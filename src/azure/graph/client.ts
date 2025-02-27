import {
  IntegrationLogger,
  IntegrationProviderAPIError,
  IntegrationProviderAuthorizationError,
  IntegrationValidationError,
} from '@keystone-labs/integration-sdk-core';
import { FetchError } from 'node-fetch';
import "isomorphic-fetch";
import {
  AuthenticationProvider,
  Client,
  GraphRequest,
} from '@microsoft/microsoft-graph-client';
import { Organization } from '@microsoft/microsoft-graph-types';

import { IntegrationConfig } from '../../types';
import { authenticate } from './authenticate';
import { retry } from '@lifeomic/attempt';
import { permissions } from '../../steps/constants';

export type IterableGraphResponse<T> = {
  value: T[];
};

type AzureGraphResponse<TResponseType = any> = TResponseType & {
  ['@odata.nextLink']?: string;
};

/**
 * Pagination: https://docs.microsoft.com/en-us/graph/paging
 * Throttling with retry after: https://docs.microsoft.com/en-us/graph/throttling
 * Batching requests: https://docs.microsoft.com/en-us/graph/json-batching
 */
export abstract class GraphClient {
  readonly client: Client;
  readonly authenticationProvider: GraphAuthenticationProvider;

  constructor(
    readonly logger: IntegrationLogger,
    readonly config: IntegrationConfig,
  ) {
    this.authenticationProvider = new GraphAuthenticationProvider(config);
    this.client = Client.initWithMiddleware({
      authProvider: this.authenticationProvider,
    });
  }

  private async getAccessTokenPermissions(): Promise<string[]> {
    const accessToken = await this.authenticationProvider.getAccessToken();
    return getRolesFromAccessToken(accessToken);
  }

  protected async enforceApiPermission(
    endpoint: string,
    permission: string,
  ): Promise<void> {
    if (!(await this.getAccessTokenPermissions()).includes(permission)) {
      throw new IntegrationProviderAuthorizationError({
        endpoint,
        status: 'MISSING_API_PERMISSION',
        statusText: 'Missing API permission required to access this endpoint',
        resourceType: [permission],
      });
    }
  }

  public async validate(): Promise<void> {
    try {
      await this.authenticationProvider.getAccessToken();
    } catch (e) {
      throw new IntegrationValidationError(e);
    }
  }

  public async validateDirectoryPermissions(): Promise<void> {
    try {
      await this.enforceApiPermission(
        '/organization',
        permissions.graph.DIRECTORY_READ_ALL,
      );
    } catch (e) {
      throw new IntegrationValidationError(e);
    }
  }

  public async fetchMetadata(): Promise<AzureGraphResponse | undefined> {
    return this.request(this.client.api('/'));
  }

  public async fetchOrganization(): Promise<Organization> {
    const response = await this.request<IterableGraphResponse<Organization>>(
      this.client.api('/organization'),
    );
    return response!.value[0];
  }

  async retryRequest<TResponseType = any>(
    graphRequest: GraphRequest,
  ): Promise<AzureGraphResponse<TResponseType>> {
    let shouldRefreshAccessToken: boolean;
    return retry(
      async () => {
        if (shouldRefreshAccessToken) {
          /**
           * The `@lifeomic/attempt` package does not await the `handleError` method, which
           * prevents this project from calling `await this.authenticationProvider.refreshAccessToken()`
           * within handleError itself.
           *
           * Instead, to make `handleError` synchronous, we simply set `shouldRefreshAccessToken = true` and
           * await the token refresh within the `attemptFunc` itself.
           */
          await this.authenticationProvider.refreshAccessToken();
          shouldRefreshAccessToken = false;
        }
        return graphRequest.get();
      },
      {
        maxAttempts: 3,
        delay: 2000,
        handleError: (err, context, options) => {
          const endpoint = (graphRequest as any).buildFullUrl?.();
          this.logger.info(
            {
              err,
              endpoint,
              attemptsRemaining: context.attemptsRemaining,
            },
            'Encountered retryable error in Azure Graph API.',
          );

          if (
            err.code === 'Authentication_ExpiredToken' ||
            err.code === 'InvalidAuthenticationToken'
          ) {
            this.logger.info('Refreshing access token');
            shouldRefreshAccessToken = true;
          }
        },
      },
    );
  }

  public async request<TResponseType = any>(
    graphRequest: GraphRequest,
  ): Promise<AzureGraphResponse<TResponseType> | undefined> {
    try {
      const response = await this.retryRequest(graphRequest);
      return response;
    } catch (err) {
      const endpoint = (graphRequest as any).buildFullUrl?.();

      // Fetch errors include the properties code, errno, message, name, stack, type.
      if (err instanceof FetchError) {
        this.logger.warn(
          { err, resourceUrl: endpoint },
          'Encountered fetch error in Azure Graph client.',
        );
        throw new IntegrationProviderAPIError({
          cause: err,
          endpoint,
          status: err.code!,
          statusText: err.message,
        });
      }

      if (err.statusCode === 403) {
        this.logger.warn(
          { err, resourceUrl: endpoint },
          'Encountered auth error in Azure Graph client.',
        );
        throw new IntegrationProviderAuthorizationError({
          cause: err,
          endpoint,
          status: err.statusCode,
          statusText: err.statusText || err.message,
        });
      }
      if (err.statusCode !== 404) {
        this.logger.warn(
          { err, resourceUrl: endpoint },
          'Encountered error in Azure Graph client.',
        );
        throw new IntegrationProviderAPIError({
          cause: err,
          endpoint,
          status: err.statusCode,
          statusText: err.statusText || err.message,
        });
      }

      this.logger.warn(
        { err, resourceUrl: endpoint },
        'Encountered non-fatal error in Azure Graph client.',
      );
    }
  }
}

class GraphAuthenticationProvider implements AuthenticationProvider {
  private accessToken: string | undefined;

  constructor(readonly config: IntegrationConfig) {}

  /**
   * Obtains an accessToken (in case of success) or rejects with error (in case
   * of failure). Currently does not track token expiration/support token
   * refresh.
   */
  public async getAccessToken(): Promise<string> {
    if (!this.accessToken) {
      this.accessToken = await authenticate(this.config);
    }
    return this.accessToken;
  }

  public async refreshAccessToken(): Promise<void> {
    this.accessToken = await authenticate(this.config);
  }
}

function getRolesFromAccessToken(accessToken: string) {
  function parseJwtPayload(jwt: string): { roles?: string[] } | undefined {
    try {
      const encodedPayload = jwt.split('.')[1];
      const decodedPayload = Buffer.from(encodedPayload, 'base64').toString();
      return JSON.parse(decodedPayload);
    } catch (e) {
      return undefined;
    }
  }

  const payload = parseJwtPayload(accessToken);
  return payload?.roles || [];
}

export const testFunctions = {
  getRolesFromAccessToken,
};
