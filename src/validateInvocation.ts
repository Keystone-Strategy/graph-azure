import {
  IntegrationExecutionContext,
  IntegrationValidationError,
} from '@keystone-labs/integration-sdk-core';

import { DirectoryGraphClient } from './steps/active-directory/client';
import { IntegrationConfig } from './types';

export default async function validateInvocation(
  context: IntegrationExecutionContext<IntegrationConfig>,
): Promise<void> {
  const config = context.instance.config;

  if (!config.clientId || !config.clientSecret || !config.directoryId) {
    throw new IntegrationValidationError(
      'Integration configuration requires all of {clentId, clientSecret, directoryId}',
    );
  }

  const directoryClient = new DirectoryGraphClient(context.logger, config);
  await directoryClient.validate();

  if (config.ingestActiveDirectory) {
    await directoryClient.validateDirectoryPermissions();
  }
}
