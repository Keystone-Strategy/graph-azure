import {
  IntegrationInstanceAuthenticationError,
  IntegrationInstanceConfigError,
  IntegrationValidationContext,
} from "@jupiterone/jupiter-managed-integration-sdk";

import { AzureClient } from "./azure";
import { AzureIntegrationInstanceConfig } from "./types";

/**
 * Performs validation of the execution before the execution handler function is
 * invoked.
 *
 * At a minimum, integrations should ensure that the
 * `context.instance.config` is valid. Integrations that require
 * additional information in `context.invocationArgs` should also
 * validate those properties. It is also helpful to perform authentication with
 * the provider to ensure that credentials are valid.
 *
 * The function will be awaited to support connecting to the provider for this
 * purpose.
 *
 * @param context
 */
export default async function invocationValidator(
  validationContext: IntegrationValidationContext,
) {
  const config = validationContext.instance
    .config as AzureIntegrationInstanceConfig;

  if (!config.clientId || !config.clientSecret || !config.directoryId) {
    throw new IntegrationInstanceConfigError(
      "clentId or clientSecret or directoryId missing in config",
    );
  }
  try {
    const client = new AzureClient(
      config.clientId,
      config.clientSecret,
      config.directoryId,
      validationContext.logger,
    );
    await client.authenticate();
  } catch (err) {
    throw new IntegrationInstanceAuthenticationError(err);
  }
}
