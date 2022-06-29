import {
  IntegrationExecutionContext,
  StepStartStates,
  StepStartState,
} from '@keystone-labs/integration-sdk-core';

import { IntegrationConfig } from './types';

import {
  STEP_AD_MESSAGES
} from './steps/exchange/constants';

function makeStepStartStates(
  stepIds: string[],
  stepStartState: StepStartState,
): StepStartStates {
  const stepStartStates: StepStartStates = {};
  for (const stepId of stepIds) {
    stepStartStates[stepId] = stepStartState;
  }
  return stepStartStates;
}

interface GetApiSteps {
  executeFirstSteps: string[];
  executeLastSteps: string[];
}

export function getActiveDirectorySteps(): GetApiSteps {
  return {
    executeFirstSteps: [
      // STEP_AD_GROUPS,
      // STEP_AD_GROUP_MEMBERS,
      // STEP_AD_USER_REGISTRATION_DETAILS,
      // STEP_AD_USERS,
      STEP_AD_MESSAGES,
      // STEP_AD_SERVICE_PRINCIPALS,
    ],
    executeLastSteps: [],
  };
}

export default function getStepStartStates(
  executionContext: IntegrationExecutionContext<IntegrationConfig>,
): StepStartStates {
  const activeDirectory = { disabled: true };

  const {
    executeFirstSteps: adFirstSteps,
    executeLastSteps: adLastSteps,
  } = getActiveDirectorySteps();
  return {
    ...makeStepStartStates([...adFirstSteps, ...adLastSteps], activeDirectory),
  };
}
