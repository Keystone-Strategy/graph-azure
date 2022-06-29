import {
    Step,
    IntegrationStepExecutionContext,
    RelationshipClass,
    IntegrationValidationError,
    StepStartStates
  } from '@keystone-labs/integration-sdk-core';
  import { IntegrationExecutionContext, IntegrationInvocationConfig } from '@keystone-labs/integration-sdk-core';
  import { DirectoryGraphClient } from '../active-directory/client';

  import { IntegrationConfig } from '../../types';

  import {
    STEP_AD_MESSAGES,
    MESSAGE_ENTITY_TYPE,
    DOMAIN_ENTITY_TYPE, 
    EMAIL_ADDRESS_ENTITY_TYPE,
    ATTACHMENT_ENTITY_TYPE,
    CONVERSATION_ENTITY_TYPE,
  } from './constants'
  
  import fetchMessages from '../exchange/fetchMessages';


  const steps: Step<IntegrationStepExecutionContext<IntegrationConfig>>[] = [
    {
      id: STEP_AD_MESSAGES,
      name: 'Active Directory Messages',
      entities: [
        {
          resourceName: 'Message',
          _type: MESSAGE_ENTITY_TYPE,
        },
        {
          resourceName: 'Email Address',
          _type:  EMAIL_ADDRESS_ENTITY_TYPE,
        },
        {
          resourceName: 'Domain',
          _type:  DOMAIN_ENTITY_TYPE,
        },
        {
          resourceName: 'Attachment',
          _type:  ATTACHMENT_ENTITY_TYPE,
        },
        {
          resourceName: 'Conversation',
          _type:  CONVERSATION_ENTITY_TYPE,
        }
      ],
      relationships: [
        {
          sourceType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.SENT_FROM,
          targetType: EMAIL_ADDRESS_ENTITY_TYPE,
        },
        {
          sourceType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.SENT_TO,
          targetType: EMAIL_ADDRESS_ENTITY_TYPE,
        },
        {
          sourceType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.CC_TO,
          targetType: EMAIL_ADDRESS_ENTITY_TYPE,
        },
        {
          sourceType: EMAIL_ADDRESS_ENTITY_TYPE,
          _type: RelationshipClass.BELONGS_TO,
          targetType: DOMAIN_ENTITY_TYPE,
        },
        {
          sourceType: EMAIL_ADDRESS_ENTITY_TYPE,
          _type: RelationshipClass.CONTAINS,
          targetType: ATTACHMENT_ENTITY_TYPE,
        },
        {
          sourceType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.BELONGS_TO,
          targetType: CONVERSATION_ENTITY_TYPE
        }
      ],
      executionHandler: fetchMessages
    }
  ];
  
    
  export const invocationConfig : IntegrationInvocationConfig<IntegrationConfig> = {
    instanceConfigFields: {
      exchangeUserId: {
        type: 'string',
        mask: false,
      },
      exchangeStartDate: {
        type: 'string',
        mask: false,
      },
      exchangeEndDate: {
        type: 'string',
        mask: false,
      }    
    },

    validateInvocation: async (context: IntegrationExecutionContext<IntegrationConfig>): Promise<void>  => {
      const config = context.instance.config;

      if (!config.clientId || !config.clientSecret || !config.directoryId) {
        throw new IntegrationValidationError(
          'Integration configuration requires all of {clentId, clientSecret, directoryId}',
        );
      }

      if (!config.exchangeUserId || !config.exchangeStartDate || !config.exchangeEndDate) {
        throw new IntegrationValidationError(
          'Exchange integration configuration requires all of {exchangeUserId, exchangeStartDate, exchangeEndDate}',
        );
      }
    
      const directoryClient = new DirectoryGraphClient(context.logger, config);
      await directoryClient.validate();
    },

    getStepStartStates: (
      executionContext: IntegrationExecutionContext<IntegrationConfig>,
    ): StepStartStates => {

      return {
        [STEP_AD_MESSAGES]: { disabled: false },
      };
    },

    integrationSteps: [...steps]
  
  }

  