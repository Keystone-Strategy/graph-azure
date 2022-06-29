
import {
  createIntegrationEntity,
  Entity,
} from '@keystone-labs/integration-sdk-core';

import { generateEntityKey } from '../../utils/generateKeys';

import{
  MESSAGE_ENTITY_TYPE,
  EMAIL_ADDRESS_ENTITY_TYPE,
  DOMAIN_ENTITY_TYPE,
  ATTACHMENT_ENTITY_TYPE,
  CONVERSATION_ENTITY_TYPE
} from '../exchange/constants'

export function createMessageEntity (
  data: any
): Entity {
  const subject = data.subject || 'NO SUBJECT'
  return createIntegrationEntity({
    entityData: {
      source: {},
      assign: {
        _key: generateEntityKey(data.id),
        _type: MESSAGE_ENTITY_TYPE,
        name: subject,
        subject,
        receivedDateTime: data.receivedDateTime
      },
    },
  });
}

export function createConversationEntity (
  conversationId: string
): Entity {
  return createIntegrationEntity({
    entityData: {
      source: {},
      assign: {
        _key: generateEntityKey(conversationId),
        _type: CONVERSATION_ENTITY_TYPE,
        name: conversationId,
      },
    },
  });
}

export function createAttachmentEntity(
  data
): Entity {
  const name = data.name ? data.name : 'No attachment name'
  return createIntegrationEntity({
    entityData: {
      source: {},
      assign: {
        _key: generateEntityKey(data.id),
        _type:  ATTACHMENT_ENTITY_TYPE,
        name,
        contentType: data.contentType,
        size: data.size,
        isInLine: data.isInLine,
        lastModifiedDateTime: data.lastModifiedDateTime
      },
    },
  });
}

export function createEmailAddressEntity(
  address: string,
  name: string
): Entity {
  const data = {
    address, name
  }
  const emailAddressKey = address.padEnd(10, '_').toLowerCase()
  return createIntegrationEntity({
    entityData: {
      source: data,
      assign: {
        _key: generateEntityKey(emailAddressKey),
        _type: EMAIL_ADDRESS_ENTITY_TYPE,
        name: name,
        address: address
      },
    },
  });
}

export function createDomainEntity(
  domain: string
): Entity {
  const data = { domain }
  return createIntegrationEntity({
    entityData: {
      source: data,
      assign: {
        _key: generateEntityKey(domain),
        _type:  DOMAIN_ENTITY_TYPE,
        name: domain,
      },
    },
  });
}
