import {
  ExplicitRelationship,
  RelationshipClass,
  createDirectRelationship,
  Entity,
  JobState,
} from '@keystone-labs/integration-sdk-core';

import { IntegrationStepContext } from '../../../types';
import {
  createMessageEntity,
  createEmailAddressEntity,
  createDomainEntity,
  createAttachmentEntity,
  createConversationEntity,
} from '../../active-directory/converters';

import { DirectoryGraphClient } from '../../active-directory/client';

import * as _ from 'lodash';
import {
  ATTACHMENT_ENTITY_TYPE,
  CONVERSATION_ENTITY_TYPE,
  DOMAIN_ENTITY_TYPE,
  EMAIL_ADDRESS_ENTITY_TYPE,
  MESSAGE_ENTITY_TYPE,
} from '../constants';

export default async function fetchMessages(
  executionContext: IntegrationStepContext,
): Promise<void> {
  const { logger, instance, jobState } = executionContext;
  const graphClient = new DirectoryGraphClient(logger, instance.config);

  await graphClient.iterateUserMessages(
    {
      userId: instance.config.exchangeUserId as string,
      startDate: instance.config.exchangeStartDate as string,
      endDate: instance.config.exchangeEndDate as string,
    },
    async (message: any) => {
      const alreadyExistingMessageEntity = await jobState.findEntity(
        message.id,
      );
      if (alreadyExistingMessageEntity) return;

      const messageEntity = createMessageEntity(message);
      const {
        fromEmailAddressEntity,
        fromDomainEntity,
        messageSentFromEmailAddressRelationship,
        emailAddressBelongsToDomainRelationship,
      }: any = processSender(message, messageEntity);

      const conversationEntity = createConversationEntity(
        message.conversationId,
      );
      const messageBelongsToConversationRelationship = createDirectRelationship(
        {
          fromKey: messageEntity._key,
          fromType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.BELONGS_TO,
          toKey: conversationEntity._key,
          toType: CONVERSATION_ENTITY_TYPE,
        },
      );

      const {
        toEmailAddressEntities,
        toDomainEntities,
        messageSentToEmailAddressRelationships,
        emailAddressBelongsToDomainRelationships:
          emailAddressBelongsToDomainRelationshipsFromRecipients,
      }: any = processRecipients(message, messageEntity, 'to');

      const {
        toEmailAddressEntities: ccEmailAddressEntities,
        toDomainEntities: ccDomainEntities,
        messageSentToEmailAddressRelationships:
          messageSentCCEmailAddressRelationships,
        emailAddressBelongsToDomainRelationships:
          emailAddressBelongsToDomainRelationshipsFromCCs,
      } = processRecipients(message, messageEntity, 'cc');

      const entities: Entity[] = [];
      entities.push(messageEntity);
      entities.push(conversationEntity);
      if (fromEmailAddressEntity !== null)
        entities.push(fromEmailAddressEntity);
      if (fromDomainEntity !== null) entities.push(fromDomainEntity);
      entities.push(...toEmailAddressEntities);
      entities.push(...toDomainEntities);
      entities.push(...ccEmailAddressEntities);
      entities.push(...ccDomainEntities);

      const relationships: ExplicitRelationship[] = [];
      relationships.push(messageBelongsToConversationRelationship);
      if (messageSentFromEmailAddressRelationship !== null)
        relationships.push(messageSentFromEmailAddressRelationship);
      if (emailAddressBelongsToDomainRelationship !== null)
        relationships.push(emailAddressBelongsToDomainRelationship);
      relationships.push(...messageSentToEmailAddressRelationships);
      relationships.push(
        ...emailAddressBelongsToDomainRelationshipsFromRecipients,
      );
      relationships.push(...messageSentCCEmailAddressRelationships);
      relationships.push(...emailAddressBelongsToDomainRelationshipsFromCCs);

      await storeEntities(entities, jobState);
      await storeRelationships(relationships, jobState);

      if (!message.hasAttachments) return;

      const attachments = await graphClient.listAttachments(
        instance.config.exchangeUserId,
        message.id,
      );

      for (const attachment of attachments) {
        const attachmentEntity = createAttachmentEntity(attachment);
        const attachmentRelationship = createDirectRelationship({
          fromKey: messageEntity._key,
          fromType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.CONTAINS,
          toKey: attachmentEntity._key,
          toType: ATTACHMENT_ENTITY_TYPE,
        });
        await storeEntities([attachmentEntity], jobState);
        await storeRelationships([attachmentRelationship], jobState);
      }
    },
  );
}

/**
 * A person dict as it comes from Microsoft.
 */
interface MSEmailAddress {
  address: string;
  name: string;
}

const getDomainFromEmailAddress = (address: string): string => {
  const afterAt = address.split('@')[1];
  const splitDot = afterAt.split('.');
  const domain = `${splitDot[splitDot.length - 2]}.${
    splitDot[splitDot.length - 1]
  }`;
  return domain.toLowerCase();
};

const processSender = (message: { from: MSEmailAddress }, messageEntity) => {
  const fromEmailAddressAddress = _.get(
    message,
    'from.emailAddress.address',
    null,
  );
  if (fromEmailAddressAddress === null) {
    return {
      fromEmailAddressEntity: null,
      fromDomainEntity: null,
      messageSentFromEmailAddressRelationship: null,
      emailAddressBelongsToDomainRelationship: null,
    };
  }
  const fromEmailAddressName = _.get(
    message,
    'from.emailAddress.name',
    'NOT DEFINED',
  );
  const fromEmailAddressEntity = createEmailAddressEntity(
    fromEmailAddressAddress,
    fromEmailAddressName,
  );

  const fromDomain = getDomainFromEmailAddress(fromEmailAddressAddress);
  const fromDomainEntity =
    fromDomain !== null ? createDomainEntity(fromDomain) : null;

  const emailAddressBelongsToDomainRelationship =
    fromDomainEntity !== null
      ? createDirectRelationship({
          fromKey: fromEmailAddressEntity._key,
          fromType: EMAIL_ADDRESS_ENTITY_TYPE,
          _type: RelationshipClass.BELONGS_TO,
          toKey: fromDomainEntity._key,
          toType: DOMAIN_ENTITY_TYPE,
        })
      : null;

  // MESSAGE_SENT_FROM_EMAIL_ADDRESS_RELATIONSHIP
  const messageSentFromEmailAddressRelationship = createDirectRelationship({
    fromKey: messageEntity._key,
    fromType: MESSAGE_ENTITY_TYPE,
    _type: RelationshipClass.SENT_FROM,
    toKey: fromEmailAddressEntity._key,
    toType: EMAIL_ADDRESS_ENTITY_TYPE,
  });

  return {
    fromEmailAddressEntity,
    fromDomainEntity,
    messageSentFromEmailAddressRelationship,
    emailAddressBelongsToDomainRelationship,
  };
};

const processRecipients = (message, messageEntity, type: 'to' | 'cc') => {
  const toEmailAddressEntities: any = [];
  const toDomainEntities: any = [];
  const messageSentToEmailAddressRelationships: any = [];
  const emailAddressBelongsToDomainRelationships: any = [];

  for (const toEmailRecipient of message.toRecipients) {
    const toEmailAddressAddress = _.get(
      toEmailRecipient,
      'emailAddress.address',
      null,
    );

    if (!toEmailAddressAddress) continue;

    const toEmailAddressName = _.get(
      toEmailRecipient,
      'emailAddress.name',
      'NOT DEFINED',
    );
    const toEmailAddressEntity = createEmailAddressEntity(
      toEmailAddressAddress,
      toEmailAddressName,
    );

    const toDomain = getDomainFromEmailAddress(toEmailAddressAddress);
    const toDomainEntity =
      toDomain !== null ? createDomainEntity(toDomain) : null;

    const emailAddressBelongsToDomainRelationship =
      toDomainEntity !== null
        ? createDirectRelationship({
            fromKey: toEmailAddressEntity._key,
            fromType: EMAIL_ADDRESS_ENTITY_TYPE,
            _type: RelationshipClass.BELONGS_TO,
            toKey: toDomainEntity._key,
            toType: DOMAIN_ENTITY_TYPE,
          })
        : null;

    // MESSAGE_SENT_TO_EMAIL_ADDRESS_RELATIONSHIP
    const messageSentToEmailAddressRelationship = createDirectRelationship({
      fromKey: messageEntity._key,
      fromType: MESSAGE_ENTITY_TYPE,
      _type:
        type === 'cc' ? RelationshipClass.CC_TO : RelationshipClass.SENT_TO,
      toKey: toEmailAddressEntity._key,
      toType: EMAIL_ADDRESS_ENTITY_TYPE,
    });

    toEmailAddressEntities.push(toEmailAddressEntity);
    if (toDomainEntity !== null) toDomainEntities.push(toDomainEntity);
    messageSentToEmailAddressRelationships.push(
      messageSentToEmailAddressRelationship,
    );
    if (emailAddressBelongsToDomainRelationship !== null)
      emailAddressBelongsToDomainRelationships.push(
        emailAddressBelongsToDomainRelationship,
      );
  }

  return {
    toEmailAddressEntities,
    toDomainEntities,
    messageSentToEmailAddressRelationships,
    emailAddressBelongsToDomainRelationships,
  };
};

const storeEntities = async (entities, jobState: JobState) => {
  const uniqueEntities = _.uniqBy(entities, (e: Entity) => e._key);

  const newEntities: Entity[] = [];
  for (const entity of uniqueEntities) {
    const existingEntity = await jobState.hasKey(entity._key);
    if (!existingEntity) newEntities.push(entity);
  }

  await jobState.addEntities(newEntities);
};

const storeRelationships = async (relationships, jobState: JobState) => {
  const uniqueRelationships = _.uniqBy(
    relationships,
    (r: ExplicitRelationship) => r._key,
  );

  const newRelationships: ExplicitRelationship[] = [];
  for (const relationship of uniqueRelationships) {
    const existingRelationship = await jobState.hasKey(relationship._key);
    if (!existingRelationship) newRelationships.push(relationship);
  }

  await jobState.addRelationships(newRelationships);
};
