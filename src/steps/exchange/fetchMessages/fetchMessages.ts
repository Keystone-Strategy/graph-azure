import {
  RelationshipClass,
  createDirectRelationship,
  Entity,
  JobState,
  Relationship,
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
      }: any = processRecipients(message, messageEntity);

      const {
        ccEmailAddressEntities,
        ccDomainEntities,
        messageSentCCEmailAddressRelationships,
        emailAddressBelongsToDomainRelationships:
          emailAddressBelongsToDomainRelationshipsFromCCs,
      } = processCCs(message, messageEntity);

      const entities: any = [];
      entities.push(messageEntity);
      entities.push(conversationEntity);
      if (fromEmailAddressEntity !== null)
        entities.push(fromEmailAddressEntity);
      if (fromDomainEntity !== null) entities.push(fromDomainEntity);
      entities.push(...toEmailAddressEntities);
      entities.push(...toDomainEntities);
      entities.push(...ccEmailAddressEntities);
      entities.push(...ccDomainEntities);

      const relationships: any = [];
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
        const attachmentReationship = createDirectRelationship({
          fromKey: messageEntity._key,
          fromType: MESSAGE_ENTITY_TYPE,
          _type: RelationshipClass.CONTAINS,
          toKey: attachmentEntity._key,
          toType: ATTACHMENT_ENTITY_TYPE,
        });
        await storeEntities([attachmentEntity], jobState);
        await storeRelationships([attachmentReationship], jobState);
      }
    },
  );
}

const getDomainFromEmailAddress = (address) => {
  const afterAt = address.split('@')[1];
  if (afterAt === undefined) return null;
  const splitDot = afterAt.split('.');
  const domain = `${splitDot[splitDot.length - 2]}.${
    splitDot[splitDot.length - 1]
  }`;
  return domain.toLowerCase();
};

const processSender = (message, messageEntity) => {
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

const processRecipients = (message, messageEntity) => {
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
      _type: RelationshipClass.SENT_TO,
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

const processCCs = (message, messageEntity) => {
  const ccEmailAddressEntities: any = [];
  const ccDomainEntities: any = [];
  const messageSentCCEmailAddressRelationships: any = [];
  const emailAddressBelongsToDomainRelationships: any = [];

  for (const ccEmailRecipient of message.ccRecipients) {
    const ccEmailAddressAddress = _.get(
      ccEmailRecipient,
      'emailAddress.address',
      null,
    );

    if (!ccEmailAddressAddress) continue;

    const ccEmailAddressName = _.get(
      ccEmailRecipient,
      'emailAddress.name',
      'NOT DEFINED',
    );
    const ccEmailAddressEntity = createEmailAddressEntity(
      ccEmailAddressAddress,
      ccEmailAddressName,
    );

    const ccDomain = getDomainFromEmailAddress(ccEmailAddressAddress);
    const ccDomainEntity =
      ccDomain !== null ? createDomainEntity(ccDomain) : null;

    const emailAddressBelongsToDomainRelationship =
      ccDomainEntity !== null
        ? createDirectRelationship({
            fromKey: ccEmailAddressEntity._key,
            fromType: EMAIL_ADDRESS_ENTITY_TYPE,
            _type: RelationshipClass.BELONGS_TO,
            toKey: ccDomainEntity._key,
            toType: DOMAIN_ENTITY_TYPE,
          })
        : null;

    // MESSAGE_SENT_TO_EMAIL_ADDRESS_RELATIONSHIP
    const messageSentCCEmailAddressRelationship = createDirectRelationship({
      fromKey: messageEntity._key,
      fromType: MESSAGE_ENTITY_TYPE,
      _type: RelationshipClass.CC_TO,
      toKey: ccEmailAddressEntity._key,
      toType: EMAIL_ADDRESS_ENTITY_TYPE,
    });

    ccEmailAddressEntities.push(ccEmailAddressEntity);
    if (ccDomainEntity !== null) ccDomainEntities.push(ccDomainEntity);
    messageSentCCEmailAddressRelationships.push(
      messageSentCCEmailAddressRelationship,
    );
    if (emailAddressBelongsToDomainRelationship !== null)
      emailAddressBelongsToDomainRelationships.push(
        emailAddressBelongsToDomainRelationship,
      );
  }

  return {
    ccEmailAddressEntities,
    ccDomainEntities,
    messageSentCCEmailAddressRelationships,
    emailAddressBelongsToDomainRelationships,
  };
};

const storeEntities = async (entities, jobState: JobState) => {
  const uniqueEntities = _.uniqBy(entities, (e: Entity) => e._key);

  const newEntities: any = [];
  for (const entity of uniqueEntities) {
    const existingEntity = await jobState.hasKey(entity._key);
    if (!existingEntity) newEntities.push(entity);
  }

  await jobState.addEntities(newEntities);
};

const storeRelationships = async (relationships, jobState: JobState) => {
  const uniqueRelationships = _.uniqBy(
    relationships,
    (r: Relationship) => r._key,
  );

  const newRelationships: any = [];
  for (const relationship of uniqueRelationships) {
    const existingRelationship = await jobState.hasKey(relationship._key);
    if (!existingRelationship) newRelationships.push(relationship);
  }

  await jobState.addRelationships(newRelationships);
};
