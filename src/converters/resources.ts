import map from "lodash.map";

import { VirtualMachine } from "@azure/arm-compute/esm/models";
import {
  IPConfiguration,
  NetworkInterface,
  PublicIPAddress,
} from "@azure/arm-network/esm/models";

import { AzureWebLinker } from "../azure";
import {
  NETWORK_INTERFACE_ENTITY_CLASS,
  NETWORK_INTERFACE_ENTITY_TYPE,
  NetworkInterfaceEntity,
  PUBLIC_IP_ADDRESS_ENTITY_CLASS,
  PUBLIC_IP_ADDRESS_ENTITY_TYPE,
  PublicIPAddressEntity,
  VIRTUAL_MACHINE_ENTITY_CLASS,
  VIRTUAL_MACHINE_ENTITY_TYPE,
  VirtualMachineEntity,
} from "../jupiterone";
import { resourceGroup } from "./utils";
import { assignTags } from "@jupiterone/jupiter-managed-integration-sdk";

export function createNetworkInterfaceEntity(
  webLinker: AzureWebLinker,
  data: NetworkInterface,
): NetworkInterfaceEntity {
  const privateIps = privateIpAddresses(data.ipConfigurations);

  const entity: NetworkInterfaceEntity = {
    _key: data.id as string,
    _type: NETWORK_INTERFACE_ENTITY_TYPE,
    _class: NETWORK_INTERFACE_ENTITY_CLASS,
    _rawData: [{ name: "default", rawData: data }],
    resourceGuid: data.resourceGuid,
    resourceGroup: resourceGroup(data.id),
    displayName: data.name,
    virtualMachineId: data.virtualMachine && data.virtualMachine.id,
    type: data.type,
    region: data.location,
    publicIp: undefined,
    publicIpAddress: undefined,
    privateIp: privateIps,
    privateIpAddress: privateIps,
    macAddress: data.macAddress,
    securityGroupId: data.networkSecurityGroup && data.networkSecurityGroup.id,
    ipForwarding: data.enableIPForwarding,
    webLink: webLinker.portalResourceUrl(data.id),
  };

  assignTags(entity, data.tags);

  return entity;
}

export function createPublicIPAddressEntity(
  webLinker: AzureWebLinker,
  data: PublicIPAddress,
): PublicIPAddressEntity {
  const entity = {
    _key: data.id as string,
    _type: PUBLIC_IP_ADDRESS_ENTITY_TYPE,
    _class: PUBLIC_IP_ADDRESS_ENTITY_CLASS,
    _rawData: [{ name: "default", rawData: data }],
    resourceGuid: data.resourceGuid,
    resourceGroup: resourceGroup(data.id),
    displayName: data.name,
    type: data.type,
    region: data.location,
    publicIp: data.ipAddress,
    publicIpAddress: data.ipAddress,
    public: true,
    webLink: webLinker.portalResourceUrl(data.id),
    sku: data.sku && data.sku.name,
  };

  assignTags(entity, data.tags);

  return entity;
}

export function createVirtualMachineEntity(
  webLinker: AzureWebLinker,
  data: VirtualMachine,
): VirtualMachineEntity {
  const entity = {
    _key: data.id as string,
    _type: VIRTUAL_MACHINE_ENTITY_TYPE,
    _class: VIRTUAL_MACHINE_ENTITY_CLASS,
    _rawData: [{ name: "default", rawData: data }],
    displayName: data.name,
    vmId: data.vmId,
    type: data.type,
    region: data.location,
    resourceGroup: resourceGroup(data.id),
    vmSize: data.hardwareProfile && data.hardwareProfile.vmSize,
    webLink: webLinker.portalResourceUrl(data.id),
  };

  assignTags(entity, data.tags);

  return entity;
}

function privateIpAddresses(
  ipConfigurations: IPConfiguration[] | undefined,
): string[] {
  const configs =
    ipConfigurations && ipConfigurations.filter(c => c.privateIPAddress);
  return map(configs, a => a.privateIPAddress) as string[];
}
