/**
 * Microsoft Intune device management tools, read-only.
 *
 * Every DeviceManagementManagedDevices.* scope needs tenant admin consent even
 * in delegated mode, so the group stays disabled unless `--org-mode` is set.
 * The tenant must also be licensed for Intune: an unlicensed or non-Intune
 * tenant answers these paths with 404 or 403 rather than an empty collection.
 */

import { z } from 'zod';
import type { ToolDefinition, ToolDeps, ToolGroupMeta, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { truncateText } from '../util/truncate.js';

// The group catalogue is a static literal; `intune` is always present.
const GROUP: ToolGroupMeta = GROUPS['intune']!;
const READ_SCOPES: string[] = [...GROUP.readScopes];

const DEVICES_PATH = '/deviceManagement/managedDevices';

/** Projected fields when the caller passes no `select`. */
const DEFAULT_DEVICE_FIELDS: string[] = [
  'id',
  'deviceName',
  'operatingSystem',
  'osVersion',
  'complianceState',
  'lastSyncDateTime',
  'userPrincipalName',
  'manufacturer',
  'model',
];

/** Non-compliant devices are usually being triaged, so carry the grace period too. */
const NONCOMPLIANT_FIELDS: string[] = [
  ...DEFAULT_DEVICE_FIELDS,
  'complianceGracePeriodExpirationDateTime',
];

const NONCOMPLIANT_FILTER = "complianceState eq 'noncompliant'";

/** Bare OData property names only — no paths, no functions, no literals. */
const FIELD_NAME = /^[A-Za-z][A-Za-z0-9_]*$/;

// ---------------------------------------------------------------------------
// Graph payload shapes (only the fields these tools project)
// ---------------------------------------------------------------------------

interface ManagedDevice {
  id?: string;
  deviceName?: string;
  managedDeviceName?: string;
  userPrincipalName?: string;
  userDisplayName?: string;
  emailAddress?: string;
  operatingSystem?: string;
  osVersion?: string;
  complianceState?: string;
  complianceGracePeriodExpirationDateTime?: string;
  enrolledDateTime?: string;
  lastSyncDateTime?: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  imei?: string;
  meid?: string;
  udid?: string;
  phoneNumber?: string;
  subscriberCarrier?: string;
  wiFiMacAddress?: string;
  ethernetMacAddress?: string;
  managementAgent?: string;
  managedDeviceOwnerType?: string;
  deviceEnrollmentType?: string;
  deviceRegistrationState?: string;
  deviceCategoryDisplayName?: string;
  azureADDeviceId?: string;
  azureADRegistered?: boolean | null;
  isEncrypted?: boolean;
  isSupervised?: boolean;
  jailBroken?: string;
  androidSecurityPatchLevel?: string;
  exchangeAccessState?: string;
  exchangeAccessStateReason?: string;
  partnerReportedThreatState?: string;
  totalStorageSpaceInBytes?: number;
  freeStorageSpaceInBytes?: number;
  physicalMemoryInBytes?: number;
  notes?: string | null;
}

interface DetectedApp {
  id?: string;
  displayName?: string;
  version?: string;
  publisher?: string;
  platform?: string;
  sizeInByte?: number;
  deviceCount?: number;
}

interface CompliancePolicy {
  '@odata.type'?: string;
  id?: string;
  displayName?: string;
  description?: string | null;
  version?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

interface DeviceOperatingSystemSummary {
  androidCount?: number;
  iosCount?: number;
  macOSCount?: number;
  windowsCount?: number;
  windowsMobileCount?: number;
  unknownCount?: number;
}

interface DeviceExchangeAccessStateSummary {
  allowedDeviceCount?: number;
  blockedDeviceCount?: number;
  quarantinedDeviceCount?: number;
  unknownDeviceCount?: number;
  unavailableDeviceCount?: number;
}

interface ManagedDeviceOverview {
  id?: string;
  enrolledDeviceCount?: number;
  mdmEnrolledCount?: number;
  dualEnrolledDeviceCount?: number;
  lastModifiedDateTime?: string;
  deviceOperatingSystemSummary?: DeviceOperatingSystemSummary | null;
  deviceExchangeAccessStateSummary?: DeviceExchangeAccessStateSummary | null;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Intune returns `""` for most unset string properties rather than omitting
 * them, so empty strings are dropped alongside null to keep projections small.
 */
function compact<T extends Record<string, unknown>>(value: T): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) {
    if (v !== undefined && v !== null && v !== '') out[key] = v;
  }
  return out;
}

function pick(source: Record<string, unknown>, fields: readonly string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const field of fields) {
    const value = source[field];
    if (value !== undefined && value !== null && value !== '') out[field] = value;
  }
  return out;
}

function collectionResult(
  items: Record<string, unknown>[],
  key: string,
  nextLink: string | undefined,
): Record<string, unknown> {
  return compact({ count: items.length, [key]: items, nextLink });
}

/**
 * Splits a caller-supplied `$select` into validated property names. Rejecting
 * anything but a bare property name keeps `$select` from becoming a second
 * place a model can smuggle OData syntax into the query string.
 */
function parseSelectFields(select: string): string[] {
  const fields = select
    .split(',')
    .map((part) => part.trim())
    .filter((part) => part.length > 0);

  if (fields.length === 0) {
    throw new Error('select must name at least one managedDevice property, e.g. "deviceName,complianceState".');
  }

  for (const field of fields) {
    if (!FIELD_NAME.test(field)) {
      throw new Error(
        `Invalid select field ${JSON.stringify(field)}: pass bare managedDevice property names such as deviceName or osVersion.`,
      );
    }
  }

  // The device id is the handle every other intune tool takes, so a custom
  // select must not be able to strip it out.
  if (!fields.includes('id')) fields.unshift('id');
  return fields;
}

/** Byte counts cost tokens and read badly; storage is reasoned about in GB. */
function toGb(bytes: number | undefined | null): number | undefined {
  if (typeof bytes !== 'number' || !Number.isFinite(bytes) || bytes <= 0) return undefined;
  return Math.round((bytes / 1024 ** 3) * 10) / 10;
}

function toMb(bytes: number | undefined | null): number | undefined {
  if (typeof bytes !== 'number' || !Number.isFinite(bytes) || bytes <= 0) return undefined;
  return Math.round((bytes / (1024 * 1024)) * 10) / 10;
}

/**
 * `deviceCompliancePolicy` is an abstract type: the platform is only visible in
 * the `@odata.type` discriminator, so surface its last segment.
 */
function shortODataType(value: string | undefined): string | undefined {
  if (typeof value !== 'string' || value === '') return undefined;
  const last = value.split('.').pop();
  return last === undefined || last === '' ? undefined : last;
}

function projectDeviceDetail(device: ManagedDevice): Record<string, unknown> {
  return compact({
    id: device.id,
    deviceName: device.deviceName,
    managedDeviceName: device.managedDeviceName,
    userPrincipalName: device.userPrincipalName,
    userDisplayName: device.userDisplayName,
    emailAddress: device.emailAddress,
    operatingSystem: device.operatingSystem,
    osVersion: device.osVersion,
    androidSecurityPatchLevel: device.androidSecurityPatchLevel,
    complianceState: device.complianceState,
    complianceGracePeriodExpirationDateTime: device.complianceGracePeriodExpirationDateTime,
    enrolledDateTime: device.enrolledDateTime,
    lastSyncDateTime: device.lastSyncDateTime,
    manufacturer: device.manufacturer,
    model: device.model,
    serialNumber: device.serialNumber,
    imei: device.imei,
    meid: device.meid,
    udid: device.udid,
    phoneNumber: device.phoneNumber,
    subscriberCarrier: device.subscriberCarrier,
    wiFiMacAddress: device.wiFiMacAddress,
    ethernetMacAddress: device.ethernetMacAddress,
    managementAgent: device.managementAgent,
    ownerType: device.managedDeviceOwnerType,
    deviceEnrollmentType: device.deviceEnrollmentType,
    deviceRegistrationState: device.deviceRegistrationState,
    deviceCategory: device.deviceCategoryDisplayName,
    azureADDeviceId: device.azureADDeviceId,
    azureADRegistered: device.azureADRegistered,
    isEncrypted: device.isEncrypted,
    isSupervised: device.isSupervised,
    jailBroken: device.jailBroken === 'Unknown' ? undefined : device.jailBroken,
    exchangeAccessState: device.exchangeAccessState,
    exchangeAccessStateReason: device.exchangeAccessStateReason,
    partnerReportedThreatState: device.partnerReportedThreatState,
    totalStorageGb: toGb(device.totalStorageSpaceInBytes),
    freeStorageGb: toGb(device.freeStorageSpaceInBytes),
    physicalMemoryGb: toGb(device.physicalMemoryInBytes),
    notes: typeof device.notes === 'string' ? truncateText(device.notes, 500) : undefined,
  });
}

function projectDetectedApp(app: DetectedApp): Record<string, unknown> {
  return compact({
    id: app.id,
    displayName: app.displayName,
    version: app.version,
    publisher: app.publisher,
    platform: app.platform,
    sizeMb: toMb(app.sizeInByte),
    deviceCount: app.deviceCount,
  });
}

function projectCompliancePolicy(policy: CompliancePolicy): Record<string, unknown> {
  const description = policy.description;
  return compact({
    id: policy.id,
    displayName: policy.displayName,
    policyType: shortODataType(policy['@odata.type']),
    description: typeof description === 'string' ? truncateText(description, 300) : undefined,
    version: policy.version,
    createdDateTime: policy.createdDateTime,
    lastModifiedDateTime: policy.lastModifiedDateTime,
  });
}

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const listManagedDevicesSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(25)
    .describe('Maximum number of devices to return in this page. Defaults to 25.'),
  filter: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Raw OData $filter over managedDevice properties, e.g. \"operatingSystem eq 'iOS'\" or \"complianceState eq 'noncompliant'\". Intune only filters on a small set of properties (deviceName, userId, userPrincipalName, complianceState, operatingSystem, osVersion, managementAgent, managedDeviceOwnerType, serialNumber, azureADDeviceId among them) and rejects startswith/contains, so stick to equality.",
    ),
  select: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Comma-separated managedDevice properties to return instead of the default set (id, deviceName, operatingSystem, osVersion, complianceState, lastSyncDateTime, userPrincipalName, manufacturer, model). Bare property names only; id is always included.',
    ),
});

const getManagedDeviceSchema = z.object({
  deviceId: z
    .string()
    .min(1)
    .describe('Intune managedDevice id (a GUID), as returned by intune_list_managed_devices. Not the Entra device id.'),
});

const listDeviceAppsSchema = z.object({
  deviceId: z
    .string()
    .min(1)
    .describe('Intune managedDevice id whose detected app inventory to read, from intune_list_managed_devices.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(50)
    .describe('Maximum number of detected apps to return in this page. Defaults to 50.'),
});

const listCompliancePoliciesSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of compliance policies to return in this page. Defaults to 50.'),
});

const getDeviceSummarySchema = z.object({});

const listNoncompliantSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(25)
    .describe('Maximum number of non-compliant devices to return in this page. Defaults to 25.'),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const intuneModule: ToolModule = {
  group: GROUP,

  build({ graph }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'intune_list_managed_devices',
        title: 'List managed devices',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Lists Intune-managed devices in the tenant with id, name, operating system and version, compliance state, last sync time, primary user, manufacturer and model, plus count and nextLink when more pages exist. Page size defaults to 25. Requires tenant admin consent (DeviceManagementManagedDevices.Read.All); complianceState and lastSyncDateTime reflect the last device check-in, typically within the last eight hours, so a device offline for days reports stale values.',
        inputSchema: listManagedDevicesSchema,
        handler: async (args) => {
          const { top, filter, select } = listManagedDevicesSchema.parse(args);
          const fields = select === undefined ? DEFAULT_DEVICE_FIELDS : parseSelectFields(select);

          const res = await graph.request({
            path: DEVICES_PATH,
            query: { $top: top, $filter: filter, $select: fields.join(',') },
            scopes: READ_SCOPES,
          });

          const devices = extractCollection<Record<string, unknown>>(res.data).map((device) =>
            pick(device, fields),
          );
          return collectionResult(devices, 'devices', res.nextLink);
        },
      },

      {
        name: 'intune_get_managed_device',
        title: 'Get a managed device',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Reads one Intune-managed device in detail: hardware identity, primary user, enrollment and management state, encryption and supervision flags, Exchange access state, storage and memory in GB, and the last sync time. Takes the Intune managedDevice id from intune_list_managed_devices, not the Entra device id or the device name. Requires tenant admin consent.',
        inputSchema: getManagedDeviceSchema,
        handler: async (args) => {
          const { deviceId } = getManagedDeviceSchema.parse(args);
          const res = await graph.request<ManagedDevice>({
            path: `${DEVICES_PATH}/${encodeURIComponent(deviceId)}`,
            scopes: READ_SCOPES,
          });
          return projectDeviceDetail(res.data);
        },
      },

      {
        name: 'intune_list_device_apps',
        title: 'List apps detected on a device',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Lists the applications Intune has detected on one managed device, returning display name, version, size in MB and the tenant-wide deviceCount for each, with count and nextLink. Page size defaults to 50. Requires tenant admin consent. The app inventory is collected on a slow cycle (roughly weekly for a full scan) and covers only managed platforms, so a freshly installed app can be missing and an uninstalled one can linger.',
        inputSchema: listDeviceAppsSchema,
        handler: async (args) => {
          const { deviceId, top } = listDeviceAppsSchema.parse(args);
          const res = await graph.request({
            // detectedApps ignores $select and $filter, so the projection happens here.
            path: `${DEVICES_PATH}/${encodeURIComponent(deviceId)}/detectedApps`,
            query: { $top: top },
            scopes: READ_SCOPES,
          });

          const apps = extractCollection<DetectedApp>(res.data).map(projectDetectedApp);
          return compact({ deviceId, count: apps.length, apps, nextLink: res.nextLink });
        },
      },

      {
        name: 'intune_list_compliance_policies',
        title: 'List device compliance policies',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Lists the tenant's Intune device compliance policies with id, display name, platform-specific policyType (from the @odata.type discriminator, e.g. windows10CompliancePolicy), description, version and modification times, plus count and nextLink. Page size defaults to 50. Requires tenant admin consent. Only the shared base fields come back — the actual rule settings live on the derived type and assignments are not included.",
        inputSchema: listCompliancePoliciesSchema,
        handler: async (args) => {
          const { top } = listCompliancePoliciesSchema.parse(args);
          const res = await graph.request({
            path: '/deviceManagement/deviceCompliancePolicies',
            query: { $top: top },
            scopes: READ_SCOPES,
          });

          const policies = extractCollection<CompliancePolicy>(res.data).map(projectCompliancePolicy);
          return collectionResult(policies, 'policies', res.nextLink);
        },
      },

      {
        name: 'intune_get_device_summary',
        title: 'Get tenant device summary',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Returns the tenant-wide Intune device counts in one call: total enrolled devices, MDM-enrolled and dual-enrolled counts, a per-operating-system breakdown, and the Exchange access state breakdown. Requires tenant admin consent. This is a singleton with no paging or filtering, and the service recomputes it periodically, so the totals can trail intune_list_managed_devices by a few hours.',
        inputSchema: getDeviceSummarySchema,
        handler: async () => {
          const res = await graph.request<ManagedDeviceOverview>({
            path: '/deviceManagement/managedDeviceOverview',
            scopes: READ_SCOPES,
          });
          const overview = res.data;
          const os = overview.deviceOperatingSystemSummary;
          const exchange = overview.deviceExchangeAccessStateSummary;

          return compact({
            enrolledDeviceCount: overview.enrolledDeviceCount,
            mdmEnrolledCount: overview.mdmEnrolledCount,
            dualEnrolledDeviceCount: overview.dualEnrolledDeviceCount,
            lastModifiedDateTime: overview.lastModifiedDateTime,
            byOperatingSystem:
              os === null || os === undefined
                ? undefined
                : compact({
                    android: os.androidCount,
                    ios: os.iosCount,
                    macOS: os.macOSCount,
                    windows: os.windowsCount,
                    windowsMobile: os.windowsMobileCount,
                    unknown: os.unknownCount,
                  }),
            exchangeAccess:
              exchange === null || exchange === undefined
                ? undefined
                : compact({
                    allowed: exchange.allowedDeviceCount,
                    blocked: exchange.blockedDeviceCount,
                    quarantined: exchange.quarantinedDeviceCount,
                    unavailable: exchange.unavailableDeviceCount,
                    unknown: exchange.unknownDeviceCount,
                  }),
          });
        },
      },

      {
        name: 'intune_list_noncompliant',
        title: 'List non-compliant devices',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Lists managed devices whose compliance state is exactly 'noncompliant', with the usual device fields plus the compliance grace period expiry, count and nextLink. Page size defaults to 25. Requires tenant admin consent. Devices in the other failing states — 'error', 'conflict', 'inGracePeriod' and 'unknown' (never checked in) — are deliberately excluded; use intune_list_managed_devices with a filter to see those.",
        inputSchema: listNoncompliantSchema,
        handler: async (args) => {
          const { top } = listNoncompliantSchema.parse(args);
          const res = await graph.request({
            path: DEVICES_PATH,
            query: {
              $top: top,
              $filter: NONCOMPLIANT_FILTER,
              $select: NONCOMPLIANT_FIELDS.join(','),
            },
            scopes: READ_SCOPES,
          });

          const devices = extractCollection<Record<string, unknown>>(res.data).map((device) =>
            pick(device, NONCOMPLIANT_FIELDS),
          );
          return collectionResult(devices, 'devices', res.nextLink);
        },
      },
    ];
  },
};
