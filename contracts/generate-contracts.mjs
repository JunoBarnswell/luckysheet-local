import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const contractsRoot = path.dirname(fileURLToPath(import.meta.url));
const repositoryRoot = path.dirname(contractsRoot);
const source = JSON.parse(await readFile(path.join(contractsRoot, 'workbook-contract.json'), 'utf8'));
const errors = source.errors.map((value) => JSON.stringify(value)).join(', ');
const entries = Object.entries(source.mutations);
const permissionSource = source.permissions;
if (!permissionSource || !permissionSource.allowFields || !permissionSource.commands || !permissionSource.commandPrefixes || !permissionSource.mutations) {
  throw new Error('workbook-contract.json must declare command and mutation permission policies');
}
const allowFieldEntries = Object.entries(permissionSource.allowFields);
const protectionActionValues = ['none', ...allowFieldEntries.map(([action]) => action)];
const protectionActionUnion = protectionActionValues.map((value) => JSON.stringify(value)).join(' | ');
const coreProtectionActionUnion = allowFieldEntries.map(([action]) => JSON.stringify(action)).join(' | ');
for (const [id] of entries) {
  if (!permissionSource.mutations[id]) throw new Error(`Mutation ${id} is missing a canonical permission policy`);
}
const permissionEntries = Object.entries(permissionSource.mutations);
const commandEntries = Object.entries(permissionSource.commands);
const commandPrefixEntries = permissionSource.commandPrefixes;
function normalizePermission(value) {
  const objectScope = value.objectScope ?? (value.protectionAction === 'edit-objects'
    ? 'drawing'
    : value.protectionAction === 'none' ? 'workbook' : 'range');
  return {
    ...value,
    affectedRangeMode: value.affectedRangeMode ?? (value.checksProtection ? 'declared' : 'none'),
    objectScope,
  };
}
const permissionPolicy = (value) => {
  const policy = normalizePermission(value);
  return `new PermissionPolicy(${JSON.stringify(policy.capability)}, ${JSON.stringify(policy.protectionAction)}, ${Boolean(policy.checksProtection)}, ${JSON.stringify(policy.affectedRangeMode)}, ${JSON.stringify(policy.objectScope)})`;
};
const permissionPolicyJson = (value) => {
  const policy = normalizePermission(value);
  return `{ capability: ${JSON.stringify(policy.capability)}, protectionAction: ${JSON.stringify(policy.protectionAction)}, checksProtection: ${Boolean(policy.checksProtection)}, affectedRangeMode: ${JSON.stringify(policy.affectedRangeMode)}, objectScope: ${JSON.stringify(policy.objectScope)} }`;
};
const javaAllowFields = allowFieldEntries.map(([action, field]) =>
  `        Map.entry(${JSON.stringify(action)}, ${JSON.stringify(field)})`,
).join(',\n');
const coreAllowFields = allowFieldEntries.map(([action, field]) =>
  `  ${JSON.stringify(action)}: ${JSON.stringify(field)},`,
).join('\n');
const javaMutations = entries.map(([id, capability]) =>
  `        Map.entry(${JSON.stringify(id)}, new MutationCapability(${JSON.stringify(capability.durability)}, ${Boolean(capability.remote)}, ${JSON.stringify(capability.schema)}, ${JSON.stringify(capability.minRole)}, ${JSON.stringify(capability.rebasePolicy)}, ${Boolean(capability.javaReducer)}, ${JSON.stringify(normalizePermission(permissionSource.mutations[id]).protectionAction)}, ${Boolean(normalizePermission(permissionSource.mutations[id]).checksProtection)}, ${JSON.stringify(normalizePermission(permissionSource.mutations[id]).affectedRangeMode)}, ${JSON.stringify(normalizePermission(permissionSource.mutations[id]).objectScope)}))`,
).join(',\n');
const javaPermissionPolicies = permissionEntries.map(([id, policy]) =>
  `        Map.entry(${JSON.stringify(id)}, ${permissionPolicy(policy)})`,
).join(',\n');
const tsMutations = entries.map(([id, capability]) =>
  `  ${JSON.stringify(id)}: { durability: ${JSON.stringify(capability.durability)}, remote: ${Boolean(capability.remote)}, schema: ${JSON.stringify(capability.schema)}, minRole: ${JSON.stringify(capability.minRole)}, rebasePolicy: ${JSON.stringify(capability.rebasePolicy)}, javaReducer: ${Boolean(capability.javaReducer)}, protectionAction: ${JSON.stringify(normalizePermission(permissionSource.mutations[id]).protectionAction)}, checksProtection: ${Boolean(normalizePermission(permissionSource.mutations[id]).checksProtection)}, affectedRangeMode: ${JSON.stringify(normalizePermission(permissionSource.mutations[id]).affectedRangeMode)}, objectScope: ${JSON.stringify(normalizePermission(permissionSource.mutations[id]).objectScope)}${capability.replacement ? `, replacement: ${JSON.stringify(capability.replacement)}` : ''}${capability.collaborationKind ? `, collaborationKind: ${JSON.stringify(capability.collaborationKind)}` : ''} },`,
).join('\n');
const tsCommandPermissions = commandEntries.map(([id, policy]) =>
  `  ${JSON.stringify(id)}: ${permissionPolicyJson(policy)},`,
).join('\n');
const tsCommandPrefixes = commandPrefixEntries.map((policy) =>
  `  { prefix: ${JSON.stringify(policy.prefix)}, ...${permissionPolicyJson(policy)} },`,
).join('\n');
const tsMutationPermissions = permissionEntries.map(([id, policy]) =>
  `  ${JSON.stringify(id)}: ${permissionPolicyJson(policy)},`,
).join('\n');

const java = `// Generated by contracts/generate-contracts.mjs. Do not edit manually.
package com.xc.luckysheet.server.contract;

import java.util.Map;
import java.util.Set;

public final class GeneratedWorkbookContract {
    public static final String API_VERSION = ${JSON.stringify(source.apiVersion)};
    public static final String SNAPSHOT_SCHEMA = ${JSON.stringify(source.workbook.snapshotSchema)};
    public static final int SNAPSHOT_VERSION = ${source.workbook.snapshotVersion};
    public static final int MAX_WORKBOOK_NAME_LENGTH = ${source.workbook.maxNameLength};
    public static final int MAX_DRAWING_SOURCE_CELLS = ${source.workbook.maxDrawingSourceCells};
    public static final Set<String> ERROR_CODES = Set.of(${errors});
    public static final Map<String, MutationCapability> MUTATIONS = Map.ofEntries(
${javaMutations}
    );
    public static final Map<String, PermissionPolicy> MUTATION_PERMISSIONS = Map.ofEntries(
${javaPermissionPolicies}
    );
    public static final Map<String, String> PROTECTION_ALLOW_FIELDS = Map.ofEntries(
${javaAllowFields}
    );

    private GeneratedWorkbookContract() {}

    public static MutationCapability mutation(String id) {
        return MUTATIONS.get(id);
    }

    public static PermissionPolicy mutationPermission(String id) {
        return MUTATION_PERMISSIONS.get(id);
    }

    public static String protectionAllowField(String action) {
        return PROTECTION_ALLOW_FIELDS.get(action);
    }

    public record MutationCapability(String durability, boolean remote, String schema, String minRole, String rebasePolicy, boolean javaReducer, String protectionAction, boolean checksProtection, String affectedRangeMode, String objectScope) {}
    public record PermissionPolicy(String capability, String protectionAction, boolean checksProtection, String affectedRangeMode, String objectScope) {}
}
`;

const typescript = `// Generated by contracts/generate-contracts.mjs. Do not edit manually.
export const WORKBOOK_CONTRACT_API_VERSION = ${JSON.stringify(source.apiVersion)} as const;
export const WORKBOOK_SNAPSHOT_SCHEMA = ${JSON.stringify(source.workbook.snapshotSchema)} as const;
export const WORKBOOK_SNAPSHOT_VERSION = ${source.workbook.snapshotVersion} as const;
export const MAX_WORKBOOK_NAME_LENGTH = ${source.workbook.maxNameLength} as const;
export const CONTRACT_ERROR_CODES = [${errors}] as const;
export type ContractErrorCode = typeof CONTRACT_ERROR_CODES[number];
export type MutationDurability = 'transient' | 'local' | 'remote';
export type PermissionCapability = 'navigate' | 'edit-cell' | 'format' | 'structure' | 'drawing' | 'protect' | 'share' | 'comment' | 'restore' | 'query' | 'script';
export type ProtectionAction = ${protectionActionUnion};
export interface PermissionPolicy {
  capability: PermissionCapability;
  protectionAction: ProtectionAction;
  checksProtection: boolean;
  affectedRangeMode: 'none' | 'declared' | 'exact';
  objectScope: 'cell' | 'range' | 'row' | 'column' | 'drawing' | 'worksheet' | 'workbook';
}
export interface MutationCapability {
  durability: MutationDurability;
  remote: boolean;
  schema: string;
  minRole: 'owner' | 'editor' | 'commenter' | 'viewer';
  rebasePolicy: 'none' | 'range' | 'exact';
  javaReducer: boolean;
  protectionAction: ProtectionAction;
  checksProtection: boolean;
  affectedRangeMode: 'none' | 'declared' | 'exact';
  objectScope: 'cell' | 'range' | 'row' | 'column' | 'drawing' | 'worksheet' | 'workbook';
  replacement?: string;
  collaborationKind?: 'cell-value' | 'cell-style' | 'clear' | 'insert-rows' | 'delete-rows' | 'insert-columns' | 'delete-columns' | 'move-range' | 'sort' | 'merge' | 'table-resize' | 'drawing' | 'comment' | 'pivot-config' | 'unknown';
}
export const MUTATION_CAPABILITIES = {
${tsMutations}
} as const satisfies Record<string, MutationCapability>;
export const COMMAND_PERMISSION_POLICIES = {
${tsCommandPermissions}
} as const satisfies Record<string, PermissionPolicy>;
export const COMMAND_PERMISSION_PREFIXES = [
${tsCommandPrefixes}
] as const satisfies readonly (PermissionPolicy & { prefix: string })[];
export const MUTATION_PERMISSION_POLICIES = {
${tsMutationPermissions}
} as const satisfies Record<string, PermissionPolicy>;
export function mutationCapability(id: string): MutationCapability | undefined {
  return MUTATION_CAPABILITIES[id as keyof typeof MUTATION_CAPABILITIES];
}
export function mutationPermission(id: string): PermissionPolicy | undefined {
  return MUTATION_PERMISSION_POLICIES[id as keyof typeof MUTATION_PERMISSION_POLICIES];
}
export function commandPermission(id: string): PermissionPolicy | undefined {
  const exact = COMMAND_PERMISSION_POLICIES[id as keyof typeof COMMAND_PERMISSION_POLICIES];
  if (exact) return exact;
  return COMMAND_PERMISSION_PREFIXES.find(({ prefix }) => id.startsWith(prefix));
}
`;

const coreLimits = `// Generated by contracts/generate-contracts.mjs. Do not edit manually.
export const MAX_DRAWING_SOURCE_CELLS = ${source.workbook.maxDrawingSourceCells} as const;
`;
const coreProtection = `// Generated by contracts/generate-contracts.mjs. Do not edit manually.
export type ProtectionAction = ${coreProtectionActionUnion};
export const PROTECTION_ACTION_ALLOW_FIELD = {
${coreAllowFields}
} as const;
`;

await Promise.all([
  writeFile(path.join(repositoryRoot, 'backend/src/main/java/com/xc/luckysheet/server/contract/GeneratedWorkbookContract.java'), java, 'utf8'),
  writeFile(path.join(repositoryRoot, 'frontend-react/packages/protocol/src/generated-contract.ts'), typescript, 'utf8'),
  writeFile(path.join(repositoryRoot, 'frontend-react/packages/core-model/src/generated-workbook-limits.ts'), coreLimits, 'utf8'),
  writeFile(path.join(repositoryRoot, 'frontend-react/packages/core-model/src/generated-protection.ts'), coreProtection, 'utf8'),
]);
