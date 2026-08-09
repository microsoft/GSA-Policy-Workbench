/**
 * Definition registry — maps every known GSA object type to its definition.
 *
 * Import this when you need to look up a definition by @odata.type at runtime
 * (e.g. to render a policy's friendly type label without hard-coding strings
 * in the UI). Definition files remain the single source of truth — this is
 * just an index over them.
 */

import { SecurityProfileDefinition } from './definitions/SecurityProfile.definition';
import { WebContentFilteringPolicyDefinition } from './definitions/WebContentFilteringPolicy.definition';
import { TlsInspectionPolicyDefinition } from './definitions/TlsInspectionPolicy.definition';
import { PromptPolicyDefinition } from './definitions/PromptPolicy.definition';
import { CloudFirewallPolicyDefinition } from './definitions/CloudFirewallPolicy.definition';
import { ThreatIntelligencePolicyDefinition } from './definitions/ThreatIntelligencePolicy.definition';
import { FilePolicyDefinition } from './definitions/FilePolicy.definition';
import { FilteringRuleDefinition } from './definitions/FilteringRule.definition';
import {
  PrivateAccessAppDefinition,
  AppProxyAppDefinition,
  AuthenticationStrengthPolicyDefinition,
} from './definitions/PrivateAccessApp.definition';
import { ForwardingProfileDefinition } from './definitions/ForwardingProfile.definition';

export const definitionRegistry = {
  [SecurityProfileDefinition.odataType]: SecurityProfileDefinition,
  [WebContentFilteringPolicyDefinition.odataType]: WebContentFilteringPolicyDefinition,
  [TlsInspectionPolicyDefinition.odataType]: TlsInspectionPolicyDefinition,
  [PromptPolicyDefinition.odataType]: PromptPolicyDefinition,
  [CloudFirewallPolicyDefinition.odataType]: CloudFirewallPolicyDefinition,
  [ThreatIntelligencePolicyDefinition.odataType]: ThreatIntelligencePolicyDefinition,
  [FilePolicyDefinition.odataType]: FilePolicyDefinition,
  [FilteringRuleDefinition.odataType]: FilteringRuleDefinition,
  [PrivateAccessAppDefinition.odataType]: PrivateAccessAppDefinition,
  [AppProxyAppDefinition.odataType]: AppProxyAppDefinition,
  [AuthenticationStrengthPolicyDefinition.odataType]: AuthenticationStrengthPolicyDefinition,
  [ForwardingProfileDefinition.odataType]: ForwardingProfileDefinition,
} as const;

export type KnownOdataType = keyof typeof definitionRegistry;

/**
 * Friendly display name for a policy's `@odata.type`, sourced from its
 * definition file. Returns `undefined` when the type is not in the registry
 * so callers can render an "Unsupported type" placeholder.
 */
export function policyTypeLabel(odataType: string | undefined): string | undefined {
  if (!odataType) return undefined;
  const def = (definitionRegistry as Record<string, { displayName: string }>)[odataType];
  return def?.displayName;
}

/**
 * Friendly label for a Security Profile policy **link** `@odata.type`.
 *
 * The link is the reliable discriminant: live Graph omits `@odata.type` on the
 * nested filtering policy, so the link type is what tells us which kind of
 * policy it is. Prompt/forwarding policies have no definition file yet, so
 * their labels live here directly.
 */
const LINK_TYPE_LABEL: Record<string, string> = {
  '#microsoft.graph.networkaccess.filteringPolicyLink':
    WebContentFilteringPolicyDefinition.displayName,
  '#microsoft.graph.networkaccess.tlsInspectionPolicyLink':
    TlsInspectionPolicyDefinition.displayName,
  '#microsoft.graph.networkaccess.threatIntelligencePolicyLink':
    ThreatIntelligencePolicyDefinition.displayName,
  '#microsoft.graph.networkaccess.cloudFirewallPolicyLink':
    CloudFirewallPolicyDefinition.displayName,
  '#microsoft.graph.networkaccess.promptPolicyLink': PromptPolicyDefinition.displayName,
  '#microsoft.graph.networkaccess.filePolicyLink': FilePolicyDefinition.displayName,
  '#microsoft.graph.networkaccess.forwardingPolicyLink': 'Forwarding Policy',
};

/**
 * Resolve a policy-type label from the link type and/or the nested policy type.
 * Prefers the nested policy's own type when present, then the link type.
 * Returns `undefined` when neither is recognised (→ "Unsupported type" in UI).
 */
export function resolvePolicyTypeLabel(
  linkOdataType: string | undefined,
  policyOdataType: string | undefined,
): string | undefined {
  return (
    policyTypeLabel(policyOdataType) ??
    (linkOdataType ? LINK_TYPE_LABEL[linkOdataType] : undefined)
  );
}
