/**
 * Forwarding-profile definition — Tier 2a.
 *
 * A GSA **traffic-forwarding profile** decides whether a class of traffic is
 * routed through Global Secure Access at all, and — via its **acquisition
 * rules** (forwarding policy rules) — which specific destinations within that
 * traffic class are actually forwarded vs. explicitly bypassed. Before any
 * filtering or Private Access policy can take effect, the relevant profile
 * must be `enabled` AND the destination must be acquired by one of its rules.
 * This is What-If's client-side ("stage 1") evaluation.
 *
 * There is one GA profile per `trafficForwardingType`:
 *   • `internet` — Internet Access (Secure Web Gateway)
 *   • `m365`     — Microsoft 365 traffic
 *   • `private`  — Private Access
 *
 * Private previews add *multiple* profiles per type:
 *   • Preview A (`private`): different PA app subsets per user/device group
 *   • Preview B (`internet`): fail-close posture via `clientFallbackAction: block`
 * See docs/Support-Multiple-Forwarding-Profiles.md for the full design.
 *
 * Read-only. Uses `NetworkAccess.Read.All` (already required for the policy
 * tree), so it loads live as well as from a fixture — no new scopes.
 *
 * API version: beta (`/networkAccess/forwardingProfiles`).
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-forwardingprofile
 */

import type { RuleDestination } from './FilteringRule.definition';

export type TrafficForwardingType = 'internet' | 'm365' | 'private' | string;

export type ForwardingProfileState = 'enabled' | 'disabled' | string;

/**
 * A forwarding rule's disposition for traffic matching its destinations.
 * `bypass` = explicitly NOT sent to Global Secure Access (a traffic escape);
 * `forward` = tunneled to the GSA edge.
 */
export type ForwardingRuleAction = 'bypass' | 'forward' | string;

/**
 * What happens to in-scope traffic when the GSA edge is unreachable.
 * `bypass` = fail-open (traffic allowed through direct); `block` = fail-close
 * (traffic dropped). Preview B field — not yet in public Graph schema.
 */
export type ClientFallbackAction = 'bypass' | 'block' | string;

/**
 * An explicit PA app reference on a `private` forwarding profile.
 * Replaces destination-matching as the app→profile join mechanism.
 * Stored as appId references only — destinations are read at runtime from
 * `privateAccessApps[].applicationSegments`.
 */
export interface LinkedApp {
  appId: string;
  displayName?: string;
}

/**
 * One **acquisition rule** inside a forwarding policy — decides whether a
 * destination is forwarded to GSA or bypassed. Reuses the same
 * `RuleDestination` union as filtering rules (fqdn / url / ipAddress /
 * ipRange / ipSubnet / webCategory).
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-forwardingrule
 * M365 subtype: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-m365forwardingrule
 */
export interface ForwardingRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  /** `url` | `fqdn` | `ipAddress` | `ipRange` | `ipSubnet`. */
  ruleType?: string;
  action?: ForwardingRuleAction;
  destinations?: RuleDestination[];
  /** M365 Network Connectivity Principles category. `m365ForwardingRule` only. */
  category?: 'default' | 'optimized' | 'allow' | string;
  /** Port(s) this rule applies to. `m365ForwardingRule` only. */
  ports?: string[];
  /** Network protocol. `m365ForwardingRule` only. */
  protocol?: string;
  /** Per-rule fail-close posture when the edge is unreachable. `m365ForwardingRule` only. */
  clientFallbackAction?: ClientFallbackAction;
}

/** A forwarding policy — a named set of acquisition rules for one traffic type. */
export interface ForwardingPolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  version?: string;
  trafficForwardingType?: TrafficForwardingType;
  /** Fetched separately (`/networkAccess/forwardingPolicies/{id}/policyRules`) — beta `$expand` cannot nest 3 levels deep. */
  policyRules?: ForwardingRule[];
}

/** Links a forwarding policy to a forwarding profile. */
export interface ForwardingPolicyLink {
  '@odata.type'?: string;
  id: string;
  /** `enabled` | `disabled` — controls whether this service group's rules are applied. */
  state?: string;
  /** Evaluation order within the profile (IA profile only; null for M365/PA). */
  priority?: number | null;
  policy?: ForwardingPolicy;
}

/**
 * Who/what this profile's traffic applies to. Today Graph only documents the
 * `associatedBranch` (remote network) derived type publicly; the private
 * preview's per-user / per-device / device-platform assignment is not yet in
 * the public schema — captured as an opaque passthrough until confirmed.
 */
export interface ForwardingAssociation {
  '@odata.type'?: string;
  [key: string]: unknown;
}

/** A GSA traffic-forwarding profile (one per traffic class, today). */
export interface ForwardingProfile {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  version?: string;
  lastModifiedDateTime?: string;
  /** `enabled` when this traffic class is routed through GSA. */
  state?: ForwardingProfileState;
  /** Which traffic class this profile governs. */
  trafficForwardingType?: TrafficForwardingType;
  /** Profile evaluation priority (lower number = higher precedence, matching Security Profile convention — unconfirmed for the multi-profile PA preview). */
  priority?: number;
  /** Linked forwarding policies — each policy's `policyRules` are the acquisition rules. */
  policies?: ForwardingPolicyLink[];
  /** User/device/branch scoping — see `ForwardingAssociation` caveat above. */
  associations?: ForwardingAssociation[];
  /**
   * Fail-open (`bypass`) or fail-close (`block`) when the GSA edge is
   * unreachable. Preview B field — absent on GA profiles (treat as `bypass`).
   */
  clientFallbackAction?: ClientFallbackAction;
  /**
   * Explicit PA app membership for `private` forwarding profiles.
   * Fixture-level join: appId references only — no segment duplication.
   * Absent on `internet` / `m365` profiles (no discrete app objects).
   */
  linkedApps?: LinkedApp[];
}

export const ForwardingProfileDefinition = {
  odataType: '#microsoft.graph.networkaccess.forwardingProfile',
  displayName: 'Forwarding Profile',

  properties: {
    id:                    { label: 'ID' },
    name:                  { label: 'Name' },
    state:                 { label: 'State' },
    trafficForwardingType: { label: 'Traffic Type' },
    priority:              { label: 'Priority' },
    clientFallbackAction:  { label: 'Client Fallback Action' },
  },

  operations: {
    list: {
      method: 'GET',
      urlTemplate: '/networkAccess/forwardingProfiles',
      // `associations` is a plain property, not a navigation property — it
      // must be requested via $select, not $expand (a live-tenant 400
      // confirmed this: "Property 'associations' ... is not a navigation
      // property or complex property"). Only `policies` is expandable here.
      query:
        '$select=id,name,description,version,lastModifiedDateTime,priority,state,trafficForwardingType,associations' +
        '&$expand=policies($expand=policy)',
    },
    listPolicyRules: {
      method: 'GET',
      urlTemplate: '/networkAccess/forwardingPolicies/{id}/policyRules',
    },
  },
} as const;
