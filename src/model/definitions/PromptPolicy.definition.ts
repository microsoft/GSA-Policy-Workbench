/**
 * Prompt Protection Policy definition — Tier 2a.
 *
 * Single source of truth for the Prompt (Generative AI) Protection Policy GSA
 * object type. A profile can carry a promptPolicyLink.
 *
 * Shape mirrors the TLS inspection policy:
 *   • the policy-level default lives in `settings.defaultAction`
 *     ('allow' | 'block'), not a top-level `action`;
 *   • each rule (`promptRule`) carries its OWN `action`;
 *   • rule conditions live under `matchingConditions`, with a `scanResult`
 *     trigger and a list of `conversationSchemes` (predefined AI services by
 *     `schemeName`, or custom by `url`).
 *
 * Graph docs: https://learn.microsoft.com/en-us/graph/api/resources/networkaccess-promptpolicy
 * API version: beta
 */

// ---------------------------------------------------------------------------
// Supporting types
// ---------------------------------------------------------------------------

/** Policy-level settings controlling default prompt-protection behaviour. */
export interface PromptPolicySettings {
  /** What happens to prompts not matched by any rule. */
  defaultAction?: 'allow' | 'block';
  [key: string]: unknown;
}

/**
 * One conversation scheme a prompt rule matches against — either a predefined
 * AI service (`schemeName`, e.g. "chatGpt", "claude", "gemini") or a custom
 * endpoint (`url`).
 */
export interface ConversationScheme {
  '@odata.type'?: string;
  /** Predefined scheme name (predefinedConversationScheme). */
  schemeName?: string;
  /** Custom endpoint URL (customConversationScheme). */
  url?: string;
  jsonPath?: string;
  [key: string]: unknown;
}

/** The conditions that select prompts for a prompt rule. */
export interface PromptRuleMatchingConditions {
  /** The scan verdict that triggers the rule, e.g. "maliciousPromptDetected". */
  scanResult?: string;
  conversationSchemes?: ConversationScheme[];
  [key: string]: unknown;
}

/**
 * A single prompt protection rule. Note `action` is PER RULE.
 */
export interface PromptRule {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  priority?: number;
  action?: 'allow' | 'block';
  /** Whether matched prompts are logged ("always" | "never" | ...). */
  promptLogging?: string;
  settings?: { status?: 'enabled' | 'disabled'; [key: string]: unknown };
  matchingConditions?: PromptRuleMatchingConditions | null;
}

// ---------------------------------------------------------------------------
// Domain type
// ---------------------------------------------------------------------------

/**
 * A Prompt Protection Policy — scans generative-AI prompts. Rules (promptRule)
 * are in `policyRules`, fetched from
 * /networkAccess/promptPolicies/{id}/policyRules and inlined by the exporter
 * (the profile $expand does not return them).
 */
export interface PromptPolicy {
  '@odata.type'?: string;
  id: string;
  name?: string;
  description?: string;
  /** Read-only. Set by the API. */
  version?: string;
  lastModifiedDateTime?: string;
  settings?: PromptPolicySettings;
  /** Populated by the exporter from the policy's /policyRules sub-route. */
  policyRules?: PromptRule[];
}

// ---------------------------------------------------------------------------
// Definition
// ---------------------------------------------------------------------------

export const PromptPolicyDefinition = {
  odataType: '#microsoft.graph.networkaccess.promptPolicy',
  displayName: 'Prompt Policy',

  properties: {
    id:                   { label: 'ID' },
    name:                 { label: 'Policy Name' },
    description:          { label: 'Description' },
    version:              { label: 'Version' },
    lastModifiedDateTime: { label: 'Last Modified' },
    settings:             { label: 'Settings' },
    policyRules:          { label: 'Rules' },
  },

  operations: {
    list:   { method: 'GET',    urlTemplate: '/networkAccess/promptPolicies' },
    get:    { method: 'GET',    urlTemplate: '/networkAccess/promptPolicies/{id}' },
    create: { method: 'POST',   urlTemplate: '/networkAccess/promptPolicies' },
    update: { method: 'PATCH',  urlTemplate: '/networkAccess/promptPolicies/{id}' },
    delete: { method: 'DELETE', urlTemplate: '/networkAccess/promptPolicies/{id}' },
  },
} as const;
