/**
 * Shared policy-type display metadata — Tier 1 (UI).
 *
 * Colour + short-label vocabulary for the GSA policy types, keyed by the
 * definition files' full `displayName` (the single source of truth — see
 * architecture.md §4). Extracted so the unified policy table and the tenant
 * posture strip render the same badge for the same policy type.
 */

import type { BadgeProps } from '@fluentui/react-components';

/** Badge colour for each policy type (keyed by definition `displayName`). */
export const POLICY_TYPE_COLOR: Record<string, BadgeProps['color']> = {
  'Web Content Filtering Policy': 'brand',
  'TLS Inspection Policy': 'informative',
  'Threat Intelligence Policy': 'severe',
  'Cloud Firewall Policy': 'warning',
  'Prompt Policy': 'success',
  'Content Policy': 'important',
  'Forwarding Policy': 'subtle',
};

/**
 * Short, table-friendly label for each policy type (UI-only — the definition
 * files keep the full `displayName` as the source of truth). Falls back to the
 * full label when a type is not mapped.
 */
export const POLICY_TYPE_SHORT: Record<string, string> = {
  'Web Content Filtering Policy': 'Web Content',
  'TLS Inspection Policy': 'TLS inspection',
  'Threat Intelligence Policy': 'Threat Intel',
  'Cloud Firewall Policy': 'Cloud Firewall',
  'Prompt Policy': 'Prompt',
  'Content Policy': 'Content',
  'Forwarding Policy': 'Forwarding',
};

/** Short label for a policy type, falling back to the full label. */
export function shortPolicyType(label: string): string {
  return POLICY_TYPE_SHORT[label] ?? label;
}
