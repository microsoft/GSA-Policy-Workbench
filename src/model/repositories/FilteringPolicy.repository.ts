/**
 * Filtering Policy repository — Tier 2b.
 *
 * Exposes read operations for Filtering Policies as a flat list
 * (all policy types: WCF, TLS, CloudFirewall, ThreatIntelligence).
 *
 * Note: For V0 the policy tree (profiles → policies → rules) is loaded in
 * one shot by the SecurityProfile repository via the adaptive loader.
 * This repository provides the standalone flat policy list — useful for
 * policy management views and V1+ write operations.
 *
 * See docs/architecture.md §5 for the repository contract.
 */

import { useQuery } from '@tanstack/react-query';
import { connectorList } from '../../adapters/graph/connector';
import { mapCollection } from '../../adapters/graph/mapper';
import { WebContentFilteringPolicyDefinition } from '../definitions/WebContentFilteringPolicy.definition';
import type { WebContentFilteringPolicy } from '../definitions/WebContentFilteringPolicy.definition';
import { queryKeys } from '../../query/keys';
import { REQUIRED_SCOPES } from '../../auth/scopes';

const FEATURE = 'filtering-policies';

/**
 * Fetch and cache the flat list of all filtering policies in the tenant.
 *
 * Returns all policy types (WCF, TLS, Cloud Firewall, Threat Intelligence)
 * as `WebContentFilteringPolicy[]` — the `@odata.type` discriminant on each
 * item identifies the concrete type.
 *
 * Does NOT include policyRules inline. Use `useProfileTree()` from the
 * SecurityProfile repository to get the full tree with rules.
 *
 * @example
 * const { data: policies = [], isLoading } = useFilteringPolicies();
 */
export function useFilteringPolicies() {
  return useQuery<WebContentFilteringPolicy[], Error>({
    queryKey: queryKeys.filteringPolicies(),
    queryFn: async ({ signal }) => {
      const raw = await connectorList<WebContentFilteringPolicy>(
        WebContentFilteringPolicyDefinition.operations.list.urlTemplate,
        {
          scopes: REQUIRED_SCOPES,
          feature: FEATURE,
          signal,
        },
      );
      return mapCollection<WebContentFilteringPolicy>(
        raw,
        WebContentFilteringPolicyDefinition.properties,
      );
    },
    staleTime: 60_000,
  });
}

