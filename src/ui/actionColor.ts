import type { BadgeProps } from '@fluentui/react-components';

/**
 * Single source of truth for policy / rule action badge colours.
 *
 * block = danger · allow|grant = success · inspect = informative ·
 * bypass = warning · scanPurview = important (Microsoft Purview content scan) ·
 * everything else = subtle. Shared by the policy table, the What-If panel, and
 * any other surface that renders an action so the posture vocabulary stays
 * identical everywhere (see the ux-design-review skill).
 */
export function actionColor(action: string): BadgeProps['color'] {
  switch (action) {
    case 'block':
      return 'danger';
    case 'allow':
    case 'grant':
      return 'success';
    case 'inspect':
      return 'informative';
    case 'bypass':
      return 'warning';
    case 'scanPurview':
      return 'important';
    default:
      return 'subtle';
  }
}

/**
 * Coerce a policy/rule `action` to a display string.
 *
 * Most actions are simple strings (`allow` / `block` / `inspect` / …), but some
 * GSA policy kinds (e.g. M365 / forwarding rules) carry a structured action
 * **object** (`{ '@odata.type', headerSettings, … }`) instead. Rendering such an
 * object directly throws "Objects are not valid as a React child", so every
 * surface that displays an action must funnel it through here first. For an
 * object we surface the `@odata.type` suffix as a readable token; anything
 * unrecognised becomes an empty string (rendered as an em-dash).
 */
export function actionToString(action: unknown): string {
  if (typeof action === 'string') return action;
  if (action && typeof action === 'object') {
    const odataType = (action as { '@odata.type'?: unknown })['@odata.type'];
    if (typeof odataType === 'string') return odataType.split('.').pop() ?? '';
  }
  return '';
}
