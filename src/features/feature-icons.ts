// src/features/feature-icons.ts
// Inline SVG paths (24x24, stroke-based, matches the sidebar's existing icon style)
// keyed by FeatureIcon so the auto-generated menu can render an icon without
// each feature needing to ship its own markup.
import { FeatureIcon } from './feature.types';

export const FEATURE_ICON_PATHS: Record<FeatureIcon, string> = {
  iron: 'M4 15h11a5 5 0 005-5V8.5A2.5 2.5 0 0017.5 6H9a5 5 0 00-5 5v4zm0 0v2a1 1 0 001 1h2m6-3l2 3',
  truck: 'M3 7h11v8H3V7zm11 3h4l3 3v2h-7v-5zM6.5 20a1.5 1.5 0 100-3 1.5 1.5 0 000 3zm11 0a1.5 1.5 0 100-3 1.5 1.5 0 000 3z',
  'shopping-bag': 'M6 7V6a4 4 0 118 0v1h3l1 14H2L3 7h3zm2 0h4V6a2 2 0 10-4 0v1z',
  'file-text': 'M7 3h7l5 5v13a1 1 0 01-1 1H7a1 1 0 01-1-1V4a1 1 0 011-1zm7 0v5h5M9 13h6m-6 4h6',
  'clipboard-list': 'M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 012-2h2a2 2 0 012 2M9 5a2 2 0 002 2h2a2 2 0 002-2m-6 8h6m-6 4h4',
  'archive-box': 'M3 7h18M4 7v12a1 1 0 001 1h14a1 1 0 001-1V7M3 7l1.5-3h15L21 7M10 12h4',
  default: 'M4 6h16M4 12h16M4 18h16',
};
