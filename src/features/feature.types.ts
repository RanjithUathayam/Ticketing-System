// src/features/feature.types.ts
//
// Shared contract for the auto-discovered "Operations Tools" feature pages.
//
// Convention: every page lives in its own folder under `src/features/<key>/`
// and exports two things from an `index.ts` barrel:
//   - `meta`      : a `FeaturePageMeta` object (see below)
//   - `Component` : the standalone Angular component class for the page
//
// `scripts/generate-features.mjs` scans this folder at build/serve time and
// generates `src/features/features.generated.ts`, which is the only file the
// rest of the app (routes + layout menu) actually imports. Nothing outside
// this folder needs to be touched to add, remove, reorder, or disable a page —
// see `src/features/README.md` for the full walkthrough.

export type FeatureIcon =
  | 'iron'
  | 'truck'
  | 'shopping-bag'
  | 'file-text'
  | 'clipboard-list'
  | 'archive-box'
  | 'default';

export interface FeaturePageMeta {
  /** Unique, stable identifier — must match the folder name under src/features/. */
  key: string;
  /** Document/browser tab title. */
  title: string;
  /** Label shown in the sidebar navigation. */
  menuLabel: string;
  /** Route segment, relative to the `/ops` parent route (no leading slash). */
  path: string;
  /** Icon key rendered next to the menu label — see feature-icons.ts. */
  icon: FeatureIcon;
  /** Menu section heading these pages are grouped under. */
  group: string;
  /** Ascending sort order within the group. */
  order: number;
  /** Set to false to keep the page registered (so old links don't 404) but hide it from the menu and block navigation. */
  enabled: boolean;
  /** Short description shown as a subtitle in the menu (optional). */
  description?: string;
}

/** What the generator emits per discovered feature — meta plus a lazy loader for the route table. */
export interface GeneratedFeatureEntry extends FeaturePageMeta {
  loadComponent: () => Promise<unknown>;
}
