# Operations Tools — auto-discovered feature pages

This folder holds every "Operations Tools" page (ironing/packing, shipment
dashboards, manifest extraction, box packing, etc.). Each page is a
self-contained Angular feature folder that gets picked up automatically —
**you never need to edit `app.routes.ts` or the sidebar to add, remove, or
reorder a page.**

## How discovery works

`scripts/generate-features.mjs` scans every folder in `src/features/` for a
`feature.meta.json`. For each one it finds, it writes an entry into
`src/features/features.generated.ts`, which exports:

- `featureRoutes` — an Angular `Routes` array, spread into the `/ops/*`
  route in `src/app.routes.ts` (that one line is the only static wiring in
  the whole app; it never needs to change when you add a page).
- `FEATURE_MENU` — the sorted metadata array the sidebar
  (`layout.component.ts`/`.html`) renders into nav links, grouped by
  `group` and ordered by `order`.

The generator runs automatically before `npm run dev` and `npm run build`
(via the `predev`/`prebuild` npm lifecycle hooks), so a normal `npm run dev`
always reflects whatever pages currently exist on disk. While `ng serve` is
already running, run `npm run generate:features:watch` alongside it in a
second terminal to pick up new/edited pages without restarting.

`src/features/features.generated.ts` is committed so a fresh checkout builds
even before anyone has run the generator — treat it as a build artifact,
though: don't hand-edit it, your changes will be overwritten.

## Adding a new page

1. Create `src/features/<your-key>/`.
2. Add `feature.meta.json` (shape documented by `FeaturePageMeta` in
   `feature.types.ts`):

   ```json
   {
     "key": "your-key",
     "title": "Your Tool Title",
     "menuLabel": "Sidebar Label",
     "path": "your-key",
     "icon": "clipboard-list",
     "group": "Operations Tools",
     "order": 70,
     "enabled": true,
     "description": "One line shown as the nav-link tooltip"
   }
   ```

   `icon` must be one of the keys in `feature-icons.ts` (add a new one there
   if none fit). `path` and `key` must each be unique across all features —
   the generator throws a clear error if two features collide.

3. Add `index.ts`:

   ```ts
   export { YourComponent as FeatureComponent } from './your-key.component';
   ```

4. Build your standalone Angular component as `your-key.component.ts` (+
   `.component.html`). It will be mounted at `/ops/your-key`, behind
   `OpsGuard` (Admin/Manager, same as the other Operations Tools pages —
   see `src/guards/ops.guard.ts` if you need different access rules) and
   inside the existing app shell (`layout.component`).
5. Run `npm run generate:features` (or just start `npm run dev`, which does
   it for you). The route and the sidebar entry both appear with no other
   changes.

## Porting an existing standalone HTML tool

Several pages here (`box-packing`, `ironing-packing`, `manifest-extractor`,
`pending-po`, `store-orders`, `shipment-dashboard`) started as
self-contained HTML files that manipulated the DOM directly rather than
using Angular bindings. Rather than risk subtly breaking their business
logic in a full rewrite, they're ported using the pattern in
`legacy-tool-base.ts`:

- The component's template is the tool's original body markup, wrapped in a
  `<div class="fx-<key>-scope" #root>` and a matching `<style>` block at the
  bottom of the `.component.html` (Angular's emulated view encapsulation
  scopes everything in it automatically — the only two things you must
  rewrite by hand are `:root { ... }` → `.fx-<key>-scope { ... }` and
  `body { ... }` → `.fx-<key>-scope { ... }`, since Angular does **not**
  scope `:root`/`html`/`body` selectors and they would otherwise leak
  page-wide).
- The component class extends `LegacyToolBase` and keeps the original
  script's logic almost verbatim: top-level `let` state becomes instance
  fields, top-level functions become private methods referencing `this`,
  and `document.getElementById('x')` becomes `this.q('x')` (scoped to the
  component's own root element, so IDs from one tool can never collide with
  another tool or the app shell).
- External libraries are swapped for the app's existing dependencies
  instead of loading a duplicate copy from a CDN: `xlsx` (already a
  dependency), `pdfjs-dist` (already configured — see
  `LabelPdfService`/`angular.json`'s `pdf.worker.min.mjs` asset), and
  `Chart` (already loaded globally in `index.html`, referenced the same way
  `dashboard.component.ts`/`reports.component.ts` already do via
  `declare var Chart: any;`). `exceljs` was added as a new dependency only
  for `pending-po`, since it's the one tool whose visual styling (merged,
  colored, bordered cells) genuinely depends on it — the community `xlsx`
  package silently ignores cell styles when writing a workbook.
- Everything runs under this app's `provideZonelessChangeDetection()`.
  These components deliberately never rely on Angular's template bindings
  or change detection to update the DOM — they do their own imperative DOM
  updates exactly as the original script did, which is what makes them
  correct under zoneless change detection without extra signal plumbing.
