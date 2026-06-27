import { defineConfig } from 'tsup';

// Single bundled distribution: inline @eigenpal/docx-editor-{core,react,i18n,agents}
// (and their third-party deps like docxtemplater/jszip/dompurify) into one package,
// keeping ONLY react/react-dom/prosemirror as externals so the consumer dedupes a
// single shared copy (the docx-editor renderer + a headless engine must agree on
// one React/ProseMirror instance).
export default defineConfig({
  entry: { index: 'src/index.ts' },
  format: ['esm', 'cjs'],
  // resolve:true inlines the re-exported @eigenpal/* types so the published
  // .d.ts is self-contained (no dangling `from '@eigenpal/...'`).
  dts: { resolve: true },
  splitting: true,
  treeshake: true,
  minify: true,
  clean: true,
  sourcemap: false,
  external: [
    'react',
    'react-dom',
    'prosemirror-commands',
    'prosemirror-dropcursor',
    'prosemirror-history',
    'prosemirror-keymap',
    'prosemirror-model',
    'prosemirror-state',
    'prosemirror-tables',
    'prosemirror-transform',
    'prosemirror-view',
  ],
  // Force-bundle the workspace packages even though they resolve via node_modules.
  noExternal: [/^@eigenpal\/docx-editor-/],
});
