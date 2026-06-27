# @sqren/docx-editor

A single bundled distribution of the [docx-editor](https://github.com/eigenpal/docx-editor)
fork — a read-only DOCX renderer / editor for React. It inlines the fork's
`@eigenpal/docx-editor-{core,react,i18n,agents}` packages into one published
package, so consumers install **one** dependency instead of four scoped ones.

`react`, `react-dom`, and `prosemirror-*` stay as peer dependencies so the
renderer shares a single React / ProseMirror instance with the rest of your app
(and with any companion engine, e.g. a headless editing engine).

## Install

```sh
npm install @sqren/docx-editor
# peers (if not already present):
npm install react react-dom prosemirror-{commands,dropcursor,history,keymap,model,state,tables,transform,view}
```

## Usage

```tsx
import { DocxEditor, type DocxEditorRef } from '@sqren/docx-editor';
import '@sqren/docx-editor/styles.css'; // pre-built (Tailwind utilities expanded)

<DocxEditor documentBuffer={bytes} readOnly />;
```

## Building / publishing (maintainers)

This package re-exports `@eigenpal/docx-editor-react` and bundles it with tsup.
Build the four `@eigenpal/*` workspace packages first, then:

```sh
bun run build      # tsup bundle + copy ../react/dist/styles.css
npm publish        # runs prepublishOnly (build) first
```
