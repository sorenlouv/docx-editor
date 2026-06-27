/**
 * @sqren/docx-editor
 *
 * Single bundled distribution of the docx-editor fork. Re-exports the curated
 * React editor surface from `@eigenpal/docx-editor-react`; the build inlines the
 * core / i18n / agents packages so consumers install just this one package
 * (plus react/react-dom/prosemirror as peers).
 *
 * Stylesheet ships separately: `import "@sqren/docx-editor/styles.css"`.
 */
export * from '@eigenpal/docx-editor-react';
