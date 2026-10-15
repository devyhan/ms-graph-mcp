/**
 * mammoth ships no TypeScript declarations and this repo adds no `@types/*`
 * packages for one call site, so the untyped import resolves to `any` here.
 * `extract-docx.ts` narrows it to the single function it uses.
 */
declare module 'mammoth';
