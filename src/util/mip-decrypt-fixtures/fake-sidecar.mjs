#!/usr/bin/env node
/**
 * Fake MIP decryptor sidecar for tests. Speaks the real argv protocol
 * (`<sidecar> <inPath> <outPath>`, exit 0/1/2, `[mip]` stderr) with synthetic
 * behaviour selected by FAKE_MIP_MODE. Never touches the network.
 *
 * Modes: success | fail | usage | no-rights | hang | check-mode | empty
 */

import { copyFileSync, readFileSync, statSync, writeFileSync } from 'node:fs';

const mode = process.env['FAKE_MIP_MODE'] ?? 'success';

if (mode === 'hang') {
  // A bare unsettled promise lets the loop drain and node exits 13; the timer
  // keeps it alive so the launcher's timeout is what fires.
  await new Promise((resolve) => setTimeout(resolve, 3600_000));
}

const [, , inPath, outPath] = process.argv;
if (mode === 'usage' || inPath === undefined || outPath === undefined) {
  process.stderr.write('[mip] usage: fake-sidecar <input> <output>\n');
  process.exit(2);
}

if (mode === 'check-mode') {
  // The launcher must create the input file 0600; anything wider fails here.
  const got = (statSync(inPath).mode & 0o777).toString(8);
  if (got !== '600') {
    process.stderr.write(`[mip] input mode is ${got}, want 600\n`);
    process.exit(1);
  }
  copyFileSync(inPath, outPath);
  process.stderr.write('[mip] mode ok\n');
  process.exit(0);
}

if (mode === 'no-rights') {
  process.stderr.write('[mip] NO_RIGHTS: identity has no usage rights on this file\n');
  process.exit(1);
}

if (mode === 'fail') {
  process.stderr.write('[mip] ERROR: synthetic decrypt failure\n');
  process.exit(1);
}

if (mode === 'empty') {
  writeFileSync(outPath, Buffer.alloc(0));
  process.stderr.write('[mip] wrote empty output\n');
  process.exit(0);
}

// success: prefix the bytes so tests can tell output from input.
const input = readFileSync(inPath);
writeFileSync(outPath, Buffer.concat([Buffer.from('PLAIN:', 'utf8'), input]));
process.stderr.write('[mip] decrypted synthetic fixture\n');
process.exit(0);
