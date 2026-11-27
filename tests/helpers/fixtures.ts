/**
 * Test fixture helpers (execution context §6, phase 2.3).
 *
 * - Local fixtures: the owner's real WPS files in a.fixtures/. They are never
 *   committed, so they are missing in CI: wrap tests that need one in
 *   `describeWithLocalFixture(...)`, which SKIPS (does not fail) when the file is absent.
 * - Generated fixtures: synthetic files in tests/fixtures/generated/, written by
 *   the Vitest globalSetup (tests/fixtures/make-fixtures.ts). Always available.
 */
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe } from 'vitest';
import { GENERATED_DIR, type FixtureName } from '../fixtures/make-fixtures.ts';

const PROJECT_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');

/** `KH13_LOCAL_FIXTURES` overrides the folder, e.g. to simulate CI (a missing folder). */
export const LOCAL_FIXTURES_DIR =
  process.env.KH13_LOCAL_FIXTURES ?? join(PROJECT_ROOT, 'a.fixtures');
export const GENERATED_FIXTURES_DIR = GENERATED_DIR;

/** The owner's real WPS files (local only). */
export type LocalFixtureName = 'test.docx' | 'fx-a-headings-list.docx';

export function hasLocalFixture(name: LocalFixtureName): boolean {
  return existsSync(join(LOCAL_FIXTURES_DIR, name));
}

/**
 * Reads a file from a.fixtures/. Throws when it is missing: call it only inside
 * `describeWithLocalFixture`, which skips the block in that case.
 */
export function readLocalFixture(name: LocalFixtureName): Uint8Array {
  const path = join(LOCAL_FIXTURES_DIR, name);
  if (!existsSync(path)) {
    throw new Error(
      `Local fixture ${name} is missing (${path}). Wrap the test in describeWithLocalFixture().`,
    );
  }
  return readFileSync(path);
}

/** `describe` that is skipped (not failed) when the local fixture is absent, e.g. in CI. */
export function describeWithLocalFixture(
  name: LocalFixtureName,
  title: string,
  fn: () => void,
): void {
  describe.skipIf(!hasLocalFixture(name))(`${title} [local: ${name}]`, fn);
}

/** Reads a synthetic fixture from tests/fixtures/generated/. */
export function readGeneratedFixture(name: FixtureName): Uint8Array {
  const path = join(GENERATED_FIXTURES_DIR, name);
  if (!existsSync(path)) {
    throw new Error(
      `Generated fixture ${name} is missing (${path}). It is written by the Vitest globalSetup; ` +
        'run the tests through `npm test`, or run `npm run fixtures`.',
    );
  }
  return readFileSync(path);
}
