import { makeFixtures } from './fixtures/make-fixtures.ts';

/** Vitest globalSetup: regenerates tests/fixtures/generated/ once per test run. */
export default async function setup(): Promise<void> {
  await makeFixtures();
}
