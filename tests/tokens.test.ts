import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '..', 'src', 'app', 'globals.css'),
  'utf8',
);

/** Values from a.ui-spec.md §1 and phase 2.4. */
const TOKENS: Record<string, string> = {
  '--color-bg': '#f3f3f1',
  '--color-surface': '#ffffff',
  '--color-page': '#ffffff',
  '--color-page-text': '#111111',
  '--color-line': '#e2e2dd',
  '--color-text': '#1b1b1b',
  '--color-muted': '#6b6b66',
  '--color-accent': '#808000',
  '--color-accent-hover': '#6b6b00',
  '--color-marker': '#808000',
  '--color-marker-text': '#000000',
  '--color-gap-tint': 'rgb(128 128 0 / 0.33)',
  '--color-danger': '#b3261e',
  '--color-saved': '#3f8f3f',
  '--radius-card': '10px',
  '--radius-mark': '4px',
};

describe('design tokens (globals.css)', () => {
  it.each(Object.entries(TOKENS))('%s is %s', (name, value) => {
    expect(css).toContain(`${name}: ${value};`);
  });

  it('sets touch-action: manipulation on the app root', () => {
    expect(css).toMatch(/#root\s*\{[^}]*touch-action:\s*manipulation;/);
  });
});
