import { cleanup, render, screen } from '@testing-library/react';
import { afterEach, expect, it } from 'vitest';
import { App } from './App';

afterEach(cleanup);

it('shows only the app name', () => {
  render(<App />);
  const heading = screen.getByRole('heading', { level: 1 });
  expect(heading.textContent).toBe('Kams Highlights 13');
});
