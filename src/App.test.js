import { render, screen } from '@testing-library/react';
import App from './App';

test('renders the landing experience', () => {
  render(<App />);
  expect(screen.getByText(/make spreadsheets feel alive/i)).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /open dashboard/i })).toBeInTheDocument();
});
