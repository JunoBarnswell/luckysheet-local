/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./apps/web/**/*.{ts,tsx}', './packages/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        ink: '#172033',
        muted: '#657089',
        canvas: '#f5f7fb',
        line: '#dfe5ef',
        accent: '#375dfb',
      },
      boxShadow: {
        panel: '0 14px 36px rgba(23, 32, 51, 0.08)',
      },
    },
  },
  plugins: [],
};
