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
        brand: {
          DEFAULT: '#107C41',
          dark: '#0B6A37',
          deep: '#07572D',
          soft: '#EEF5F0',
          pale: '#F7FBF8',
          line: '#D5E7DC',
          ink: '#173326',
        },
        chrome: '#FBFBFC',
      },
      boxShadow: {
        panel: '0 14px 36px rgba(23, 32, 51, 0.08)',
        'hub-card': '0 1px 3px rgba(28, 48, 38, 0.05), 0 7px 20px rgba(28, 48, 38, 0.05)',
        'brand-sm': '0 3px 9px rgba(16, 124, 65, 0.18)',
      },
    },
  },
  plugins: [],
};
