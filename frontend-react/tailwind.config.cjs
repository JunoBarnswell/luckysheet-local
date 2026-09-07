/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./apps/web/**/*.{ts,tsx}', './packages/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        ink: '#3d3c41',
        muted: '#6c696d',
        canvas: '#ffffff',
        line: '#d9d9d9',
        accent: '#107C41',
        brand: {
          DEFAULT: '#107C41',
          dark: '#0B6A37',
          deep: '#07502A',
          soft: '#E5F2EB',
          pale: '#F5F5F5',
          line: '#D1D1D1',
          ink: '#242424',
        },
        chrome: '#f3f3f3',
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
