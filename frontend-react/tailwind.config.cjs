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
        accent: '#217345',
        brand: {
          DEFAULT: '#217345',
          dark: '#185b37',
          deep: '#12482b',
          soft: '#eef4f0',
          pale: '#f8faf9',
          line: '#cbd8d0',
          ink: '#3a4b42',
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
