import { tokens } from "@fluentui/react-components";

/** @type {import('tailwindcss').Config} */
export default {
  content: ["./src/**/*.{js,jsx,ts,tsx}"],
  theme: {
    extend: {
      colors: {
        primary: tokens.colorBrandBackground,
      },
      fontSize: {
        xxs: "0.5rem", // 8px
      },
    },
  },
  plugins: [],
};
