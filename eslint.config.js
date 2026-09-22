const js = require("@eslint/js");
const globals = require("globals");

module.exports = [
  { ignores: ["node_modules/", "eslint.config.js"] },
  js.configs.recommended,
  {
    files: ["**/*.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: { ...globals.googleappsscript },
    },
    rules: {
      "no-unused-vars": "off",
      "no-undef": "off",
    },
  },
];