// Flat ESLint config (ESLint 8+ / powerbi-visuals-tools 5.6).
// Composes the Power BI Visuals recommended ruleset with TypeScript parsing.

import powerbiVisuals from "eslint-plugin-powerbi-visuals";
import tsParser from "@typescript-eslint/parser";
import tsPlugin from "@typescript-eslint/eslint-plugin";

export default [
    {
        ignores: [".tmp/**", "dist/**", "node_modules/**", "**/*.js", "**/*.mjs"]
    },
    powerbiVisuals.configs.recommended,
    {
        files: ["src/**/*.ts"],
        languageOptions: {
            parser: tsParser,
            ecmaVersion: 2017,
            sourceType: "module"
        },
        plugins: {
            "@typescript-eslint": tsPlugin
        },
        rules: {
            "@typescript-eslint/no-explicit-any": "warn",
            "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_" }]
        }
    }
];
