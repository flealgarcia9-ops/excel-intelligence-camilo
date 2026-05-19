# Skill Registry — tareacam

Generated: 2026-05-18

## Matched Skills

| Skill | Scope | Trigger | Why Matched |
|-------|-------|---------|-------------|
| react-19 | User | React 19 components | Project uses react ^19.1.1 |
| tailwind-4 | User | Tailwind CSS styling | index.css is 42KB — likely Tailwind-generated |
| vercel-react-best-practices | User | React/Next.js performance | React SPA optimization guidelines |
| playwright | User | E2E testing | Available for future E2E needs |
| typescript | User | TypeScript code | @types/react in devDependencies |

## Project Conventions

- **ES Modules**: `"type": "module"` in package.json
- **Flat ESLint config**: `eslint.config.js` with `@eslint/js`
- **React Hooks rules**: `eslint-plugin-react-hooks` configured
- **Vitest for testing**: `vitest.config.js` with jsdom environment
- **Vite build tool**: Standard Vite + React plugin setup

## Agent Instructions

- No `AGENTS.md` or `.cursorrules` found in project root
- Follow standard React 19 patterns (no useMemo/useCallback needed per react-19 skill)
