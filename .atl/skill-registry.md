# Skill Registry — tareacam

Generated: 2026-06-02

## Available Skills (User Scope)

| Skill | Path | Trigger |
|-------|------|---------|
| branch-pr | `~/.codex/skills/branch-pr/SKILL.md` | creating, opening, or preparing PRs for review |
| chained-pr | `~/.codex/skills/chained-pr/SKILL.md` | PRs over 400 lines, stacked PRs, review slices |
| cognitive-doc-design | `~/.codex/skills/cognitive-doc-design/SKILL.md` | writing guides, READMEs, RFCs, onboarding, architecture, or review-facing docs |
| comment-writer | `~/.codex/skills/comment-writer/SKILL.md` | PR feedback, issue replies, reviews, Slack messages, or GitHub comments |
| issue-creation | `~/.codex/skills/issue-creation/SKILL.md` | creating GitHub issues, bug reports, or feature requests |
| judgment-day | `~/.codex/skills/judgment-day/SKILL.md` | judgment day, dual review, adversarial review, juzgar |
| skill-creator | `~/.codex/skills/skill-creator/SKILL.md` | new skills, agent instructions, documenting AI usage patterns |
| skill-improver | `~/.codex/skills/skill-improver/SKILL.md` | improve skills, audit skills, refactor skills, skill quality |
| work-unit-commits | `~/.codex/skills/work-unit-commits/SKILL.md` | implementation, commit splitting, chained PRs, or keeping tests and docs with code |

## SDD Skills (System)

| Skill | Path | Trigger |
|-------|------|---------|
| sdd-apply | `~/.codex/skills/sdd-apply/SKILL.md` | orchestrator launches apply for one or more change tasks |
| sdd-archive | `~/.codex/skills/sdd-archive/SKILL.md` | orchestrator launches archive after implementation and verification |
| sdd-design | `~/.codex/skills/sdd-design/SKILL.md` | orchestrator launches design for a change |
| sdd-explore | `~/.codex/skills/sdd-explore/SKILL.md` | orchestrator launches exploration or requirement clarification |
| sdd-init | `~/.codex/skills/sdd-init/SKILL.md` | sdd init, iniciar sdd, openspec init |
| sdd-onboard | `~/.codex/skills/sdd-onboard/SKILL.md` | orchestrator launches onboarding for the full SDD cycle |
| sdd-propose | `~/.codex/skills/sdd-propose/SKILL.md` | orchestrator launches proposal work for a change |
| sdd-spec | `~/.codex/skills/sdd-spec/SKILL.md` | orchestrator launches spec work for a change |
| sdd-tasks | `~/.codex/skills/sdd-tasks/SKILL.md` | orchestrator launches task planning for a change |
| sdd-verify | `~/.codex/skills/sdd-verify/SKILL.md` | SDD verification phase, verify change |

## Project Conventions

- **ES Modules**: `"type": "module"` in package.json
- **Flat ESLint config**: `eslint.config.js` with `@eslint/js`, `eslint-plugin-react-hooks`, `eslint-plugin-react-refresh`
- **React 19 patterns**: functional components with hooks
- **Custom CSS design system**: `index.css` contains tokens and component styles (not Tailwind)
- **Vitest for testing**: `vitest.config.js` with jsdom environment
- **Vite build tool**: base `./`, manual chunks for `xlsx` and `chart.js`
- **JavaScript project**: no TypeScript compiler; `@types/react` for IDE support only
- **No formatter** (Prettier/Biome) installed
- **No coverage tool** installed
- **CI/CD**: GitHub Actions → GitHub Pages (`deploy.yml`)

## Agent Instructions

- No `AGENTS.md`, `.cursorrules`, `CLAUDE.md`, `GEMINI.md`, or `copilot-instructions.md` found in project root.
