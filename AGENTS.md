# AGENTS.md

## Project

Outlook add-in built on Office.js ("add-in-pfi"), TypeScript + React 18, bundled via Webpack 5, styled with Tailwind CSS. Deployed to Vercel. Host is `MailHost`/`Mailbox` (see `manifest.xml`) — not Excel.

## Structure

- `src/taskpane/` — React entry point and UI (`taskpane.tsx`, `taskpane.html`, `taskpane.css`)
- `src/services/` — business logic (`mailService.ts`, `analysisService.ts`)
- `src/global.d.ts` — global type declarations
- `manifest.xml` — Office add-in manifest
- `assets/` — static assets

## Commands

- `npm run dev-server` / `npm start` — webpack dev server
- `npm run build` — production build
- `npm run build:dev` — development build

## Conventions

- TypeScript strict mode enabled (`tsconfig.json`).
- No test runner or linter configured yet — do not assume Jest/Vitest/ESLint exist.
- Office.js types via `@types/office-js`.

## SDD

This project uses Spec-Driven Development with Engram as the persistence backend (no OpenSpec files — architecture/decisions are stored in Engram memory, topic key prefix `sdd-init/Add_in_PFI`).
