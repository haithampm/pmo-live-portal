# PMO Live Portal

External PMO web portal for project portfolio tracking.

## What this portal does

- Shows a professional PMO dashboard
- Opens Microsoft Lists for live project/task/risk/report data
- Opens Microsoft Forms for task updates and weekly reports
- Links to Teams / OneDrive / SharePoint project files
- Can be hosted on GitHub Pages, Cloudflare Pages, Netlify, or Vercel

## Local development

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

## Where to update links

Edit `src/main.jsx` and replace placeholder links inside the `CONFIG` object:

```js
const CONFIG = {
  microsoftLists: {
    projects: '...',
    tasks: '...',
    risks: '...',
    reports: '...',
    resources: '...',
  },
  forms: {
    addTask: '...',
    updateTask: '...',
    weeklyReport: '...',
    registerRisk: '...',
  },
};
```

Also update each project object with real Teams file links:

```js
files: '...',
schedule: '...',
reports: '...',
docs: '...'
```

## Recommended hosting

Cloudflare Pages is recommended for quick free hosting.

Build command:

```bash
npm run build
```

Build output directory:

```text
dist
```
