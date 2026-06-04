# KK Brand & Content Hub

A self-contained, GitHub Pages–ready static site that assembles the 2026 content
strategy for **Kalpesh Kinariwala** — Founder & Chairman of **Pantheon Development**
and Founder of **HOP Events** ([@kalpesh.kinariwala](https://www.instagram.com/kalpesh.kinariwala/)).

Everything lives in one place: a landing page, three interactive content hubs, the
written strategy documents, and the underlying research — all built around a 3‑pillar
system: **Emotional (40%)**, **Growth (35%)**, and **Controversy (25%)**.

## What's inside

| Page | Path | What it is |
|------|------|------------|
| **Landing page** | [`index.html`](index.html) | Entry point linking everything below |
| **Content Production Hub** | [`hubs/content-production-hub.html`](hubs/content-production-hub.html) | Production playbook — 10 ready-to-shoot templates (E1–E4, G1–G3, C1–C3) with briefs & scripts |
| **3-Pillar Content Reference** | [`hubs/3-pillar-reference.html`](hubs/3-pillar-reference.html) | Tabbed reference — KK reels vs. competitor reels, caption formulas, weekly plan, content bot |
| **Brand Intelligence Hub** | [`hubs/brand-intelligence-hub.html`](hubs/brand-intelligence-hub.html) | Post-level dataset across KK + 4 peer accounts, diagnostics, controversy ideas |
| **Document viewer** | [`docs.html`](docs.html) | Renders the Markdown docs below in-browser |

### Strategy documents
- [`strategy/instagram-growth-strategy-2026.md`](strategy/instagram-growth-strategy-2026.md) — current state, what's working, gaps, the 90-day action plan, LinkedIn cross-posting, KPIs, quick wins
- [`strategy/social-media-content-plan-2026.md`](strategy/social-media-content-plan-2026.md) — the full 3-pillar system, 30 post ideas, weekly calendar, monthly rhythm, caption templates, red lines

### Research & reference
- [`research/uae-real-estate-founders-linkedin.md`](research/uae-real-estate-founders-linkedin.md) — Top 50 UAE developer & PropTech founders with LinkedIn profiles, tiered for outreach
- [`research/proptech-event-companies-uae.md`](research/proptech-event-companies-uae.md) — UAE & global PropTech events to attend/speak at + event partners, with a pitch priority order
- [`research/proptech-ai-thought-leadership.md`](research/proptech-ai-thought-leadership.md) — LinkedIn strategy, five thought-leadership pillars, and six ready-to-post articles

## Project structure

```
.
├── index.html                  # Landing page
├── docs.html                   # In-browser Markdown viewer
├── .nojekyll                   # Serve files as-is on GitHub Pages
├── hubs/
│   ├── content-production-hub.html
│   ├── 3-pillar-reference.html
│   └── brand-intelligence-hub.html
├── strategy/
│   ├── instagram-growth-strategy-2026.md
│   └── social-media-content-plan-2026.md
└── research/
    ├── uae-real-estate-founders-linkedin.md
    ├── proptech-event-companies-uae.md
    └── proptech-ai-thought-leadership.md
```

## View it locally

The three hubs are fully standalone — just open any `hubs/*.html` in a browser.

`index.html` and `docs.html` load the Markdown files with `fetch()`, which browsers
block over `file://`. To preview the whole site, serve the folder over HTTP:

```bash
# from the repo root
python3 -m http.server 8000
# then open http://localhost:8000/
```

## Deploy on GitHub Pages

**Live site:** https://aistudio-pantheon.github.io/kk/

This site is published from the **`gh-pages`** branch (GitHub Pages → *Deploy from a branch* →
`gh-pages` / root). To update the live site, publish the latest files to that branch — for
example, from the site source:

```bash
git switch --orphan gh-pages   # first time only
git add -A && git commit -m "Publish site"
git push -f origin gh-pages
```

The included `.nojekyll` file disables Jekyll processing so every file is served exactly as-is.

## Notes

- The hubs embed live Instagram posts via Instagram's official `embed.js` (needs internet at
  view time). The doc viewer renders Markdown with [`marked`](https://marked.js.org/), which is
  **vendored locally** at `assets/vendor/marked.min.js` — so `docs.html` works with no external
  dependency.
- Engagement figures and post references reflect the analysis window of **May 2025 – May 2026**.
- This is a content-strategy reference site — there is no build step and no backend.
