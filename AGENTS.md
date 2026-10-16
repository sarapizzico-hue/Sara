# AGENTS.md

## Cursor Cloud specific instructions

This repository is a **static website** (deployed via GitHub Pages — note the `.nojekyll` marker). It is a flat collection of self-contained, Portuguese-language (`pt-BR`) HTML documents at the repo root: a strategic sales dashboard (`index.html`), commercial proposals (`proposta-*.html`), sales scripts (`script-*.html`), playbooks/action plans (`playbook-*.html`, `plano-acao-*.html`), and analyses (`analise-*.html`, `visao-*.html`). Companion `.docx`/`.pptx`/`.pdf`/`.png` files are downloadable exports, not part of the runtime. Shared media lives in `assets/`.

### Key facts

- **No build system, no package manager, no backend, no database.** There are no lockfiles (`package.json`, `requirements.txt`, etc.) and nothing to install. The update/startup script is intentionally a no-op.
- Each HTML file is fully self-contained: inlined CSS and inline vanilla JS. There is no shared bundle.
- Interactivity (tabs, charts) is client-side. Charts use **Chart.js** loaded from `cdn.jsdelivr.net`, and the Inter font loads from Google Fonts. These are the only external dependencies and are non-blocking — pages still render offline (charts/fonts degrade gracefully).

### Run / preview (development)

Serve the repo root over HTTP to reproduce GitHub Pages hosting (Python 3 and Node are preinstalled in the environment):

```
python3 -m http.server 8000
```

Then open e.g. `http://localhost:8000/index.html`. Files can also be opened directly via `file://`, but serving over HTTP most closely matches production.

### Lint / test / build

There are no automated tests, linters, or build steps in this repository. "Testing" means opening the HTML pages in a browser and confirming they render (styling, KPIs, charts) and that interactive tabs update content.
