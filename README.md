# BISC Union — current files
Generated 22 August 2026

## What's here

| File | Where it goes |
|---|---|
| `index.html` | Repo root — this is bisc.work |
| `big-2026-toolkit.html` | Same folder as index.html |
| `bisc-intake-backend.gs` | Apps Script editor (paste, replacing everything) |

## Live settings baked in

- **Endpoint** both apps post to: `.../AKfycbzunUCz71Mfak.../exec`
- **Sheet** the script writes to: `1LmKjohQDfLLKVfEyjI1O4Yg0pccZvSef0-_91LoNH7I`
- Filenames are case-sensitive on GitHub Pages. Use them exactly as written.

## Updating the script later

Deploy → Manage deployments → **pencil icon** → Version: **New version** → Deploy.

Do NOT use "New deployment" — that creates a different URL and the apps
keep talking to the old one. That caused a long debugging detour once already.

## Checking what's actually live

Open the /exec URL in a browser. The response tells you the version, the
script id, and which sheet it writes to. Add `?probe=feedback` to write a
test row without involving the app.

## Two things still unresolved in the copy

- The benefits-eligibility FAQ answer is marked "needs review before launch"
- The Active Member definition is marked "pending confirmation from structures"

## Turning off co-design feedback at launch

In both HTML files, find `const FEEDBACK_ON = true;` and set it to `false`.
The button and panel remove themselves.
