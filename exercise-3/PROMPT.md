# Exercise 3: Scrollytelling

You have research files from an investigation into forced evictions in Assam,
India. Your task: turn them into an interactive scrollytelling page.

## Files

- `findings.json` — 11 structured findings with sources and confidence levels
- `article.md` — full article draft

## Prompt

```
Read the files in this folder. Build a single standalone index.html file
that tells this story as a scrollytelling piece:

- 6-8 sections, each with a short editorial text block (2-3 sentences)
- Sticky full-viewport backgrounds that change as you scroll (use colors,
  data visualizations, or typographic treatments — no external images needed)
- IntersectionObserver to trigger section transitions
- Dark theme, clean typography
- Source citations at the bottom of each text card
- No frameworks, no build step — just HTML/CSS/JS that opens in a browser

The narrative arc: scale of the crisis → who is behind it → what the
evidence shows → the human cost.
```
