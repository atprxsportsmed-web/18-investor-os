# Deck source

14 slides, in order (see `deck.json`): cover, problem, story, offer, edge, market, tiers, team, unit, projections, funds, returns, ask, next.

Live copy: Slides artifact https://claude.ai/artifact/1sFwrvCpMgo6QJ4cmj42HM. Jeremy also edits that page by hand, so the repo can be behind it.

## Sync rules
- Before editing, read the live files from the artifact (Artifact tool, action `read`, with `paths` and an `out_dir`), copy them over `deck/`, then edit.
- Small changes: publish a slide with `edits` (old/new text pairs). Several slides, order changes or `deck.json`: publish `deck.json` with a `files` list of the changed slides, `root` set to the folder holding `project/`.
- After publishing, commit the changed files here.

## Slide format rules (a violation holds the slide and breaks page editing)
- Allowed tags: section, h1-h3, p, ul/ol/li, div, img, svg, table/tr/th/td, hr, aside, and inline b/i/u/a/span/br.
- Inline tags only inside text blocks (p, li, td), never directly inside a div.
- Page numbers are hardcoded in the bottom-right `<p>` of each slide. Renumber after inserting, removing or moving a slide.
- Speaker notes are the `<aside>` at the end of each slide. Notes carry the verify-before-use warnings.
