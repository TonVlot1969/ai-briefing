# AI Briefing

Ton's weekly bilingual (NL/EN) spoken AI briefing, built for listening in the car on Monday morning.

- **Live app:** https://tonvlot1969.github.io/ai-briefing/ai-briefing.html
- **Archive:** https://tonvlot1969.github.io/ai-briefing/archive.html

## How it works

`ai-briefing.html` is a single self-contained page. Every edition lives inline in a
`<script id="data" type="application/json">` block, with an NL/EN toggle and an
in-app edition picker. Each paragraph carries two versions: `t` is what you read,
`s` is what the speech synthesiser says, with things like "A I" and "M C P" spelled
out so the voice does not mangle them.

`briefings/briefing-YYYY-MM-DD.html` is a dated snapshot of the whole app.
`archive.html` links to every snapshot, newest first.

## Publishing

Generated and pushed every Friday at 12:00 by a scheduled Claude Cowork task, so it
is ready to listen to on Monday.

Two rules exist because breaking them lost four editions:

1. **The live site is not the source of truth.** The local working folder is. Each run
   merges local and published editions by date and asserts the count never drops.
2. **Push everything in one commit.** A partial push orphaned 31 July 2026. The 12, 19
   and 26 June editions went missing the same way: published as pages, never listed
   anywhere, invisible for two months.
