# Token Usage & Cost Log

Figures come from the [Anthropic Console → Usage](https://console.anthropic.com/settings/usage) page.
This file is a manual record — update it after each session.

## Pricing reference (as of 2025-08)

| Model | Input | Output | Cache write | Cache read |
|---|---|---|---|---|
| claude-sonnet-4-6 | $3.00 / MTok | $15.00 / MTok | $3.75 / MTok | $0.30 / MTok |

*MTok = 1 million tokens. Prices subject to change — verify at console.anthropic.com/settings/usage.*

---

## Session log

| Date | Session ID | Model | Input tok | Output tok | Cache write | Cache read | Est. cost (USD) | Notes |
|---|---|---|---|---|---|---|---|---|
| 2026-09-01 | 75810552-7b60-40a3-98c0-df8bc44ce941 | claude-sonnet-4-6 | — | — | — | — | — | PerfWeeklyReport dashboard enhancements (scripting task Ready/Design counts, RW history, filter/UI tweaks) |

---

## How to update this log

1. Go to [console.anthropic.com/settings/usage](https://console.anthropic.com/settings/usage)
2. Filter by date and/or API key to find the session
3. Copy input, output, and cache token counts into the row above
4. Compute cost: `(input * 3 + output * 15 + cache_write * 3.75 + cache_read * 0.30) / 1_000_000`
