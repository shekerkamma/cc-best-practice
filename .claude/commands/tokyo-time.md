---
description: Display the current time in Tokyo, Japan (JST, UTC+9)
model: haiku
---

# Tokyo Time Command

Display the current time in Japan Standard Time by invoking the `time-tokyo` skill.

## Workflow

Use the Skill tool to invoke the `time-tokyo` skill:
- skill: time-tokyo

The skill provides the bash command and output format. Return its output directly to the user — no extra commentary.

## Critical Requirements

1. **Use Skill tool**: Do NOT run bash directly. Invoke via the `Skill` tool so the skill stays the single source of truth for the timezone logic.
2. **No extra output**: Just the formatted time line the skill specifies.
