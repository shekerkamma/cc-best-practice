---
name: presentation
description: Update, modify, or fix presentation slides, structure, styling, or weights. Use when the user mentions presentation, slides, slide deck, or wants to edit presentation content.
argument-hint: "[what to change — e.g., 'add a new slide about hooks', 'fix slide 12 layout']"
---

# Presentation Skill

Delegates all presentation work to the `presentation-curator` agent, which has three preloaded sub-skills:

- **presentation-structure** — slide format, weight system, navigation, section structure
- **presentation-styling** — CSS classes, component patterns, syntax highlighting
- **vibe-to-agentic-framework** — conceptual framework, narrative arc, level system

## Workflow

When the user asks to update, modify, or fix a presentation:

1. **Delegate** — use the Agent tool to invoke the `presentation-curator` agent:

```
Agent(subagent_type="presentation-curator", description="Update presentation", prompt="<user's request>")
```

2. **Report back** — relay the agent's summary: slides changed, total count, level transitions, renumbering.

## Why Delegate

The presentation-curator agent:
- Has all three sub-skills preloaded (structure, styling, framework)
- Self-evolves after every execution, updating its own skills to prevent knowledge drift
- Enforces sequential slide numbering, level integrity, and pattern consistency
- Knows the 4-level journey system (Low → Medium → High → Pro)

Never edit presentation files directly — always go through the agent.
