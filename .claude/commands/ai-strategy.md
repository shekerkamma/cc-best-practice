---
description: Research AI market intelligence and generate a comprehensive strategy document (.docx)
---

# AI Strategy Command

Research AI business strategy for a given topic or vertical and generate a comprehensive Word document with full references.

## Workflow

### Step 1: Get the Topic

If the user provided a topic with the command (e.g., `/ai-strategy AI-native insurance brokerage`), use that. Otherwise, use the AskUserQuestion tool to ask:

"What AI strategy topic or vertical should I research? Examples:
- AI-native service provider (general)
- AI-native insurance brokerage
- Enterprise AI deployment services
- AI-native legal services
- AI-native healthcare revenue cycle
- Competitive analysis: OpenAI vs Anthropic services strategy"

### Step 2: Launch the AI Strategy Agent

Use the Agent tool to invoke the ai-strategy-agent:
- subagent_type: ai-strategy-agent
- description: AI strategy research and document generation
- prompt: |
    Research the following AI strategy topic and generate a comprehensive Word document:

    Topic: [topic from Step 1]

    Follow your preloaded ai-strategy-researcher skill instructions exactly:
    1. Execute the 4-phase research framework (Market Signals, Competitive Intelligence, Operational Intelligence, Framework Application)
    2. Run parallel WebSearch calls where possible
    3. Generate a professional Word document with all 13 required sections
    4. Include minimum 30 referenced URLs with source attribution
    5. Save the document to the project root
    6. Clean up the generator script
    7. Report back: file path, 5 key findings, reference count

Wait for the agent to complete and capture the results.

### Step 3: Deliver Results

Report to the user:
1. File path to the generated document
2. The 5 key findings from the research
3. Total references collected
4. Offer to open the document (`wslview` on WSL2, `open` on macOS)

## Critical Requirements

1. **Use Agent Tool**: DO NOT do the research inline. Delegate to the ai-strategy-agent via the Agent tool.
2. **Topic Clarity**: Ensure the topic is specific enough for focused research. If too broad, ask the user to narrow it.
3. **Document Quality**: The output must be investor-ready with professional formatting.
4. **Sequential Flow**: Get topic first, then launch agent, then deliver results.

## Example Usage

```
/ai-strategy AI-native insurance brokerage
/ai-strategy Enterprise AI deployment competitive analysis
/ai-strategy AI-native accounting services vertical deep-dive
```
