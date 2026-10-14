# NovaRetail Executive

## Overview

NovaRetail Executive Deck Generator is a Google ADK-powered AI agent that automates quarterly business review reporting.

The solution analyzes business data from Excel, generates AI-driven executive insights using Gemini, creates supporting visualizations, and automatically builds a five-slide PowerPoint presentation covering:

- Regional Performance
- Customer Experience
- Operational Risk
- Executive Recommendations

The workflow combines Google ADK, Gemini 3.5 Flash, Pandas, Matplotlib, and Python-PPTX to transform raw business data into an executive-ready presentation with minimal manual effort.

## Architecture

User
↓
ADK Agent
↓
analyze_data()
↓
Gemini Insight Generation
↓
generate_charts()
↓
build_presentation()
↓
PowerPoint Output


## Tools

### analyze_data()
Reads the NovaRetail dataset and generates business metrics.

### generate_charts()
Creates performance, customer experience, and operational risk charts.

### build_presentation()
Populates the PowerPoint template with AI-generated content, charts, and icons.


## How to Run

```bash
adk web --port 8000

## Sample Prompt

The agent can be triggered using a simple natural language request:

```text
Create the NovaRetail Q2 2026 executive presentation.
```

The agent will automatically:

1. Analyze the NovaRetail dataset
2. Generate executive insights using Gemini
3. Create supporting charts
4. Populate the PowerPoint template
5. Insert icons and visual assets
6. Generate the final presentation deck