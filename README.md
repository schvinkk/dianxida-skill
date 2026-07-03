# thesis-format-dxat

A universal AI agent skill for automatically formatting graduation theses per Dianxi Applied Technology University (DXAT) template standards.

## Overview

This skill can be installed and used by **any AI agent** that supports Markdown-based skill definitions (QoderWork, Claude, GPT, Cursor, Copilot, or any custom agent framework). It is not limited to any specific platform.

The skill provides a complete, step-by-step workflow for formatting .docx thesis documents at the OOXML level, ensuring pixel-perfect compliance with the university's official template.

## What It Does

- Formats thesis drafts to match DXAT's official template (fonts, spacing, margins, page setup)
- Handles OOXML-level operations: section breaks, headers/footers, TOC field codes, cover page, page numbering
- Dynamically resolves style IDs (never hardcodes them) to work with any document
- Preserves the user's original text content without any modification
- Includes AI detection rate checking and content improvement suggestions

## Files

| File | Description |
|------|-------------|
| `SKILL.md` | Main skill documentation 鈥?workflow, style mapping, formatting rules, iron laws |
| `ooxml-reference.md` | OOXML technical reference 鈥?XML namespaces, section break positioning, templates |
| `pitfalls-and-fixes.md` | Known issues and solutions 鈥?style ID mismatches, run-level overrides, TOC quirks |
| `scripts/merge_thesis.py` | Complete Python merge/formatting script |
| `thesis-format-dxat.zip` | Packaged skill archive for easy installation |

## How to Use

1. **Provide a thesis draft** (.docx) that needs formatting
2. **Provide a reference example** (.docx) that already has correct formatting 鈥?the skill uses this as the baseline
3. The agent reads `SKILL.md` and follows the workflow to format the draft
4. Output is a properly formatted .docx file compliant with DXAT standards

## Requirements

- Any AI agent with file read/write capabilities
- Python 3.x (for running the merge script)
- A correctly-formatted thesis example (as the format baseline)
- The official DXAT template document (included in references)

## Key Principles

- **Copy, don't create**: Always use the example document as the base, copy its XML structure
- **Dynamic style IDs**: Never hardcode style IDs 鈥?they vary across documents
- **Content preservation**: Never modify the user's text content, only formatting
- **Run-level awareness**: Check for run-level overrides that differ from style definitions

## License

MIT