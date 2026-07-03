# thesis-format-dxat

QoderWork Skill: Automatically format graduation theses per DXAT (Dianxi Applied Technology University) template standards.

## Files

- `SKILL.md` - Main skill documentation with workflow, style mapping, and formatting rules
- `ooxml-reference.md` - OOXML technical reference (namespaces, section breaks, header/footer XML templates)
- `pitfalls-and-fixes.md` - Known issues and fixes (hardcoded style IDs, section break positioning, etc.)
- `scripts/merge_thesis.py` - Complete merge/formatting script
- `thesis-format-dxat.zip` - Packaged skill archive (installable in QoderWork)

## Usage

Install this skill in QoderWork, then submit a thesis draft (.docx) along with a reference example document. The skill handles all OOXML-level formatting including section breaks, headers/footers, TOC field codes, cover page titles, and page numbering.

## License

MIT