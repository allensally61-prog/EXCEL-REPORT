# Visio schematic updater

Pushes values from the LT-WATSAN project summary Excel into the labels of
`LT-WATSAN SCHEMATIC DIAGRAM_rev03.vsdx`. Runs on Windows (the pipeline-summary
tables are drawn by `render_table.ps1` through PowerShell). Needs `pip install openpyxl`.

## Update only Option 04

```
python update_schematic.py --excel "<OPT01..OPT05_2030_2.xlsx>" --vsdx "<rev03.vsdx>" --pages "OPT 4A,OPT 4B"
```

Output: `..._UPDATED.vsdx` plus `..._UPDATED_CHANGES.csv` (every old -> new value with its Excel cell).

- **OPT 4B** (Katavi and Rukwa common intake) uses the same shapes as OPT 2b.
- **OPT 4A** (Kigoma, Malagarasi intake): Kasulu BPS -> Heru juu -> Kitambuka/Kibondo
  uses the same shapes as OPT 1A. The upstream part (Malagarasi intake, Mpeta WTP,
  Chakulu BPS, Masanza, Mutinde BPS, Rusesa, Lolawaha) is new, so those entries in
  `mapping.json` have keys starting with `?` and are skipped until you fill in their shape IDs:

  ```
  python update_schematic.py --vsdx "<rev03.vsdx>" --dump "OPT 4A"
  ```

  prints every shape ID with its text; replace each `?...` key with the matching ID.
  For Mutinde BPS, if both pump blocks sit in one label use `ID#0` and `ID#1`.

If your page names differ from `OPT 4A` / `OPT 4B`, rename the keys in `mapping.json`
(the CSV lists the drawing's real page names when a page is not found).
