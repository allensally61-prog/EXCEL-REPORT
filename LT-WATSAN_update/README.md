# Visio schematic updater

Pushes values from the LT-WATSAN project summary Excel into the labels of
`LT-WATSAN SCHEMATIC DIAGRAM_rev03.vsdx`. Runs on Windows (the pipeline-summary
tables are drawn by `render_table.ps1` through PowerShell). Needs `pip install openpyxl`.

## Option 04 (done)

`LT-WATSAN SCHEMATIC DIAGRAM_rev03_UPDATED_OPT4.vsdx` = the `_UPDATED.vsdx` with pages
OPT 4A and OPT 4B also updated from the OPTION 04 sheet; every change is listed in
`..._UPDATED_OPT4_CHANGES.csv`. To redo it:

```
python update_schematic.py --pages "OPT 4A,OPT 4B" --excel "<..._2030_2.xlsx>" ^
  --vsdx "LT-WATSAN SCHEMATIC DIAGRAM_rev03_UPDATED.vsdx" --out "LT-WATSAN SCHEMATIC DIAGRAM_rev03_UPDATED_OPT4.vsdx"
```

Always give `--out` a new name - never the same file as `--vsdx`.

- `--pages` limits the run to the listed pages; `--dump "<page>"` lists a page's shape IDs and text.
- Tables: on Windows `render_table.ps1` draws them; elsewhere `render_table.py` draws the
  same table (needs `pip install fonttools` and the Carlito font).
