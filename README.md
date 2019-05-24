# Kicad_bom_sync
KiCad BOM plugin that syncs changes from KiCad to a spreadsheet.

## Concept
Most BOM tools simply export to CSV or XLSX. You would then manually edit the spreadsheet
(adding component price info, manufacturers, style, etc). But what if you then make changes to the design?

This tool can edit an existing XLSX spreadsheet to update it with changes from KiCad.
By highligting the changes, it is essentially a diff tool for the BOM.


# How does it work

## Calling the script

This works the same as most other BOM plugins. The script expects two parameters: the KiCad Bom file (XML format)
and the name of the output file:
```
python "pathToFile/BOM.py" "%I" "%O
```

## Creating a new BOM
If the specified xlsx file does not exist, a new file will be created. If it does, it is updated as explained below.

## Syncing to an existing BOM
The script looks for a sheet called 'BOM' and assumes the first row are headers (such as 'Ref', 'Qty' etc) that are matched with the component fields from KiCad.
This means the script will work even if you rearrange the order of the columns.

Components are grouped by value and footprint:
* Matching component found? Any non-empty fields from KiCad overwrite the values in the sheet and are marked yellow.
* No matching component (matching value + footprint) found? A new row is added and marked green.
* If the sheet has a column called 'Sync' (optional but highly recommended),
this is updated to '1' for matched components, or marked red for obsolete.
