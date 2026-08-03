# 4d-plugin-xlnt-csv

`xlsx-to-csv` is a 4D plugin that converts Excel workbooks (`.xlsx`) into CSV files — one file per sheet — and can list the sheet names in a workbook, without going through 4D's own import commands or opening Excel. It's built on [xlsxio](https://github.com/brechtsanders/xlsxio), a streaming C reader for the `.xlsx` format, chosen specifically to avoid the memory growth the plugin's original implementation (built on the `xlnt` library) suffered on very large sheets on Windows. The plugin exposes two commands: [`convert xlsx to csv`](#convert-xlsx-to-csv) and [`get xlsx sheets`](#get-xlsx-sheets).

| Command | Returns | Purpose |
|---|---|---|
| [`convert xlsx to csv`](#convert-xlsx-to-csv) | — | Converts one or more sheets of a workbook into CSV files on disk. |
| [`get xlsx sheets`](#get-xlsx-sheets) | — | Returns the list of sheet names contained in a workbook. |

**Platforms:** macOS (Universal — arm64 + x86_64) · Windows (x64)

---

## Requirements & platform notes

- Both commands are declared `threadSafe: true` in the plugin manifest and can be called from a preemptive process.
- Neither command returns a value — both are declared in the plugin manifest without a return type, i.e. they behave as procedures, not functions. There is no function result to check for success.
- **Password-protected workbooks are not currently supported**, even though both commands accept a `password` parameter — see the note in each command's parameter table and in Error handling below.
- Failure is silent, not a 4D error. A bad path, a corrupt file, or a workbook the plugin can't open (including any password-protected file, since passwords aren't functional) causes the command to simply do nothing — no CSV files are written, `get xlsx sheets`'s array comes back empty, and 4D raises nothing. Always check for the expected output (files on disk, a non-empty array) rather than relying on error trapping.
- `convert xlsx to csv` accepts an `encoding` parameter controlling output character encoding and BOM (UTF‑8 with BOM by default, UTF‑8 without BOM, or Shift‑JIS). `get xlsx sheets` has no encoding option — sheet names always come back as plain 4D text.

---

## convert xlsx to csv

### Syntax
```
convert xlsx to csv ( path ; sheets ; password ; folder ; fieldDelimiter ; recordDelimiter ; encoding )
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Full path to the source `.xlsx` file. |
| `sheets` | Array Text (in) | Names of the sheets to export. Leave every element empty (or pass a zero-length array) to export **every** sheet in the workbook — see Description below for what "empty" means exactly. |
| `password` | Text | **Ignored in this build** — the value is read but never actually used to open the workbook. Pass an empty string; do not rely on this parameter to open password-protected files. |
| `folder` | Text | Destination folder for the generated CSV files. Created automatically — including any missing parent folders — if it doesn't already exist. |
| `fieldDelimiter` | Text | Field (column) separator written between cell values. Pass an empty string to use the default, a comma (`,`). |
| `recordDelimiter` | Text | Record (row) separator written at the end of each row. Pass an empty string to use the default, `\r\n`. |
| `encoding` | Longint | `0` (or any value other than `1`/`2`) — UTF‑8 **with** a byte-order mark (default). `1` — UTF‑8 **without** a BOM. `2` — Shift‑JIS, no BOM. |
| Result | — | This command does not return a value. |

### Description
One CSV file is written per exported sheet, named `<sheet name>.csv`, into `folder`. If `folder` doesn't end in a path separator, one is added automatically before the sheet files are written.

If `sheets` is left empty, the plugin discovers the workbook's sheet names itself (via `xlsxio`'s sheet-list API) and converts all of them — but that discovered list is **not** written back into your `sheets` array. If you need the actual sheet names in 4D (for a log, a report of what was exported, etc.), call [`get xlsx sheets`](#get-xlsx-sheets) separately; don't expect `convert xlsx to csv` to populate `sheets` for you. The trigger for "convert everything" is specifically *no element of `sheets` has any length* — an array containing one empty-string element behaves the same as a zero-length array.

`encoding = 2` (Shift‑JIS) only re-encodes **cell values** — `fieldDelimiter` and `recordDelimiter` are written out exactly as given, with no re-encoding. Keep custom delimiters to plain ASCII characters when using Shift‑JIS output; a non-ASCII delimiter would end up mixed-encoding in the same file as the Shift‑JIS cell data.

The conversion runs synchronously — there's no progress callback or cancellation exposed to 4D — but it periodically yields control back to 4D while iterating rows, so it doesn't freeze the interface on very large sheets the way a tight, non-yielding loop would.

### Example
From the plugin's own test method (`TEST.4dm`):
```4d
//%attributes = {}
C_TEXT:C284($src)
ARRAY TEXT:C222($sheets;0)
C_TEXT:C284($password)

$src:=System folder:C487(Desktop:K41:16)+"SB_data.xlsx"

  //get xlsx sheets ($src;$sheets;$password)

$dst:=System folder:C487(Desktop:K41:16)+"SB_data"

convert xlsx to csv ($src;$sheets;$password;$dst;"";"";2)
```
This converts every sheet in `SB_data.xlsx` (since `$sheets` is left at size 0) to Shift‑JIS CSV files, using the default `,`/`\r\n` delimiters, into a `SB_data` folder on the Desktop.

Exporting only specific sheets, with a semicolon delimiter and UTF‑8 without a BOM:
```4d
C_TEXT($path;$folder;$password)
ARRAY TEXT($sheets;2)

$sheets{1}:="Summary"
$sheets{2}:="Detail"

$path:=System folder:C487(Desktop:K41:16)+"invoice.xlsx"
$folder:=System folder:C487(Desktop:K41:16)+"invoice_csv"
$password:=""

convert xlsx to csv ($path;$sheets;$password;$folder;";";"";1)
```

---

## get xlsx sheets

### Syntax
```
get xlsx sheets ( path ; sheets ; password )
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Full path to the source `.xlsx` file. |
| `sheets` | Array Text (out) | Filled with the workbook's sheet names, in the order the file stores them. Any prior contents of the array are discarded. |
| `password` | Text | **Ignored in this build** — same caveat as `convert xlsx to csv`'s `password` parameter. Pass an empty string. |
| Result | — | This command does not return a value. |

### Description
Reads the workbook's sheet list via `xlsxio` and returns the names into `sheets`. If `path` can't be opened (missing file, corrupt file, or a password-protected file, since passwords aren't supported), `sheets` simply comes back empty — no error is raised.

### Example
```4d
C_TEXT($path;$password;$sheetName)
ARRAY TEXT($sheets;0)

$path:=System folder:C487(Desktop:K41:16)+"invoice.xlsx"
$password:=""

get xlsx sheets ($path;$sheets;$password)

For each ($sheetName;$sheets)
	ALERT:C41("Sheet found: "+$sheetName)
End for each 
```

Using the discovered list to drive a full export:
```4d
C_TEXT($path;$folder;$password)
ARRAY TEXT($sheets;0)

$path:=System folder:C487(Desktop:K41:16)+"invoice.xlsx"
$folder:=System folder:C487(Desktop:K41:16)+"invoice_csv"
$password:=""

get xlsx sheets ($path;$sheets;$password)
If (Size of array:C274($sheets)>0)
	convert xlsx to csv ($path;$sheets;$password;$folder;"";"";0)
End if 
```

---

## Error handling & troubleshooting

- **A workbook that can't be opened fails silently, not with a 4D error.** Bad path, corrupt file, or unsupported/password-protected file all result in the command simply doing nothing — no CSV files, an empty `sheets` array, nothing raised. Check for the files/array contents you expect rather than trapping an error.
- **Password-protected workbooks aren't actually supported, despite the `password` parameter existing on both commands.** The value is read internally but never passed through to the underlying file-open call in this build. Always pass an empty string, and don't expect encrypted files to open.
- **`convert xlsx to csv` never writes the discovered sheet list back to you.** Leaving `sheets` empty converts every sheet, but your `sheets` array stays exactly as you passed it in — call `get xlsx sheets` first if you need the names in 4D.
- **"Convert everything" is triggered by every sheet name being empty, not by array size alone.** An array of size 1 holding a single empty-string element is treated the same as a zero-length array — both mean "all sheets."
- **Shift‑JIS output (`encoding = 2`) only converts cell values.** Your `fieldDelimiter`/`recordDelimiter` are written verbatim; stick to ASCII delimiters when using Shift‑JIS to avoid a mixed-encoding file.
- **No progress or cancel callback.** The call blocks until the conversion finishes (it yields periodically internally so 4D doesn't freeze, but there's nothing to hook into from 4D for progress reporting).

---

## Quick reference

```4d
 // list sheets
C_TEXT($path;$folder;$password)
ARRAY TEXT($sheets;0)
$path:=System folder:C487(Desktop:K41:16)+"invoice.xlsx"
$password:=""
get xlsx sheets ($path;$sheets;$password)

 // convert everything, defaults (UTF-8 + BOM, comma / CRLF)
$folder:=System folder:C487(Desktop:K41:16)+"invoice_csv"
convert xlsx to csv ($path;$sheets;$password;$folder;"";"";0)

 // convert two named sheets, UTF-8 no BOM, semicolon-delimited
ARRAY TEXT($sheets;2)
$sheets{1}:="Summary"
$sheets{2}:="Detail"
convert xlsx to csv ($path;$sheets;$password;$folder;";";"";1)
```
