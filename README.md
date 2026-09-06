# Fuzzy Lookup 2026 Add-In for Excel

A rebuilt, 64-bit compatible replacement for the original Microsoft Fuzzy Lookup Add-In that stopped working in modern Excel versions.

Performs fuzzy matching of textual data between two Excel tables using multiple string similarity algorithms.

## Features

- **Levenshtein Distance** — edit-based similarity
- **Jaro-Winkler** — prefix-weighted character matching
- **Jaccard** — token overlap for multi-word strings
- **Auto mode** — weighted combination of all algorithms
- Configurable similarity threshold (0–100%)
- Multiple matches per row
- Text cleaning (trim, lowercase, remove punctuation)
- Custom find/replace transformations
- Worksheet functions: `=FUZZYMATCH(A1, B1)` and `=FUZZYVLOOKUP(A1, range, 1, 2, 0.6)`

## Download

Download **[FuzzyLookup.xlam](FuzzyLookup.xlam)** from this repository.

## Installation

### Step 1: Unblock the file (Windows security)

Windows blocks files downloaded from the internet. You must unblock it first:

1. Right-click **FuzzyLookup.xlam** in File Explorer
2. Click **Properties**
3. At the bottom of the General tab, check **Unblock**
4. Click **OK**

> If you skip this step, Excel will silently refuse to load the add-in or show a security error.

### Step 2: Install in Excel

1. Open **Excel**
2. Go to **File → Options → Add-ins**
3. At the bottom, next to "Manage:", select **Excel Add-ins** and click **Go...**
4. Click **Browse...**
5. Navigate to where you saved **FuzzyLookup.xlam** and select it
6. Make sure the checkbox next to **FuzzyLookup** is checked
7. Click **OK**

### Step 3: Enable macros (if prompted)

- If Excel asks about macros, click **Enable Macros**
- If the add-in doesn't load, go to **File → Options → Trust Center → Trust Center Settings → Macro Settings** and select **Enable VBA macros**

## Usage

### Task Pane

1. Click **Fuzzy Lookup → Open Fuzzy Lookup...** in the menu bar
2. Select your **Left Table** (the values to look up)
3. Select your **Right Table** (the table to match against)
4. Choose the **Match Column** for each table
5. Adjust the **Similarity Threshold** (lower = more matches, higher = stricter)
6. Click **Go!**

Results appear in a new sheet with all columns from both tables plus a similarity score.

### Worksheet Functions

Use these directly in cells:

```
=FUZZYMATCH("Microsoft Corp", "Microsoft Corporation")
→ 0.82

=FUZZYVLOOKUP("Jon Smith", A1:B100, 1, 2, 0.6)
→ Returns column 2 value of the best match in column 1 above 60% similarity
```

## Compatibility

- Excel 2016, 2019, 2021, 2024 (32-bit and 64-bit)
- Microsoft 365
- Windows only (uses VBA)

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Add-in doesn't appear in menu | Unblock the file (Step 1) and restart Excel |
| "Macros have been disabled" | Enable macros in Trust Center settings |
| Columns don't populate | Close Excel, reinstall the add-in |
| Slow on large datasets | Raise the similarity threshold to reduce comparisons |

## License

MIT License — free to use, modify, and distribute.
