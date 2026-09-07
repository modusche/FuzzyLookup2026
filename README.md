# Fuzzy Lookup 2026 — Free Excel Add-In

> **The original Microsoft Fuzzy Lookup Add-In is discontinued and doesn't work on 64-bit Excel.**
> This is a free, open-source replacement that works on all modern versions of Excel.

Fuzzy match names, addresses, company names, or any text between two Excel tables — even when the data has typos, abbreviations, or inconsistent formatting.

## Why This Exists

Microsoft's official Fuzzy Lookup Add-In (released ~2012) stopped working years ago on 64-bit Excel. There's no official fix. This project is a clean-room rebuild that:

- Works on **64-bit Excel** (2016, 2019, 2021, 2024, Microsoft 365)
- Runs **10-100x faster** than the original (bulk array processing, optimized algorithms)
- Includes a **web version** that also works on Mac and Excel Online
- Is **completely free** and open source

## Algorithms

| Algorithm | Best For | How It Works |
|-----------|----------|--------------|
| **Levenshtein** | Typos, misspellings | Counts character edits (insert, delete, replace) |
| **Jaro-Winkler** | Names, short strings | Rewards matching characters + shared prefix |
| **Jaccard** | Multi-word strings | Measures word overlap between two strings |
| **Combined** (default) | General use | Weighted blend of all three |

## Quick Start

### Option A: VBA Add-In (.xlam) — Windows

1. Download **[FuzzyLookup.xlam](FuzzyLookup.xlam)**
2. **Unblock it:** Right-click → Properties → check ✅ **Unblock** → OK
3. In Excel: **File → Options → Add-ins → Excel Add-ins → Go → Browse** → select the file
4. Click **Fuzzy Lookup → Open Fuzzy Lookup** in the menu bar

### Option B: Web Add-In — Windows, Mac & Excel Online

1. Download **[manifest.xml](web-addin/manifest.xml)**
2. In Excel: **Insert → Get Add-ins → Upload My Add-in** → select manifest.xml
3. Click **Fuzzy Lookup** on the Home tab

No macros, no unblocking, no VBA — works everywhere.

## How to Use

1. Put your data in **two sheets** (or Excel Tables)
2. Open the Fuzzy Lookup task pane
3. Select the **Left Table** (your messy data) and **Right Table** (your clean reference)
4. Pick the **Match Column** in each table
5. Set the **Similarity Threshold** (default 0.65 — lower = more matches, higher = stricter)
6. Click **Go!**

Results appear in a new sheet with all columns from both tables + a similarity score.

### Worksheet Functions

Use directly in cells — no task pane needed:

```excel
=FUZZYMATCH("Microsoft Corp", "Microsoft Corporation")
→ 0.82

=FUZZYVLOOKUP("Jon Smith", A1:B100, 1, 2, 0.6)
→ Returns best fuzzy match from column 1, returns column 2 value
```

## Use Cases

- Matching **customer lists** from different systems
- Deduplicating **contact databases**
- Reconciling **company names** (Inc vs Incorporated, Corp vs Corporation)
- Linking **product catalogs** with inconsistent naming
- Matching **addresses** with typos or abbreviations
- Any VLOOKUP that fails because the data isn't an exact match

## Compatibility

| Version | VBA (.xlam) | Web Add-In |
|---------|:-----------:|:----------:|
| Excel 2016 | ✅ | ✅ |
| Excel 2019 | ✅ | ✅ |
| Excel 2021 | ✅ | ✅ |
| Excel 2024 | ✅ | ✅ |
| Microsoft 365 | ✅ | ✅ |
| 32-bit Excel | ✅ | ✅ |
| 64-bit Excel | ✅ | ✅ |
| Mac | ❌ | ✅ |
| Excel Online | ❌ | ✅ |

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Add-in doesn't appear | Right-click .xlam → Properties → Unblock, then restart Excel |
| "Macros disabled" | File → Options → Trust Center → Enable VBA macros |
| Columns don't populate | Close Excel completely, reopen, reinstall add-in |
| Slow on large datasets | Raise threshold to 0.8+, reduce max matches to 1 |
| Web add-in error | Make sure you downloaded manifest.xml, not the .xlam |

## Keywords

fuzzy lookup, fuzzy match, fuzzy vlookup, approximate match, string matching, Excel add-in, data matching, name matching, deduplication, record linkage, Levenshtein distance, Jaro-Winkler, edit distance, Microsoft Fuzzy Lookup replacement, 64-bit Excel add-in

## License

MIT License — free to use, modify, and distribute.
