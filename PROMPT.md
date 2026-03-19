# Exercise 1: Data Extraction

Pick one of the three files in this folder and ask your AI agent to extract
structured data from it.

## Option A — BFS Domestic Violence Statistics (bfs_domestic_violence_2009_2024.xlsx)

```
Extract and clean this BFS dataset into a usable flat CSV. The file has
16 yearly sheets with 3-row merged headers, nested categories (offence type
→ gender of accused → relationship type), "X" markers for suppressed counts,
and footnotes in the bottom rows.

I need a single CSV with columns: Year, Offence_Type, Gender_Accused,
Relationship_Type, Total_Offences, Cleared_Offences, Clearance_Rate.
Replace "X" with null. Flatten the hierarchy so each row is one combination.
```

## Option B — Court Filing (gauhati_high_court_wp4581_2025.pdf)

```
Extract all petitioner data from this 45-page court filing into a CSV.
Each entry has: number, name, relation (S/O or W/O + name), village, PO,
PS, district, PIN. There are 220 entries across ~15 pages.
```

## Option C — Scanned Government Contract (loa_excavators_contract.pdf)

```
Extract structured data from this scanned government contract. I need:
1. Contract metadata: issuing authority, contractor, date, tender ref, total value
2. Schedule of work table: item, quantity, unit, rate, amount
3. Distribution list of officials who received copies
```
