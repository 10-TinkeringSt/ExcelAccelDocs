# Spec 01: Solution skeleton and cube metadata extractor (XYZ.AQBE)

## Purpose
Create the solution skeleton, then build the metadata extractor. It connects to ONE SSAS Multidimensional cube, reads its STRUCTURAL metadata (dimensions, hierarchies, levels, measures) and saves it as one deterministic JSON file (the "metadata snapshot"). Later projects (MDX generator, DAX generator, Excel add-in) load that file as plain data. Read-only. No members, no data queries, no MDX, no UI.

## Rules for you
- Read AGENTS.md first. Windows corporate machine, no admin rights, dotnet CLI only. Target net10.0 (net9.0 only if net10.0 is unavailable; record it in DECISIONS.md).
- You may not be able to reach the SSAS server or the network. If a command needs the network (for example NuGet restore) and fails, STOP and tell me exactly what failed.
- Never invent NuGet package IDs, method signatures or rowset column names. Verify from package metadata or documentation. Anything you cannot verify goes behind one small class and is listed in DECISIONS.md as UNVERIFIED.
- Authentication is Windows Integrated. Never write passwords or connection strings to files, logs or console output, including error messages.
- Do not query $SYSTEM DMVs (they need server admin rights). Use the client library's schema-rowset API, the discovery mechanism Excel's field list relies on. I expect it works with ordinary read access; docs/VERIFY.md must tell me how to confirm that.
- SSAS client: ADOMD.NET from NuGet. I believe the modern-.NET package id is Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64 (verify it).

## Solution to create
Root: `XYZ.AQBE.sln`, `Directory.Build.props` (net10.0, nullable enabled, TreatWarningsAsErrors, Company XYZ, deterministic builds), `.gitignore` if missing, README.md, DECISIONS.md, docs/METADATA-SCHEMA.md, docs/VERIFY.md.
Projects (assembly name = namespace = project name; file-scoped namespaces, XML docs on public types, immutable records):
- `src/XYZ.AQBE.Model`: ZERO dependencies. Namespace `XYZ.AQBE.Model.Metadata`: the snapshot records, JSON load/save, and the snapshot integrity validator. Later specs add more types to this project.
- `src/XYZ.AQBE.Metadata`: IRowsetSource, AdomdRowsetSource, CsvRowsetSource, SnapshotBuilder, connection-string helper. References Model and the ADOMD package. This is the ONLY project that may reference ADOMD.
- `src/XYZ.AQBE.Cli`: console app, assembly name `aqbe`, command group `metadata`. Hand-rolled argument parsing (no CLI framework packages). Keep it thin; logic lives in the libraries.
- `tests/XYZ.AQBE.Model.Tests` and `tests/XYZ.AQBE.Metadata.Tests` (xUnit only).
Do NOT create other projects (Mdx, Sampling, Dax, App, ExcelAddin come from later specs); just list them in the README as planned.

## Architecture (the live connection is a thin, isolated layer)
- `IRowsetSource.GetRows(string rowsetName, IReadOnlyDictionary<string,string>? restrictions)` returns rows as case-insensitive column-name to value maps.
- `AdomdRowsetSource`: live implementation using the ADOMD schema-rowset API. The only class that touches ADOMD.
- `CsvRowsetSource`: reads `<dir>/<ROWSET_NAME>.csv` (or `.tsv`), header row = column names, RFC-4180 quoting, UTF-8 with or without BOM. It lets me work offline or from results exported from SSMS, and it is how you test.
- `SnapshotBuilder.Build(IRowsetSource, string cubeName)`: pure mapping logic, no I/O. It filters by CUBE_NAME itself, whatever the source did with restrictions.

## Rowsets to read
Verify each column name against the OLE DB for OLAP / MS-SSAS documentation. Treat every column as optional (a missing column maps to null) except the unique-name columns, which are required.
- MDSCHEMA_CUBES: CUBE_NAME, DESCRIPTION, BASE_CUBE_NAME, CUBE_SOURCE / CUBE_TYPE (whichever exists)
- MDSCHEMA_DIMENSIONS: DIMENSION_NAME, DIMENSION_UNIQUE_NAME, DIMENSION_CAPTION, DIMENSION_TYPE, DIMENSION_IS_VISIBLE
- MDSCHEMA_HIERARCHIES: HIERARCHY_NAME, HIERARCHY_UNIQUE_NAME, HIERARCHY_CAPTION, DIMENSION_UNIQUE_NAME, DEFAULT_MEMBER, ALL_MEMBER, HIERARCHY_IS_VISIBLE, HIERARCHY_DISPLAY_FOLDER, HIERARCHY_ORIGIN
- MDSCHEMA_LEVELS: LEVEL_NAME, LEVEL_UNIQUE_NAME, LEVEL_CAPTION, DIMENSION_UNIQUE_NAME, HIERARCHY_UNIQUE_NAME, LEVEL_NUMBER, LEVEL_TYPE, LEVEL_CARDINALITY, LEVEL_IS_VISIBLE
- MDSCHEMA_MEASURES: MEASURE_NAME, MEASURE_UNIQUE_NAME, MEASURE_CAPTION, MEASURE_AGGREGATOR, DATA_TYPE, MEASURE_IS_VISIBLE, MEASUREGROUP_NAME, MEASURE_DISPLAY_FOLDER, DEFAULT_FORMAT_STRING

## Mapping rules
- Cube selection: match the requested name case-insensitively among rows that are real cubes (not dimension-cubes, and not perspectives, which have a non-empty BASE_CUBE_NAME different from their own name; confirm this from the docs). If the name matches only a perspective, or nothing, fail with a message listing the real cube names.
- Dimensions: exclude the Measures dimension (unique name `[Measures]`, or measure-type). Its items go into `measures`.
- Keep hidden items with `isVisible=false`; never drop anything silently.
- `isAllLevel`: true for the hierarchy's All level (per the docs: LEVEL_TYPE for All, level number 0; verify and document).
- `defaultIsAll`: true only when allMember is not null and defaultMember equals allMember (ordinal, ignore case).
- Measures: store the raw `aggregatorCode` and a mapped `aggregator` name from the documented MEASURE_AGGREGATOR values (unknown codes map to "Unknown"). `isCalculated` is true for the documented calculated code (verify; I believe 127). For dimensions store `typeCode` and a mapped `type` the same way.
- No members anywhere. Level cardinality (a count) is fine.

## Snapshot JSON (camelCase, enums as strings, UTF-8 without BOM, "\n" line endings, indented, explicit nulls)
```json
{
  "schemaVersion": 1,
  "source": { "server": "SRV01", "database": "ERS_DB", "extractedAtUtc": "2026-10-05T10:00:00Z", "extractorVersion": "1.0.0" },
  "cube": {
    "name": "ERS_Cube", "caption": "ERS_Cube", "description": null,
    "dimensions": [
      {
        "uniqueName": "[Claim]", "name": "Claim", "caption": "Claim", "typeCode": 3, "type": "Other", "isVisible": true,
        "hierarchies": [
          {
            "uniqueName": "[Claim].[Claim Status]", "name": "Claim Status", "caption": "Claim Status",
            "isVisible": true, "displayFolder": null, "origin": null,
            "allMember": "[Claim].[Claim Status].[All]", "defaultMember": "[Claim].[Claim Status].[All]", "defaultIsAll": true,
            "levels": [
              { "uniqueName": "[Claim].[Claim Status].[(All)]", "name": "(All)", "caption": "(All)", "number": 0, "typeCode": 1, "isAllLevel": true, "cardinality": 1, "isVisible": true },
              { "uniqueName": "[Claim].[Claim Status].[Claim Status]", "name": "Claim Status", "caption": "Claim Status", "number": 1, "typeCode": 0, "isAllLevel": false, "cardinality": 12, "isVisible": true }
            ]
          }
        ]
      }
    ],
    "measures": [
      { "uniqueName": "[Measures].[Accounting Outstanding Losses]", "name": "Accounting Outstanding Losses", "caption": "Accounting Outstanding Losses",
        "measureGroup": "Claims", "displayFolder": null, "aggregatorCode": 1, "aggregator": "Sum", "isCalculated": false,
        "dataTypeCode": 5, "formatString": null, "isVisible": true }
    ]
  }
}
```
Determinism: sort dimensions, hierarchies and measures by uniqueName (ordinal); levels by number. The only run-dependent field is `source.extractedAtUtc`; `--no-timestamp` writes null. Document every field in docs/METADATA-SCHEMA.md; it is the contract for downstream projects.

## Snapshot integrity validator (in Model; run by `metadata validate` and automatically after `extract`)
Unique names unique (case-insensitive) within their kind; every hierarchy belongs to a listed dimension and every level to a listed hierarchy; level numbers unique within a hierarchy; `schemaVersion` supported; no empty unique names. Report ALL problems, not just the first.

## CLI
- `aqbe metadata extract --connection "<string>" --cube <name> --out <file> [--no-timestamp]`. The connection may come from env var `CUBE_CONNECTION` instead. Accept an Excel-style string with a leading `OLEDB;` prefix by stripping it. If ADOMD rejects a `Provider=` keyword, remove it and document that.
- `aqbe metadata extract --from-csv <dir> --cube <name> --out <file> [--no-timestamp]`
- `aqbe metadata validate <file>` and `aqbe metadata summary <file>` (counts of dimensions, hierarchies, levels, measures, hidden items, and hierarchies whose default is not All).
Only server and database/catalog go into `source`, parsed from the connection string. Exit codes: 0 ok, 1 failure, 2 usage error. Error messages must never echo the connection string.

## Tests (xUnit; no server needed)
- Synthetic CSV fixtures for a small cube: 3 dimensions (one with an attribute hierarchy, one with a user hierarchy and an All level), a hidden hierarchy, a hierarchy with no All member, a perspective row and a dimension-cube row in the cubes rowset (must be ignored), measures including a calculated one, plus one fixture with some optional columns missing.
- Golden snapshot JSON compared after normalizing line endings (written only when `UPDATE_GOLDEN=1`; otherwise a missing golden fails). Review every golden by eye.
- Shuffled CSV row order gives identical output.
- Load then save equals the original.
- Integrity validator: one test per rule, and several problems reported together.
- Connection string: server and catalog parsed; a password in the input never appears in any output, log or exception message.
- CSV parser: quotes, embedded commas and newlines, BOM, `.tsv`.
- The live AdomdRowsetSource is not unit tested; keep it as small as possible.

## Docs
README (usage and planned projects); DECISIONS.md (every non-obvious decision and everything UNVERIFIED); docs/METADATA-SCHEMA.md; docs/VERIFY.md: a checklist for me against the real cube: (1) run `extract` with my own Windows account, no admin; (2) compare dimension, hierarchy, level and measure counts with the same rowsets in SSMS; (3) spot-check three hierarchies' default members with an MDX DefaultMember query; (4) confirm perspectives aren't mistaken for cubes; (5) note the run time.

## Non-goals
No members, no data queries, no MDX, no UI, no reading the connection from Excel, nothing writable on the server.

## Process
Propose a short plan and wait for my approval. Then implement in this order: skeleton and Model records with JSON; integrity validator; CSV source; SnapshotBuilder; CLI; ADOMD source; docs. Run `dotnet build` and `dotnet test` after each step; finish with no failures or warnings. Decide small things yourself and log them. End with a short report: what was built, test counts, UNVERIFIED items, open questions.