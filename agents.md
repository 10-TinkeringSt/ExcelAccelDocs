# XYZ.AQBE (Accelerator for Query Building in Excel)

Builds MDX (later DAX) queries from a user's arrangement of Filters, Dimensions and Measures, using cube metadata. End goal: an Excel-DNA add-in. Libraries stay Excel-free; only XYZ.AQBE.ExcelAddin may reference Excel or Excel-DNA.

## Constraints
- Windows corporate machine, no admin rights, no Visual Studio. dotnet CLI only. Target net10.0.
- Projects: src/XYZ.AQBE.<Name>; tests in tests/XYZ.AQBE.<Name>.Tests; namespace = project name.
- Dependencies point inward: ExcelAddin -> App -> Mdx / Dax / Metadata -> Model. Model has zero dependencies. Only Metadata references ADOMD.
- Never invent package IDs, API signatures or rowset column names. Mark anything unverifiable as UNVERIFIED in DECISIONS.md.
- Never write credentials or connection strings to files, logs or output.
- Generated output must be deterministic ("\n" line endings, no timestamps).

## Commands
- Build: dotnet build
- Test: dotnet test --nologo -v q

## Working agreements
- Specs live in docs/prompts/. Read the one named in the request and do not expand its scope.
- Propose a plan and wait for approval before editing.
- Log non-obvious decisions in DECISIONS.md.
- Finish each step with build and tests passing and no warnings.