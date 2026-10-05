# Spec 02: MDX generator and spec sampler (XYZ.AQBE)

## Purpose
Add to the existing solution: a validated MDX generator driven by two JSON inputs. (1) a QuerySpec JSON: the user's arrangement of Filters with typed member values, Dimensions and Measures. (2) the metadata snapshot JSON produced by Spec 01. Output: formatted MDX that returns a TIDY FLAT table (measures as columns, one row per dimension combination). The MDX is later pasted into Power Query / Power Pivot. There is no rows/columns choice for users.
Out of scope: UI, metadata extraction, server connections, running queries, DAX.

## Rules for you
- Read AGENTS.md and docs/prompts/01-metadata-extractor.md first. Spec 01 is already built in this repo. REUSE the metadata records in `XYZ.AQBE.Model.Metadata`; never create a second copy. If `samples/metadata/metadata.json` differs from them, extend the Model additively (existing tests must still pass) and list the differences in DECISIONS.md.
- Corporate Windows machine, dotnet CLI, net10.0. Mdx and Sampling have no runtime dependencies beyond Model.
- You cannot reach any cube. Never claim the MDX works on the server; only that it is well-formed and matches the golden files.
- Never invent APIs or package details. Log anything unverifiable as UNVERIFIED in DECISIONS.md.
- Deterministic output: "\n" line endings, no timestamps.

## Projects to add
- Extend `src/XYZ.AQBE.Model`: namespace `XYZ.AQBE.Model.Query` (QuerySpec, MeasureRef, DimensionRef, FilterSpec, FilterKind, MemberMode; System.Text.Json with case-insensitive property names and enums as strings) and `UniqueNameTokenizer` (splits unique names into segments, respecting `]]` escapes).
- `src/XYZ.AQBE.Mdx`: references Model. MetadataIndex, validator, generator, MdxOptions, ValidationIssue.
- `src/XYZ.AQBE.Sampling`: references Model. `SpecSampler`, a dev/test tool used by the CLI and the tests, never by the add-in.
- Extend `src/XYZ.AQBE.Cli`: command group `mdx`.
- `tests/XYZ.AQBE.Mdx.Tests` (references Mdx, Sampling, Model; xUnit only). Add tokenizer and QuerySpec JSON tests to `tests/XYZ.AQBE.Model.Tests`.
- Also add `samples/specs/`, `docs/VALIDATION.md` (every issue code), and extend README, DECISIONS.md and docs/VERIFY.md.
Dependencies point inward only: Mdx and Sampling reference Model; nothing references Cli.

## Input 1: QuerySpec
- `QuerySpec { Cube; Measures: [MeasureRef{Name}]; Dimensions: [DimensionRef{Name}]; Filters: [FilterSpec{Hierarchy, Kind, Mode, Values[]}] }`
- Measure name = full unique name, e.g. `[Measures].[Accounting Outstanding Losses]`.
- Dimension name = full unique LEVEL name, e.g. `[Reporting Structure].[Line Of Business].[Line Of Business]`.
- Filter `Hierarchy` = full unique HIERARCHY name, e.g. `[Source].[Source Name]`. `Kind`: Single | List | Range. `Mode`: Key (`&[value]`) | Name (`[value]`) | UniqueName (value used verbatim).
- The caller supplies full names. The generator only looks things up in the metadata to validate and canonicalize.

## Input 2: metadata snapshot
Load it with the existing Model types. Build a `MetadataIndex` with case-insensitive lookups (SSAS names are case-insensitive): measure by unique name; level by unique name; hierarchy by unique name; level to parent hierarchy. It must also report when a name exists but is the wrong kind (for example a hierarchy name given as a dimension) so error messages can be precise.

## Public API (project XYZ.AQBE.Mdx)
```csharp
public static class MdxGenerator
{
    public static IReadOnlyList<ValidationIssue> Validate(QuerySpec spec, MetadataSnapshot? metadata, MdxOptions? options = null);
    public static string Generate(QuerySpec spec, MetadataSnapshot? metadata, MdxOptions? options = null); // throws MdxValidationException carrying all Errors
}
public sealed record MdxOptions(bool NonEmptyRows = true, FilterPlacement FilterPlacement = FilterPlacement.Hybrid,
    CrossJoinStyle CrossJoinStyle = CrossJoinStyle.Operator, bool IncludeHeaderComment = false, int IndentSize = 4,
    long LargeResultThreshold = 1_000_000);
```
`metadata` null means spec-only checks (no existence checks, no canonicalization). `ValidationIssue { Severity (Error|Warning), Code, Message, Path }`. Report ALL issues. Use stable codes (`S###` spec-only, `M###` metadata) and document them in docs/VALIDATION.md.

## Generation rules
1. Measures: `{ m1, m2, ... } ON COLUMNS`.
2. Dimensions: each emitted as `<level unique name>.Members`. One dimension: used directly. Several: crossjoined in the order given. `CrossJoinStyle.Operator` joins with `*` inside braces. `CrossJoinStyle.Function` emits `CrossJoin(a, b)` and nests for three or more: `CrossJoin(CrossJoin(a, b), c)`. Zero dimensions: omit `ON ROWS` (one row of totals). `NON EMPTY` precedes the ROWS set when `NonEmptyRows`. Never emit HIERARCHIZE, DRILLDOWNLEVEL, DIMENSION PROPERTIES or CELL PROPERTIES, and never include the All member row (deliberate departures from the MDX Excel generates, because we need a flat table; record this in DECISIONS.md).
3. Filter placement `Hybrid` (default), mirroring MDX captured from Excel (a single selected item goes in the slicer; several items go in a subselect):
   - `Single` goes in the WHERE tuple, in input order, UNLESS its hierarchy is also the hierarchy of any dimension, in which case it goes in a subselect as a one-member set (a slicer cannot share a hierarchy with an axis).
   - `List` and `Range` always go in subselects. `List` keeps the given order (no sorting). `Range` takes exactly two values and emits `{ first : last }`, inclusive, never reordered.
   - Subselect filters keep input order: first = outermost, innermost selects `FROM [Cube]`. With no subselect filters emit `FROM [Cube]` directly; with no WHERE filters emit no WHERE clause.
   - `AllSubselect`: every filter, Single included, goes in a subselect; nothing in WHERE.
   - WHERE format: `WHERE ( m1, m2 )` on one line.
4. Hierarchy overlap: with metadata, a dimension's hierarchy is the parent hierarchy of its level, compared by canonical unique name. Without metadata, fall back to the first two name segments (document this limitation). Use the tokenizer, never string splitting.
5. Member references: Key `<Hierarchy>.&[<value>]`; Name `<Hierarchy>.[<value>]`; UniqueName verbatim and untouched. In Key and Name, escape `]` as `]]`.
6. Canonicalization: with metadata, emit canonical unique names from the metadata (so letter-case differences in the input never reach the MDX). Typed member VALUES are never altered.
7. Formatting: uppercase keywords, nested subselects indented by `IndentSize`, one clause per line, deterministic. The layout below is the starting point; document the final layout in DECISIONS.md and lock it with golden tests. `IncludeHeaderComment` (default false; I have not verified that Power Query tolerates comments) adds a short comment noting that measures are pre-aggregated at the queried grain and non-additive measures will not re-aggregate correctly downstream.

## Validation
Spec-only Errors: cube empty; zero measures; duplicate measure or dimension (case-insensitive); a name not starting with `[`; a filter with no values; Single not exactly 1 value; Range not exactly 2; an empty, whitespace-only or newline-containing value; the same hierarchy in two different filters.
Spec-only Warnings: List with duplicate values; dimension name with fewer than 3 segments (looks like a hierarchy, not a level).
Metadata Errors: cube name differs from the metadata cube; measure not found; dimension not found as a level (if the name is a hierarchy or dimension, say so and suggest its level names); filter hierarchy not found (if the name is a level, say to use its hierarchy); `[Measures]` used as a filter hierarchy.
Metadata Warnings: a referenced item is hidden; a dimension is an All level; two dimensions come from the same hierarchy (I believe SSAS rejects this in a crossjoin but have not confirmed it, so keep it a Warning and add the check to VERIFY.md); a measure is calculated or its aggregator is not additive (anything other than Sum, Count, Min or Max; document the mapping); the product of dimension level cardinalities exceeds `LargeResultThreshold` (state clearly that this is an upper bound before filters and NON EMPTY).

## Acceptance examples
Example A (self-contained; build a small synthetic metadata fixture for a cube named `Sales` that contains these names). Input: measures `[Measures].[Sales Amount]`, `[Measures].[Order Quantity]`; dimensions `[Date].[Calendar].[Calendar Year]`, `[Product].[Category].[Category]`; filters: `[Customer].[Country]` List Key [Australia, Canada]; `[Date].[Calendar Date]` Range Key [20260101, 20260131]; `[Sales Territory].[Region]` Single Key [Pacific]. Expected output (Hybrid, Operator):
```mdx
SELECT
    { [Measures].[Sales Amount], [Measures].[Order Quantity] } ON COLUMNS,
    NON EMPTY
        { [Date].[Calendar].[Calendar Year].Members
          * [Product].[Category].[Category].Members } ON ROWS
FROM
(
    SELECT { [Customer].[Country].&[Australia], [Customer].[Country].&[Canada] } ON COLUMNS
    FROM
    (
        SELECT { [Date].[Calendar Date].&[20260101] : [Date].[Calendar Date].&[20260131] } ON COLUMNS
        FROM [Sales]
    )
)
WHERE ( [Sales Territory].[Region].&[Pacific] )
```
Example B (illustrative; use real names from samples/metadata/metadata.json and note any that do not exist): measure `[Measures].[Accounting Outstanding Losses]`; dimensions the Line Of Business and Claim Status levels; filters: `[Source].[Source Name]` List Key [Eclipse, SBS_Eclipse]; `[Claim].[Claim Reference]` Single Key [1061009]. Expected shape: the measure on COLUMNS, the two levels crossjoined on ROWS, the Source list in a subselect over `[ERS_Cube]`, the Claim Reference single value in WHERE.

## Spec sampler (stands in for the missing UI)
`SpecSampler` (seeded, deterministic) reads a metadata snapshot and produces valid QuerySpec JSONs: 1 to 3 measures; 0 to 3 levels from DIFFERENT hierarchies (visible, not All levels); 0 to 3 filters on other hierarchies, mixing Single, List and Range, with clearly fake values (`SAMPLE_1`...). Optional values pool: a JSON map from hierarchy unique name to an array of real values that I fill in myself, so sampled specs can run against the real cube.
Tests, for many seeds: every sampled spec validates with zero Errors, generates MDX without throwing, output has balanced brackets, braces and parentheses, and repeated generation is identical.

## CLI (group `mdx`)
- `aqbe mdx generate --spec <file> --metadata <file> [--filter-placement hybrid|allsubselect] [--crossjoin operator|function] [--no-nonempty] [--header-comment]` prints MDX to stdout; warnings go to stderr; errors print every issue and exit 1.
- `aqbe mdx validate --spec <file> --metadata <file>`
- `aqbe mdx sample --metadata <file> --seed <n> --count <n> --out <dir> [--values-pool <file>]`
- Exit codes: 0 ok (warnings allowed), 1 invalid input, 2 usage error. Provide 3 sample specs in samples/specs/.

## Tests
- Golden MDX files under tests/XYZ.AQBE.Mdx.Tests/Golden, compared after normalizing line endings (written only when `UPDATE_GOLDEN=1`; otherwise a missing golden fails). Review every golden by eye.
- Generation cases: minimal; multi-measure; 1, 2 and 3 dimensions in both CrossJoin styles; zero dimensions with filters; no filters; Single, List, Range; Single in WHERE; two Singles in one WHERE tuple; Single on the same hierarchy as a dimension (must be a subselect); Single on a DIFFERENT hierarchy of the same dimension (WHERE); mixed filters and nesting order; Key, Name, UniqueName; `]` escaping everywhere; letter-case canonicalization; AllSubselect; `NonEmptyRows=false`; header comment; Examples A and B.
- Validation: one test per code, plus several issues reported together; metadata null versus present.
- JSON: spec deserializes and round-trips unchanged; snapshot with unknown properties still loads; wrong schemaVersion fails clearly.
- Property checks: balanced delimiters, no "\r", identical repeated output.

## Docs
README; DECISIONS.md (including the departures from Excel's MDX and every UNVERIFIED point); docs/VALIDATION.md; docs/VERIFY.md: checks against the real cube: (1) run each generated sample in SSMS; (2) compare numbers with an equivalent Excel pivot, including a calculated and a distinct-count measure; (3) compare Hybrid and AllSubselect on the same filters; (4) confirm whether `*` and `CrossJoin()` behave the same; (5) confirm whether two levels of one hierarchy are rejected; (6) confirm a slicer on one hierarchy works while another hierarchy of the same dimension is on an axis; (7) paste a result into Power Query to see the column naming and whether comments are tolerated.

## Non-goals
No server access, no member fetching, no WITH MEMBER/SET, ORDER, member properties, DRILLTHROUGH, exclusion filters, or pinning unmentioned hierarchies to `[All]` (every hierarchy in the cube I checked except Measures defaults to All). Put ideas in DECISIONS.md instead of building them.

## Process
Propose a short plan and wait for my approval. Implement in this order: Model additions; tokenizer; metadata index; validator; generator; golden tests; sampler; CLI; docs. Run `dotnet build` and `dotnet test` after each step; finish with no failures or warnings. Decide small things yourself and log them. End with a short report: what was built, test counts, UNVERIFIED items, golden files created, open questions.