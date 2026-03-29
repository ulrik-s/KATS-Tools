# VBA Refactor Proposal: Read → Transform → Render

## Why this refactor

The current code works, but feature work is becoming expensive because:

- Domain calculations, document parsing, and Word rendering are mixed in large procedures (especially in `Processor_KR.bas`).
- Shared state is global (`gMoneyState`, `gCategoryHours`, hearing flags), which creates ordering dependencies between processors.
- The same low-level helpers exist in multiple modules (for example table/date/number helpers in both `KATSUtils.bas` and `Processor_KR.bas`).
- Pipeline orchestration is string-based (`Application.Run`), which is flexible but weakly typed and difficult to validate.

## Target architecture

Use your suggested pattern per processor:

```vb
Public Sub Process_SOMETHING(ByVal content As Range)
    Dim t As Table
    Set t = RequireSingleTable(content)
    If t Is Nothing Then Exit Sub

    Dim model As SomethingModel
    model = ReadSomething(t)

    ApplySomethingState model, ctx
    RenderSomething t, model, ctx
End Sub
```

The key is to enforce three layers and one context object:

1. **Reader layer (Document → Model)**
   - Only reads from `Range/Table`.
   - No mutation of document text/cells.
   - Returns typed model records.

2. **Domain layer (Model → Model/State)**
   - Pure-ish business rules: sums, taxemål logic, rates, hearing calculations.
   - No direct Word object calls.

3. **Renderer layer (Model/State → Document)**
   - Applies final values and formatting to Word.
   - Handles row deletion, spacing, content controls.

4. **Processing context (`ctx`)**
   - Replaces global module state.
   - Carries cross-tag data explicitly (e.g. utlägg totals used by ARVODE_TOTAL).

## Recommended module layout

Create folders/modules by responsibility.

- `src/core/` (or naming prefix `Core_` if keeping flat)
  - `Core_Pipeline.bas` – tag registration + pipeline execution.
  - `Core_Context.bas` – `KatsContext` type + context getters/setters.
  - `Core_TagScanner.bas` – current tag scanning logic from `TagHandler.bas`.

- `src/domain/`
  - `Domain_AR.bas` – all AR/Taxa business rules (no Word API calls).
  - `Domain_Yttrande.bas` – party extraction + name normalization.
  - `Domain_Formatting.bas` – money/date/postort domain formatting rules.

- `src/io/`
  - `Read_AR.bas`, `Render_AR.bas`
  - `Read_Utlagg.bas`, `Render_Utlagg.bas`
  - `Read_Yttrande.bas`, `Render_Yttrande.bas`

- `src/shared/`
  - `Shared_Table.bas` – `CellTextSafe`, `RequireSingleTable`, row finders.
  - `Shared_Parse.bas` – `SvToCurrency`, ISO date parsing.
  - `Shared_Text.bas` – normalize line endings, trim helpers.

- `src/app/`
  - `KATSMain.bas` keeps ribbon entrypoints and calls `Core_Pipeline.Run`.

## Processor contract

Define a consistent processor signature and registration metadata:

- `Public Sub Process_<TAG>(ByVal content As Range, ByRef ctx As KatsContext)`
- Registration table stores:
  - tag name
  - processor name
  - execution order
  - optional preconditions (requires table, requires prior state key)

If you must keep `Application.Run`, pass context through a singleton `CurrentContext` object in `Core_Context`; but prefer direct calls where possible.

## Concrete refactor opportunities in current code

1. **Extract duplicated helpers first**
   - Move duplicate `CellTextSafe`, `CellSetTextSafe`, `LooksLikeIsoDate`, `SvToCurrency`, `FormatSvInt`, etc. into one shared module.
   - Make processors depend only on shared helpers.

2. **Isolate global state**
   - Current globals in `Processor_KR.bas` become fields in `KatsContext`:
     - `MoneyState`
     - `CategoryHours`
     - `IsTaxemal`, `TaxLevel`
     - `HearingStart`, `HearingMinutes`
     - `Postort`

3. **Split `Process_ARVODE`**
   - Reader: read ARVODE table rows/specs into a model.
   - Domain: choose `Taxa` vs normal model, compute rows to keep/remove, compute amounts.
   - Renderer: apply text and delete rows.

4. **Split `Process_ARVODE_TOTAL`**
   - Reader gets row indices once.
   - Domain computes ex/moms/incl.
   - Renderer writes values and optional row removal.

5. **Split `Process_YTTRANDE_PARTER`**
   - Reader captures raw block + parsed parties.
   - Domain resolves name list and defaults.
   - Renderer builds dropdown controls and underline rules.

## Migration plan (low risk)

1. **Phase 1: mechanical extraction**
   - Extract shared helpers and add wrapper calls so behavior is unchanged.

2. **Phase 2: introduce context**
   - Add `KatsContext` and replace module-level globals gradually.

3. **Phase 3: refactor one processor family at a time**
   - Start with `YTTRANDE_*` (smaller blast radius), then `UTLÄGG`, then `ARVODE*`.

4. **Phase 4: harden pipeline**
   - Add validation pass at startup: all registered processors exist, expected signature, tags unique.

5. **Phase 5: tests around domain logic**
   - Keep VBA tests minimal but add deterministic domain test procedures (`Test_Domain_AR_*`) independent from Word tables where possible.

## Design rules for future features

- New feature = add/extend model + domain rule first, renderer last.
- No business calculations inside renderer procedures.
- No direct writes in reader procedures.
- All cross-processor dependencies go through `ctx`, never through module globals.
- Keep parser/formatter functions side-effect free.

## Example: ARVODE flow after refactor

- `Process_ARVODE(content, ctx)`
  1. `table = RequireSingleTable(content)`
  2. `input = ReadArvodeTable(table)`
  3. `result = ComputeArvode(input, ctx)`
  4. `RenderArvodeTable(table, result)`
  5. `ctx.MoneyState.ArvodeExMoms = result.TotalExMoms`

This pattern will make new requests (new row types, alternative tax rules, custom output templates) mostly domain/model changes rather than risky edits across mixed parsing/rendering code.
