# Pharma Presentation Agent

Local Python CLI for generating and refining pharma business-development PowerPoint decks from a structured `context.txt`, optional Excel support files, historical reference decks, and optionally an already-existing presentation that needs improvement.

## What this repo does

This tool is built for pharma analytics and business-development presentations where we want to:
- use `context.txt` as the primary source of current approved project facts
- pull fresh supporting numbers from `excel_context/` on every run
- learn tone, flow, and layout patterns from historical decks without copying old claims
- optionally refine an existing PPT instead of drafting a new one from scratch

## Core behavior

### Claim-bearing inputs

These are allowed to drive facts, numbers, recommendations, and project-specific statements:
- `context/<Use_Case>/context.txt`
- approved files inside `context/<Use_Case>/excel_context/`
- optional user edits to `content_v1.json`

### Style/reference inputs

These help with structure and presentation style only:
- historical decks in `reference_decks/`
- an optional existing PPT passed through `--existing-pptx`

Important rule:
- old decks can influence flow, visual rhythm, and layout style
- old decks must not be reused as factual sources for current metrics or business claims

## Repo structure

```text
context/
  Referral_Analysis/
    context.txt
    reference_decks/
    excel_context/
    assets/
    notes/
  Segmentation/
    context.txt
    reference_decks/
    excel_context/
    assets/
    notes/
  Patient_Event_Prediction/
    context.txt
    reference_decks/
    excel_context/
    assets/
    notes/
output/
src/pharma_agent/
tests/
main.py
requirements.txt
README.md
.env.example
.gitignore
```

## Important files

- `main.py`
  Entry point for all commands.

- `src/pharma_agent/cli.py`
  Command-line interface for `list-use-cases`, `plan`, `build`, and `run`.

- `src/pharma_agent/context_manager.py`
  Loads `context.txt`, parses sections, validates required fields, and discovers Excel files.

- `src/pharma_agent/excel_context.py`
  Reads Excel/CSV/TSV files and creates runtime context that is regenerated every time the tool runs.

- `src/pharma_agent/pptx_reference.py`
  Extracts slide text and style patterns from reference decks and existing decks.

- `src/pharma_agent/planning.py`
  Main orchestration flow. Builds the runtime context, prepares the prompt, drafts the slide plan, and decides whether the run is a new build or a refinement pass.

- `src/pharma_agent/presentation_builder.py`
  Writes PowerPoint output. It supports both brand-new deck generation and true in-place refinement of an existing deck.

- `src/pharma_agent/llm.py`
  Minimal provider wrapper for Groq and Gemini.

- `tests/test_context_validation.py`
  Lightweight regression coverage for validation and runtime Excel context behavior.

## Setup

### 1. Create a virtual environment

```powershell
python -m venv bd_venv
```

### 2. Activate it

```powershell
.\bd_venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

## LLM provider setup

Preferred example using Groq:

```powershell
$env:GROQ_API_KEY="your-groq-key"
```

Optional local `.env` example:

```text
LLM_PROVIDER=groq
GROQ_API_KEY=your-groq-key
GROQ_MODEL=llama-3.3-70b-versatile

or

GEMINI_API_KEY=your-gemini-key
GEMINI_MODEL=gemini-2.0-flash
```

Notes:
- `.env` is ignored by git
- `.env.example` is only a sample
- do not commit a real key

## How to prepare `context.txt`

Every project should be driven by a strong `context.txt`.

Important sections:
- `Project Name`
- `Client / Brand / Indication`
- `Use-Case Type`
- `Audience`
- `Slide Count Target`
- `Template Deck`
- `Generate JSON Output`
- `Presentation Objective`
- `Current Business Question`
- `Approved Facts and Numbers`
- `Patient Funnel or Workflow Summary`
- `Analytical Outputs Available`
- `Draft Recommendations or Hypotheses`
- `Slides to Include or Avoid`
- `Constraints / Caveats`
- `Additional Instructions`

Best practice:
- put every approved current-project claim into `context.txt` or `excel_context/`
- do not expect the model to infer missing business facts
- use `Additional Instructions` to enforce client-specific writing style, talking headers, or slide behavior

## How `excel_context/` works

Supported formats:
- `.xlsx`
- `.csv`
- `.tsv`

Current behavior:
- the tool scans a small number of rows from each file
- builds numeric and text samples
- appends those findings into a runtime Excel context block on every run
- feeds that runtime block into planning without modifying `context.txt` on disk

This means:
- Excel-derived facts are refreshed dynamically each time a PPT is created
- users do not need to manually copy spreadsheet content into `context.txt` for every update

Best use cases:
- KPI tables
- state/country rankings
- patient funnel counts
- HCP or segment counts
- model outputs exported to Excel/CSV

## Commands

### List available use cases

```powershell
python main.py list-use-cases
```

### Generate a deck end to end

```powershell
python main.py run --use-case referral_analysis
```

### Generate `content_v1` for review before PPT build

```powershell
python main.py plan --use-case referral_analysis --generate-json-output
```

### Build from an edited `content_v1.json`

```powershell
python main.py build --content output\<timestamp>_referral_analysis_content_v1.json --generate-json-output
```

## Existing deck refinement

### Refine an already-existing deck

Use this when someone has already made a PPT and you want the agent to improve it using current project context.

```powershell
python main.py run --use-case referral_analysis --existing-pptx "C:\path\to\existing_deck.pptx" --output-pptx "C:\path\to\refined_deck.pptx"
```

What happens in refinement mode:
- the existing deck is read as an input artifact
- its slide text and style patterns are extracted
- the planner uses that as refinement context
- the builder updates existing slides in place where possible
- if the new outline is longer than the old deck, new slides are added at the end

### Review the refined plan before touching the PPT

```powershell
python main.py plan --use-case referral_analysis --existing-pptx "C:\path\to\existing_deck.pptx" --generate-json-output
```

This is the safest flow when:
- the existing deck has sensitive client wording
- you want to review changed talking headers first
- you want to manually edit `content_v1.json` before writing back to the PPT

## Recommended usage patterns

### Mode 1: Fast new-deck generation

1. Update `context/<Use_Case>/context.txt`
2. Add or update files in `excel_context/`
3. Run `python main.py run --use-case <use_case>`
4. Review the generated PPT

### Mode 2: Plan first, then build

1. Run `python main.py plan --use-case <use_case> --generate-json-output`
2. Review `content_v1.md`
3. Optionally edit `content_v1.json`
4. Run `python main.py build --content <content_v1.json> --generate-json-output`

### Mode 3: Refine an existing deck

1. Prepare `context.txt` and `excel_context/`
2. Run `python main.py plan --use-case <use_case> --existing-pptx <deck> --generate-json-output`
3. Review the proposed structure and slide language
4. Run `python main.py run --use-case <use_case> --existing-pptx <deck> --output-pptx <refined deck>`

## What the agent is trying to do well

- use talking headers instead of weak one-word slide titles
- build a story that feels close to client-ready
- use spreadsheet facts dynamically on each run
- preserve the spirit of an existing client deck when refining
- keep style influence and factual influence separate

## Troubleshooting

- If the CLI prints `Generation mode: groq` or `Generation mode: gemini`, the outline came from the live provider.
- If the CLI prints `Generation mode: offline`, check `Generation message` immediately below it.
- If `Generation message` mentions `HTTP 429`, the provider key was recognized but quota or billing blocked the request.
- If `Generation message` says no provider key was configured, set `GROQ_API_KEY` or `GEMINI_API_KEY` and retry.
- If `Generation message` mentions a connection error, the tool could not reach the provider endpoint from the current environment.

## Limitations

- Excel ingestion is still summary-based; it does not build final charts directly from spreadsheets.
- Existing deck refinement is conservative: it updates text-bearing regions in place and preserves slide shells, but it does not yet intelligently remap every original shape with perfect semantic understanding.
- If the refined outline is shorter than the existing deck, trailing old slides are currently left unchanged and a warning is added.
- `.xlsb` is not supported.
- Layout fidelity still depends on the quality and consistency of the reference deck or existing deck.
- If no live LLM provider is available, the tool falls back to deterministic offline drafting.

## Full-potential workflow

For the best results:
1. Write a strong `context.txt` with approved claims and clear instructions.
2. Keep `excel_context/` current so spreadsheet facts refresh automatically.
3. If refining, pass the real client deck with `--existing-pptx`.
4. Use `plan --generate-json-output` when quality matters more than speed.
5. Review `content_v1.md` or `content_v1.json` before the final write-back for important decks.
6. Treat the final output as a strong draft, then do final client polish in PowerPoint if needed.
