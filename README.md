# Pharma Presentation Agent

Local Python CLI for generating pharma business-development PowerPoint drafts from a single `context.txt` file per use case, while using prior client decks only as style and structure references.

## What this repo is for

This tool helps you create a fresh client-style PowerPoint deck for pharma analytics and business-development work.

It is designed around three ideas:
- `context.txt` is the only file that should contain current-project facts and instructions
- older client decks are used only for tone, structure, slide-master styling, and layout inspiration
- the fastest path should be simple: update `context.txt` and run one command to get a `.pptx`

## Fastest workflow

For most use cases, this is all you need:

```powershell
python main.py run --use-case referral_analysis
```

That command:
- reads `context/Referral_Analysis/context.txt`
- reads the reference decks in `context/Referral_Analysis/reference_decks/`
- uses Gemini if the API key is configured
- builds the final `.pptx`
- by default keeps only the `.pptx` in `output/`

## How the tool thinks about sources

### Claim-bearing source

This is the only place current project facts should come from:
- `context/<Use_Case>/context.txt`

Examples:
- approved numbers
- current business question
- patient funnel
- recommendations
- constraints

### Style-only source

These are not treated as truth for the new project:
- old client `.pptx` files in `reference_decks/`

They are used for:
- slide flow
- storytelling style
- wording tone
- client master/theme inheritance
- layout family selection

## Repo structure

```text
context/
  Referral_Analysis/
    context.txt
    reference_decks/
    assets/
    notes/
  Segmentation/
    context.txt
    reference_decks/
    assets/
    notes/
  Patient_Event_Prediction/
    context.txt
    reference_decks/
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

## What each file and folder does

### Root files

- `main.py`
  Thin entrypoint. You run this file for all commands.

- `requirements.txt`
  Python dependencies for the repo.

- `.env.example`
  Example environment file showing how to provide `GEMINI_API_KEY`.

- `.gitignore`
  Prevents local-only folders and secrets from being committed.

- `README.md`
  This guide.

### Source code

- `src/pharma_agent/cli.py`
  Defines the commands: `list-use-cases`, `plan`, `build`, and `run`.

- `src/pharma_agent/planning.py`
  Main orchestration logic. Reads context, validates it, selects template decks, creates the intermediate content plan, and triggers PPT generation.

- `src/pharma_agent/presentation_builder.py`
  Creates the final PowerPoint using `python-pptx`, while inheriting the selected client deck's slide master and using different layout patterns.

- `src/pharma_agent/context_manager.py`
  Reads and validates `context.txt`. Also resolves whether JSON sidecar outputs should be kept.

- `src/pharma_agent/pptx_reference.py`
  Extracts text and style cues from older decks in `reference_decks/`.

- `src/pharma_agent/llm.py`
  Handles Gemini API access through environment variables or `.env`.

- `src/pharma_agent/models.py`
  Shared data structures used across the app.

### Context folders

Each use case under `context/` is a working area.

- `context.txt`
  The main file you should edit for each new project.

- `reference_decks/`
  Old client decks that help the tool inherit structure and style.

- `assets/`
  Currently optional and unused by the app. Keep future charts, screenshots, or images here if you want a place for them.

- `notes/`
  Currently optional and unused by the app. Keep supporting notes or raw thinking here if useful.

### Output folder

- `output/`
  Stores generated artifacts.

Default behavior:
- only the final `.pptx` is kept

Optional behavior when JSON output is enabled:
- `content_v1.json`
- `content_v1.md`
- `trace.json`
- final `.pptx`

### Tests

- `tests/test_context_validation.py`
  Basic checks for use-case discovery, context validation, and content-plan generation.

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

## Gemini API key setup

Preferred option:

```powershell
$env:GEMINI_API_KEY="your-gemini-key"
```

Optional local file:

```text
GEMINI_API_KEY=your-gemini-key
```

Place that in a local `.env` file in the repo root.

Notes:
- `.env` is ignored by git
- `.env.example` is only a sample
- do not commit a real API key

## How to prepare `context.txt`

Every new project should be driven by one `context.txt` file.

Important fields:
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

### Most important practical rule

If a fact should appear in the deck, put it in `context.txt`.

Do not assume the model should infer:
- numbers
- recommendations
- current patient funnel details
- brand-specific conclusions

## Commands

### List available use cases

```powershell
python main.py list-use-cases
```

### Run end-to-end and keep only the PPT

This is the recommended default command.

```powershell
python main.py run --use-case referral_analysis
```

Use this when:
- your `context.txt` is already strong
- you want the simplest workflow
- you do not want extra intermediate files

### Generate editable intermediate files

Use this when you want to inspect or manually edit the structure before building the deck.

```powershell
python main.py plan --use-case referral_analysis --generate-json-output
```

This creates:
- `content_v1.json`
- `content_v1.md`

### Build from an edited JSON file

```powershell
python main.py build --content output\<timestamp>_referral_analysis_content_v1.json --generate-json-output
```

Use this when:
- you edited the intermediate JSON manually
- you want to preserve a trace file too

## Recommended ways to use the repo

### Mode 1: Fast production path

Best when you already know what you want.

1. Update `context/<Use_Case>/context.txt`
2. Run `python main.py run --use-case <use_case>`
3. Open the generated `.pptx`
4. Make final manual polish edits in PowerPoint if needed

### Mode 2: Review before PPT

Best when the story is still evolving.

1. Set `Generate JSON Output` to `true` in `context.txt` or pass `--generate-json-output`
2. Run `python main.py plan --use-case <use_case> --generate-json-output`
3. Review `content_v1.md` or edit `content_v1.json`
4. Run `python main.py build --content <content_v1.json> --generate-json-output`

### Mode 3: Use a different context file

Useful if you want project-specific files outside the default folder.

```powershell
python main.py run --use-case referral_analysis --context C:\path\to\context.txt
```

## Understanding outputs

### Default output

- `output/<timestamp>_<use_case>_deck.pptx`

### Optional outputs

- `content_v1.json`
  Structured editable plan of the presentation.

- `content_v1.md`
  Human-readable Markdown rendering of the same plan.

- `trace.json`
  Provenance file showing which slide pieces came from project context vs style reference.

## How to read the Markdown output locally

If you generate `content_v1.md`, the easiest way to render it locally is with VS Code.

Open it:

```powershell
code output\your_file.md
```

Then in VS Code:
- `Ctrl+Shift+V` for preview
- or `Ctrl+K` then `V` for side-by-side preview

## Slide master and formatting behavior

- The builder uses the `Template Deck` specified in `context.txt`
- If `Template Deck` is blank, it falls back to the first deck in `reference_decks/`
- The output presentation inherits the client deck's slide master by starting from that deck template
- Different slide archetypes use different layout choices and rendering styles

Current archetypes include:
- title
- agenda
- executive summary
- framework
- patient funnel
- insight
- recommendation
- appendix

So the output should not look like the same slide repeated over and over.

## About the extra folders

- `.agents/` and `.codex/`
  These are Codex/workspace metadata folders, not part of the app logic. They are ignored in git.

- `assets/` and `notes/` inside each use case
  These are optional placeholders. The current app does not require them, but they are there in case you want to store visuals or supporting notes beside a use case.

## Current limitations

- The tool does not yet extract charts or tables directly from Excel
- Generated visuals are placeholders, not final analytics charts
- Layout fidelity depends on how reusable the selected client template deck is
- If Gemini is unavailable, the tool falls back to a deterministic offline draft mode

## Sanity-check commands

Run tests:

```powershell
python -m unittest discover -s tests -v
```

See use cases:

```powershell
python main.py list-use-cases
```

Generate a deck quickly:

```powershell
python main.py run --use-case referral_analysis
```
