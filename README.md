# Pharma Presentation Agent

Local Python CLI for generating pharma business-development PowerPoint drafts from a single `context.txt` file per use case, while using prior client decks only as style and structure references.

## What the tool does

- Reads `context/<Use_Case>/context.txt`
- Reads prior `.pptx` files from `context/<Use_Case>/reference_decks/`
- Uses Gemini when `GEMINI_API_KEY` is available
- By default produces only the final `.pptx`
- Optionally keeps editable `content_v1` and trace JSON sidecars when `Generate JSON Output` is `true` or the CLI flag is used
- Builds a fresh `.pptx` that inherits the selected client deck's slide master and layout family

## Safety rules

- Historical decks are style and layout references only.
- Current project facts, numbers, funnel details, and recommendations must come from `context.txt`.
- If `context.txt` is too weak, the tool stops and tells you what to improve.
- If the content plan requests facts that are not in `context.txt`, placeholders should remain instead of inventing claims.

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
  Patient_Event_Prediction/
    context.txt
    reference_decks/
output/
src/pharma_agent/
main.py
requirements.txt
```

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

Preferred option: environment variable

```powershell
$env:GEMINI_API_KEY="your-gemini-key"
```

Optional option: `.env` file in repo root

```text
GEMINI_API_KEY=your-gemini-key
```

A starter file is included at `.env.example`.

Do not commit real API keys to git.

## How to prepare `context.txt`

Each use case should have one editable `context.txt`. That file should contain:

- project basics
- slide count target
- template deck name or path
- `Generate JSON Output` as `true` or `false`
- current business question
- approved facts and numbers
- workflow or patient funnel
- analytical outputs available
- recommendations or hypotheses
- caveats and formatting requests

The tool reads comments in the file as guidance, but only `## Heading` sections are parsed as structured inputs.

## Commands

### List available use cases

```powershell
python main.py list-use-cases
```

### Fastest workflow: build the PPT directly

This is the most convenient default path. It reads `context.txt`, builds the plan internally, and only keeps the final PowerPoint unless JSON output is enabled.

```powershell
python main.py run --use-case referral_analysis
```

### Generate editable `content_v1`

Use this only when you want to inspect or manually edit the intermediate structure before PowerPoint generation.

```powershell
python main.py plan --use-case referral_analysis --generate-json-output
```

Outputs:

- `output/<timestamp>_<use_case>_content_v1.json`
- `output/<timestamp>_<use_case>_content_v1.md`

### Build a PPT from edited `content_v1`

```powershell
python main.py build --content output\<timestamp>_referral_analysis_content_v1.json --generate-json-output
```

If you omit `--generate-json-output` here, the PPT is created without an extra trace file.

## Recommended workflow

1. Update `context/<Use_Case>/context.txt`
2. Usually run `python main.py run --use-case <use_case>`
3. Only if you want manual intermediate editing, set `Generate JSON Output` to `true` or pass `--generate-json-output`
4. If needed, edit `content_v1.json`
5. Run `build` on that edited file

## Slide master and formatting behavior

- The builder uses the `Template Deck` specified in `context.txt`.
- If `Template Deck` is blank, it falls back to the first deck in `reference_decks/`.
- The output presentation inherits the client deck's slide master by starting from that deck template.
- Different slide archetypes use different layout choices and rendering patterns, so the output deck is not a stack of identical slides.

## Current limitations

- The tool does not yet extract charts or tables directly from Excel.
- Generated visuals are placeholders, not final analytics charts.
- Layout fidelity depends on how reusable the selected client template deck is.
- If Gemini is unavailable, the tool falls back to a deterministic offline draft mode.
