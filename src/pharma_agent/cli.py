from __future__ import annotations

"""Command-line interface for the presentation agent.

This file is deliberately thin. Its job is only to:
- parse CLI arguments
- translate them into an `OutlineRequest`
- call the service layer
- print human-readable results

All real business logic should stay in `planning.py` and `presentation_builder.py`
so that this file remains easy to audit.
"""

import argparse
import json
from pathlib import Path

from .models import OutlineRequest
from .planning import PlanningService


# These help strings are long on purpose because the CLI is one of the first
# places a new user will look when they are unsure about the right mode.
OUTPUT_HELP = "Optional output PPTX path. If the file already exists, the tool updates that deck in place."
EXISTING_HELP = "Optional existing PPTX to refine. Its content flow and layouts are used as a refinement target instead of drafting from scratch."


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI surface used by `main.py`.

    We expose four modes:
    - `list-use-cases`: inspect available use-case folders
    - `plan`: create `content_v1` without touching PowerPoint
    - `build`: turn an existing `content_v1.json` into a PPT
    - `run`: do planning and PPT writing in one command
    """

    parser = argparse.ArgumentParser(description="Pharma presentation agent CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser("list-use-cases", help="List configured use cases")
    list_parser.set_defaults(handler=handle_list_use_cases)

    plan_parser = subparsers.add_parser("plan", help="Generate content_v1 from context.txt")
    plan_parser.add_argument("--use-case", required=True, help="Use case folder name under context/")
    plan_parser.add_argument("--context", default="", help="Optional explicit path to context.txt")
    plan_parser.add_argument("--existing-pptx", default="", help=EXISTING_HELP)
    plan_parser.add_argument("--generate-json-output", action="store_true", help="Force creation of content_v1 JSON, Markdown, and trace sidecar files")
    plan_parser.set_defaults(handler=handle_plan)

    build_parser_cmd = subparsers.add_parser("build", help="Build or update a PowerPoint from a content_v1 JSON file")
    build_parser_cmd.add_argument("--content", required=True, help="Path to content_v1 JSON")
    build_parser_cmd.add_argument("--generate-json-output", action="store_true", help="Force creation of the trace JSON file")
    build_parser_cmd.add_argument("--output-pptx", default="", help=OUTPUT_HELP)
    build_parser_cmd.add_argument("--existing-pptx", default="", help=EXISTING_HELP)
    build_parser_cmd.set_defaults(handler=handle_build)

    run_parser = subparsers.add_parser("run", help="Plan and build in one command using context.txt")
    run_parser.add_argument("--use-case", required=True, help="Use case folder name under context/")
    run_parser.add_argument("--context", default="", help="Optional explicit path to context.txt")
    run_parser.add_argument("--existing-pptx", default="", help=EXISTING_HELP)
    run_parser.add_argument("--generate-json-output", action="store_true", help="Keep content_v1 and trace files instead of only the PPTX")
    run_parser.add_argument("--output-pptx", default="", help=OUTPUT_HELP)
    run_parser.set_defaults(handler=handle_run)
    return parser


def handle_list_use_cases(args: argparse.Namespace) -> int:
    """Print the available use cases as JSON for easy inspection."""
    print(json.dumps(PlanningService(Path.cwd()).list_use_cases(), indent=2))
    return 0


def handle_plan(args: argparse.Namespace) -> int:
    """Create `content_v1` only.

    This mode is useful when the user wants to review or manually edit the
    planned storyline before anything writes back into PowerPoint.
    """

    service = PlanningService(Path.cwd())
    request = OutlineRequest(
        useCaseId=args.use_case,
        contextPath=str(Path(args.context).resolve()) if args.context else "",
        generateJsonOutput=args.generate_json_output,
        existingDeckPath=str(Path(args.existing_pptx).resolve()) if args.existing_pptx else "",
    )
    json_path, md_path = service.generate_content_plan(request, Path.cwd() / "output")
    data = json.loads(json_path.read_text(encoding="utf-8"))
    print(f"Generation mode: {data.get('generationMode', 'unknown')}")
    print(f"Generation message: {data.get('generationMessage', '')}")
    print(f"content_v1 JSON: {json_path}")
    print(f"content_v1 Markdown: {md_path}")
    return 0


def handle_build(args: argparse.Namespace) -> int:
    """Write or refine a PPT from an already-created `content_v1.json`."""
    service = PlanningService(Path.cwd())
    result = service.build_presentation(
        Path(args.content).resolve(),
        Path.cwd() / "output",
        generate_json_output=args.generate_json_output,
        output_pptx_path=args.output_pptx,
        existing_pptx_path=args.existing_pptx,
    )
    _print_build_result(result)
    return 0


def handle_run(args: argparse.Namespace) -> int:
    """Do the entire plan-then-build workflow in one command."""
    service = PlanningService(Path.cwd())
    request = OutlineRequest(
        useCaseId=args.use_case,
        contextPath=str(Path(args.context).resolve()) if args.context else "",
        generateJsonOutput=args.generate_json_output,
        outputPptxPath=args.output_pptx,
        existingDeckPath=str(Path(args.existing_pptx).resolve()) if args.existing_pptx else "",
    )
    result, generation_mode, generation_message = service.run_pipeline(request, Path.cwd() / "output")
    print(f"Generation mode: {generation_mode}")
    print(f"Generation message: {generation_message}")
    _print_build_result(result)
    return 0


def _print_build_result(result) -> None:
    """Keep CLI output consistent between `build` and `run`."""
    print(f"Presentation: {result.pptxPath}")
    if result.tracePath:
        print(f"Trace: {result.tracePath}")
    if result.warnings:
        print("Warnings:")
        for warning in result.warnings:
            print(f"- {warning}")


def main() -> int:
    """Entry point called by `main.py`."""
    parser = build_parser()
    args = parser.parse_args()
    return args.handler(args)
