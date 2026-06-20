from __future__ import annotations

import argparse
import json
from pathlib import Path

from .models import OutlineRequest
from .planning import PlanningService


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Pharma presentation agent CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser("list-use-cases", help="List configured use cases")
    list_parser.set_defaults(handler=handle_list_use_cases)

    plan_parser = subparsers.add_parser("plan", help="Generate content_v1 from context.txt")
    plan_parser.add_argument("--use-case", required=True, help="Use case folder name under context/")
    plan_parser.add_argument("--context", default="", help="Optional explicit path to context.txt")
    plan_parser.add_argument("--generate-json-output", action="store_true", help="Force creation of content_v1 JSON, Markdown, and trace sidecar files")
    plan_parser.set_defaults(handler=handle_plan)

    build_parser_cmd = subparsers.add_parser("build", help="Build a PowerPoint from a content_v1 JSON file")
    build_parser_cmd.add_argument("--content", required=True, help="Path to content_v1 JSON")
    build_parser_cmd.add_argument("--generate-json-output", action="store_true", help="Force creation of the trace JSON file")
    build_parser_cmd.set_defaults(handler=handle_build)

    run_parser = subparsers.add_parser("run", help="Plan and build in one command using context.txt")
    run_parser.add_argument("--use-case", required=True, help="Use case folder name under context/")
    run_parser.add_argument("--context", default="", help="Optional explicit path to context.txt")
    run_parser.add_argument("--generate-json-output", action="store_true", help="Keep content_v1 and trace files instead of only the PPTX")
    run_parser.set_defaults(handler=handle_run)

    return parser


def handle_list_use_cases(args: argparse.Namespace) -> int:
    service = PlanningService(Path.cwd())
    print(json.dumps(service.list_use_cases(), indent=2))
    return 0


def handle_plan(args: argparse.Namespace) -> int:
    service = PlanningService(Path.cwd())
    request = OutlineRequest(
        useCaseId=args.use_case,
        contextPath=str(Path(args.context).resolve()) if args.context else "",
        generateJsonOutput=args.generate_json_output,
    )
    output_dir = Path.cwd() / "output"
    json_path, md_path = service.generate_content_plan(request, output_dir)
    print(f"content_v1 JSON: {json_path}")
    print(f"content_v1 Markdown: {md_path}")
    return 0


def handle_build(args: argparse.Namespace) -> int:
    service = PlanningService(Path.cwd())
    result = service.build_presentation(
        Path(args.content).resolve(),
        Path.cwd() / "output",
        generate_json_output=args.generate_json_output,
    )
    print(f"Presentation: {result.pptxPath}")
    if result.tracePath:
        print(f"Trace: {result.tracePath}")
    if result.warnings:
        print("Warnings:")
        for warning in result.warnings:
            print(f"- {warning}")
    return 0


def handle_run(args: argparse.Namespace) -> int:
    service = PlanningService(Path.cwd())
    request = OutlineRequest(
        useCaseId=args.use_case,
        contextPath=str(Path(args.context).resolve()) if args.context else "",
        generateJsonOutput=args.generate_json_output,
    )
    result = service.run_pipeline(request, Path.cwd() / "output")
    print(f"Presentation: {result.pptxPath}")
    if result.tracePath:
        print(f"Trace: {result.tracePath}")
    if result.warnings:
        print("Warnings:")
        for warning in result.warnings:
            print(f"- {warning}")
    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.handler(args)
