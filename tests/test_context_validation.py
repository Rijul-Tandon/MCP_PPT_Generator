import unittest
from pathlib import Path

from src.pharma_agent.context_manager import ContextManager
from src.pharma_agent.excel_context import ExcelContextLibrary
from src.pharma_agent.models import OutlineRequest
from src.pharma_agent.planning import PlanningService


ROOT = Path(__file__).resolve().parents[1]


class ContextValidationTests(unittest.TestCase):
    def test_list_use_cases_includes_placeholders(self) -> None:
        manager = ContextManager(ROOT)
        use_cases = {item.id: item for item in manager.list_use_cases()}
        self.assertIn("referral_analysis", use_cases)
        self.assertTrue(use_cases["referral_analysis"].hasContextFile)
        self.assertIn("segmentation", use_cases)
        self.assertIn("patient_event_prediction", use_cases)
        self.assertIn("excelContextFileCount", use_cases["referral_analysis"].__dict__)

    def test_validation_flags_missing_sections(self) -> None:
        manager = ContextManager(ROOT)
        result = manager.validate_context("## Project Name\nOnly a title")
        self.assertFalse(result.isValid)
        self.assertIn("Audience", result.missingSections)

    def test_runtime_excel_context_block_is_generated(self) -> None:
        library = ExcelContextLibrary()
        block = library.build_runtime_context_block(
            {
                "files": [
                    {
                        "fileName": "sample.xlsx",
                        "numericSamples": ["Sheet1: 42"],
                        "textSamples": ["Sheet1: Top corridor is NY to NJ"],
                    }
                ]
            }
        )
        self.assertIn("Runtime Excel Context", block)
        self.assertIn("sample.xlsx", block)
        self.assertIn("42", block)

    def test_plan_generates_content_v1(self) -> None:
        service = PlanningService(ROOT)
        request = OutlineRequest(
            useCaseId="referral_analysis",
            contextPath=str(ROOT / "context" / "Referral_Analysis" / "context.txt"),
            generateJsonOutput=True,
        )
        json_path, md_path = service.generate_content_plan(request, ROOT / "output")
        self.assertTrue(json_path.exists())
        self.assertTrue(md_path.exists())
        self.assertIn("content_v1", json_path.name)


if __name__ == "__main__":
    unittest.main()
