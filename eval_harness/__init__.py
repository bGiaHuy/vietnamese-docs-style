"""Completely decoupled independent evaluation harness for Vietnamese DOCX workflows."""

from eval_harness.diff_ntpr import (
    audit_namespaces,
    check_non_target_preservation,
    compute_xml_ntpr,
)
from eval_harness.grader import (
    check_compliance_profile,
    crosscheck_dev_ab,
    grade_case,
    grade_directory,
)
from eval_harness.ooxml_reader import inspect_docx
from eval_harness.render_inspector import render_and_inspect
from eval_harness.schema_validator import validate_package
from eval_harness.stats import compare_arms, summarize_experiment

__version__ = "0.1.0"

__all__ = [
    "__version__",
    "audit_namespaces",
    "check_compliance_profile",
    "check_non_target_preservation",
    "compare_arms",
    "compute_xml_ntpr",
    "crosscheck_dev_ab",
    "grade_case",
    "grade_directory",
    "inspect_docx",
    "render_and_inspect",
    "summarize_experiment",
    "validate_package",
]
