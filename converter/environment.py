from __future__ import annotations

import importlib.util
import os
from dataclasses import dataclass, field


@dataclass
class EnvironmentReport:
    warnings: list[str] = field(default_factory=list)
    missing_modules: list[str] = field(default_factory=list)


def check_environment(required_environment: str = "md2docx-env", required_modules: tuple[str, ...] = ("docx", "markdown_it", "yaml", "PIL")) -> EnvironmentReport:
    report = EnvironmentReport()
    if os.environ.get("CONDA_DEFAULT_ENV") != required_environment:
        report.warnings.append(
            f"Current environment is not {required_environment}. Please run: conda activate {required_environment}"
        )
    report.missing_modules = [module for module in required_modules if importlib.util.find_spec(module) is None]
    if report.missing_modules:
        report.warnings.append("Missing Python packages: " + ", ".join(report.missing_modules) + ". Run: pip install -r requirements.txt")
    return report
