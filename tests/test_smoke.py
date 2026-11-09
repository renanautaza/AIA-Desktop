"""Dependency-free validation for the current AIA Desktop source."""

import ast
from pathlib import Path


root = Path(__file__).parent.parent
source = root / "src" / "Autaza_Inspection_Analytics.py"
tree = ast.parse(source.read_text(encoding="utf-8"), filename=str(source))

assert any(
    isinstance(node, ast.Assign)
    and any(isinstance(target, ast.Name) and target.id == "VERSION" for target in node.targets)
    for node in tree.body
), "VERSION is not defined"

assert "1.1.6" in source.read_text(encoding="utf-8"), "Expected current version 1.1.6"
print("Smoke test passed: source parses and identifies AIA Desktop 1.1.6.")
