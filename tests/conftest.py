from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from vendor_instructions.pipeline import BuildOutcome, Workspace, build

SAMPLE = Path(__file__).resolve().parents[1] / "examples" / "sample_data"
VENDOR = "Northwind Manufacturing"


@pytest.fixture
def workspace(tmp_path: Path) -> Workspace:
    """A private copy of the committed sample workspace."""
    shutil.copytree(SAMPLE, tmp_path / "ws")
    return Workspace(tmp_path / "ws")


@pytest.fixture
def built(workspace: Workspace) -> BuildOutcome:
    return build(workspace, vendor=VENDOR)
