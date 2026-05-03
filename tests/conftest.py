from pathlib import Path


TEST_TMP_DIR = Path(__file__).resolve().parents[1] / "_tmp" / "pytest"
TEST_TMP_DIR.mkdir(parents=True, exist_ok=True)
