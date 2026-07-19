import tomllib
import unittest
from pathlib import Path


class DeploymentManifestTests(unittest.TestCase):
    def test_vercel_runtime_dependencies_are_declared(self):
        manifest = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))
        dependencies = manifest["project"]["dependencies"]
        normalized = {item.split("[", 1)[0].split("=", 1)[0].split(">", 1)[0] for item in dependencies}
        self.assertTrue(
            {"anthropic", "fastapi", "psycopg", "yfinance", "openpyxl"}.issubset(normalized)
        )


if __name__ == "__main__":
    unittest.main()
