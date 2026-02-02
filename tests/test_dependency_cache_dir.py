import os
import sys
import unittest
from pathlib import Path


class TestDependencyCacheDir(unittest.TestCase):
    def test_cache_dir_resolves_to_project_root_from_subdir(self):
        repo_root = Path(__file__).resolve().parent.parent

        sys.path.insert(0, os.path.abspath("scripts"))
        import convert_document  # noqa: E402

        old_cwd = os.getcwd()
        try:
            os.chdir(repo_root / "scripts")
            cache_dir = convert_document.get_cache_dir()
            self.assertEqual(
                cache_dir,
                repo_root / ".claude" / "cache" / "docugenius-converter",
            )

            convert_document.check_dependencies(".docx")
            self.assertTrue((cache_dir / "dependencies.json").exists())
        finally:
            os.chdir(old_cwd)


if __name__ == "__main__":
    unittest.main()

