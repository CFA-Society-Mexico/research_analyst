"""Contratos del repo: lo que la documentacion promete coincide con lo que existe.

Atrapa la deriva entre docs, manifests y codigo (version en tres lugares, una
skill sin ruta en AGENTS.md, un check F documentado sin implementar, una ruta
citada que ya no existe). Stdlib unittest, sin red.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILLS_DIR = ROOT / "skills"
COMMANDS_DIR = ROOT / "commands"
CHECKS_DOC = SKILLS_DIR / "model-standards" / "references" / "integrity-checks.md"
BUILDER = ROOT / "tools" / "xlsx_builder.py"

# Documentos cuyas rutas citadas deben existir en el repo.
ROUTED_DOCS = ("AGENTS.md", "README.md", "CONTRIBUTING.md", "docs/architecture.md")

_REPO_PATH = re.compile(r"`((?:commands|skills|templates|tools|docs|tests)/[^`\s]+)`")
_CHECK_ROW = re.compile(r"^\| ([SCDF]\d+[a-z]?) \|", re.MULTILINE)
_FINDING_ID = re.compile(r'Finding\(\s*"(F\d+)[ "]')


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def _skill_names() -> list[str]:
    return sorted(p.name for p in SKILLS_DIR.iterdir() if p.is_dir())


def _command_names() -> list[str]:
    return sorted(p.name for p in COMMANDS_DIR.glob("*.md"))


class VersionTests(unittest.TestCase):

    def test_version_is_the_same_in_manifests_and_readme(self) -> None:
        claude = json.loads(_read(".claude-plugin/plugin.json"))["version"]
        codex = json.loads(_read(".codex-plugin/plugin.json"))["version"]
        readme = re.search(r"\bv(\d+\.\d+\.\d+)\b", _read("README.md"))
        self.assertIsNotNone(readme, "README.md no declara version vX.Y.Z")
        assert readme is not None
        self.assertEqual({claude, codex, readme.group(1)}, {claude},
                         "version distinta entre .claude-plugin, .codex-plugin y README")

    def test_manifests_agree_on_plugin_name(self) -> None:
        claude = json.loads(_read(".claude-plugin/plugin.json"))["name"]
        codex = json.loads(_read(".codex-plugin/plugin.json"))["name"]
        market = json.loads(_read(".claude-plugin/marketplace.json"))["plugins"]
        self.assertEqual(claude, codex)
        self.assertIn(claude, [p["name"] for p in market])


class SkillTests(unittest.TestCase):

    def test_every_skill_has_frontmatter_matching_its_folder(self) -> None:
        for name in _skill_names():
            with self.subTest(skill=name):
                text = _read(f"skills/{name}/SKILL.md")
                self.assertTrue(text.startswith("---\n"), "sin frontmatter")
                front = text.split("---\n", 2)[1]
                declared = re.search(r"^name:\s*(\S+)\s*$", front, re.MULTILINE)
                self.assertIsNotNone(declared, "frontmatter sin name")
                assert declared is not None
                self.assertEqual(declared.group(1), name)
                self.assertRegex(front, r"(?m)^description:\s*\S", "frontmatter sin description")

    def test_documented_skill_count_matches_folders(self) -> None:
        actual = len(_skill_names())
        for rel, pattern in (("README.md", r"Las (\d+) skills"),
                             ("CONTRIBUTING.md", r"tiene (\d+) skills")):
            with self.subTest(doc=rel):
                found = re.search(pattern, _read(rel))
                self.assertIsNotNone(found, f"{rel} ya no declara el numero de skills")
                assert found is not None
                self.assertEqual(int(found.group(1)), actual)


class RoutingTests(unittest.TestCase):

    def test_agents_router_covers_every_command_and_skill(self) -> None:
        agents = _read("AGENTS.md")
        expected = [f"commands/{c}" for c in _command_names()]
        expected += [f"skills/{s}/SKILL.md" for s in _skill_names()]
        missing = [path for path in expected if f"`{path}`" not in agents]
        self.assertEqual(missing, [], "sin ruta en AGENTS.md (runtimes sin skills nativas no lo encuentran)")

    def test_readme_lists_every_skill_and_command(self) -> None:
        readme = _read("README.md")
        missing = [s for s in _skill_names() if f"`{s}`" not in readme]
        missing += [c for c in _command_names() if f"/{c.removesuffix('.md')}" not in readme]
        self.assertEqual(missing, [])

    def test_repo_paths_cited_in_docs_exist(self) -> None:
        docs = list(ROUTED_DOCS)
        docs += [f"commands/{c}" for c in _command_names()]
        docs += [f"skills/{s}/SKILL.md" for s in _skill_names()]
        dangling: list[str] = []
        for rel in docs:
            for cited in _REPO_PATH.findall(_read(rel)):
                if any(ch in cited for ch in "*<>"):
                    continue  # glob o placeholder, no una ruta concreta
                target = cited.split("::", 1)[0].rstrip(".,;:)")
                if not (ROOT / target).exists():
                    dangling.append(f"{rel} -> {cited}")
        self.assertEqual(dangling, [], "rutas citadas que no existen")


class IntegrityCheckTests(unittest.TestCase):

    def test_check_ids_are_unique_in_the_doc(self) -> None:
        ids = _CHECK_ROW.findall(CHECKS_DOC.read_text(encoding="utf-8"))
        duplicated = sorted({i for i in ids if ids.count(i) > 1})
        self.assertEqual(duplicated, [], "los IDs de check no se reciclan ni se repiten")

    def test_documented_f_checks_are_exactly_the_implemented_ones(self) -> None:
        documented = {i for i in _CHECK_ROW.findall(CHECKS_DOC.read_text(encoding="utf-8"))
                      if i.startswith("F")}
        source = BUILDER.read_text(encoding="utf-8")
        implemented = set(_FINDING_ID.findall(source))
        # F20 (cobertura del audit) se emite con el ID armado en tiempo de corrida.
        implemented |= {i for i in documented if f'"{i} ' in source or f"'{i} " in source}
        self.assertEqual(documented - implemented, set(), "check F documentado sin implementar")
        self.assertEqual(implemented - documented, set(), "check F implementado sin documentar")


if __name__ == "__main__":
    unittest.main()
