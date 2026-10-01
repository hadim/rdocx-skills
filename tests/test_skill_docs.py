"""The skills against the pinned build: every Python name and keyword argument and every CLI command and flag they
cite exists, their gap pages match tests/gaps.py, and each SKILL.md stays under 200 lines. A pin that renames or
drops an API fails here even when no recipe block runs it (the recipes themselves run in test_docs_snippets.py).

Positional parameters are written as short placeholders in the skills (`bi`, `rid`, `r, c`) and are not checked:
a positional call works whatever the parameter is called. Keyword arguments are checked, since a wrong one raises."""
import ast
import builtins
import functools
import re
import subprocess
from pathlib import Path

import pytest
import rdocx
import rpptx

from conftest import BIN, EXE
from gaps import GAPS

ROOT = Path(__file__).resolve().parent.parent
SKILLS = ROOT / "skills"
DOCS = sorted(p for p in SKILLS.glob("*/SKILL.md")) + sorted(
    p for p in SKILLS.glob("*/references/*.md") if p.name != "recipes.md")
# Names the skills cite from other libraries (python-docx, python-pptx, lxml, zipfile), never from rdocx or rpptx.
FOREIGN = {"Presentation", "Document", "fromstring", "tostring", "xpath", "related_part", "ZipFile"}


def rel(path):
    return path.relative_to(ROOT)


def spans(path):
    """(line number, code span) for every inline code span of a Markdown file, fenced blocks excluded."""
    text, fenced = path.read_text(), False
    for n, line in enumerate(text.splitlines(), 1):
        if line.startswith("```"):
            fenced = not fenced
            continue
        if not fenced:
            yield from ((n, s) for s in re.findall(r"`([^`]+)`", line))


# ---------------------------------------------------------------- Python names
@functools.lru_cache(maxsize=None)
def python_api():
    """Name -> parameter-name sets (None for a class or an attribute), from the type stubs and Python modules of
    the installed packages and from the skills' scripts."""
    api = {}
    sources = [f for m in (rdocx, rpptx) for f in Path(m.__file__).parent.rglob("*.py*")]
    for f in sources + sorted(SKILLS.glob("*/scripts/*.py")):
        for node in ast.walk(ast.parse(f.read_text())):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                a = node.args
                names = {x.arg for x in a.posonlyargs + a.args + a.kwonlyargs} - {"self", "cls"}
                api.setdefault(node.name, []).append(names | ({"**"} if a.kwarg else set()))
            elif isinstance(node, ast.ClassDef):
                api.setdefault(node.name, []).append(None)
            elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
                api.setdefault(node.target.id, []).append(None)
            elif isinstance(node, ast.Assign):
                for t in node.targets:
                    if isinstance(t, ast.Name):
                        api.setdefault(t.id, []).append(None)
    return api


def calls(span):
    """(name, top-level arguments) of every call in a code span, nested calls included."""
    for m in re.finditer(r"(?<![\w])([A-Za-z_]\w*)\(", span):
        depth, args, cur = 0, [], ""
        for ch in span[m.end():]:
            if ch in "([{":
                depth += 1
            elif ch in ")]}":
                if depth == 0:
                    break
                depth -= 1
            if ch == "," and depth == 0:
                args.append(cur.strip())
                cur = ""
            else:
                cur += ch
        args.append(cur.strip())
        yield m.group(1), [a for a in args if a]


def python_problems(path):
    api = python_api()
    for n, span in spans(path):
        if re.match(r"(rdocx|rpptx|python3?|pip|bash|soffice) |\$", span):
            continue
        for name, args in calls(span):
            if hasattr(builtins, name) or name in FOREIGN:
                continue
            if name not in api:
                yield f"{rel(path)}:{n}: no {name} in the pinned build: `{span}`"
                continue
            sigs = [s for s in api[name] if s is not None]
            for a in args:
                kw = re.match(r"([a-z_]\w*)=(?!=)", a)
                if sigs and kw and not any(kw.group(1) in s or "**" in s for s in sigs):
                    yield f"{rel(path)}:{n}: {name}() takes no keyword {kw.group(1)}: `{span}`"


@pytest.mark.parametrize("path", DOCS, ids=lambda p: str(rel(p)))
def test_python_names_and_keywords_exist(path):
    assert list(python_problems(path)) == []


# ---------------------------------------------------------------- CLI commands and flags
@functools.lru_cache(maxsize=None)
def cli_help(*command):
    res = subprocess.run([str(BIN / (command[0] + EXE)), *command[1:], "--help"], capture_output=True, text=True)
    return res.stdout if res.returncode == 0 else None


def subcommands(help_text):
    section = help_text.split("Commands:", 1)[1].split("\n\n", 1)[0] if "Commands:" in help_text else ""
    return set(re.findall(r"^  (\S+)", section, re.M)) - {"help"}


def cli_problems(path):
    for n, span in spans(path):
        tokens = span.replace("[", " ").replace("]", " ").split()
        if not tokens or tokens[0] not in ("rdocx", "rpptx"):
            continue
        commands, i = [(tokens[0],)], 1
        while i < len(tokens):
            known = set.union(*(subcommands(cli_help(*c) or "") for c in commands))
            words = tokens[i].split("/")
            if not known or not all(w in known for w in words):
                break
            commands = [c + (w,) for c in commands for w in words if w in subcommands(cli_help(*c) or "")]
            i += 1
        if len(commands) == 1 and len(commands[0]) == 1 and len(tokens) > 1 and not tokens[1].startswith("-"):
            if re.fullmatch(r"[a-z]+", tokens[1]):
                yield f"{rel(path)}:{n}: no command {tokens[0]} {tokens[1]}: `{span}`"
                continue
        helps = [cli_help(*c) or "" for c in commands]
        for token in tokens[i:]:
            for flag in token.split("/"):
                if re.fullmatch(r"--?[A-Za-z][\w-]*", flag) and not any(
                        re.search(rf"(?<![\w-]){re.escape(flag)}(?![\w-])", h) for h in helps):
                    yield f"{rel(path)}:{n}: {' '.join(commands[0])} has no {flag}: `{span}`"


@pytest.mark.parametrize("path", DOCS, ids=lambda p: str(rel(p)))
def test_cli_commands_and_flags_exist(path):
    assert list(cli_problems(path)) == []


def test_the_checks_catch_a_wrong_name(tmp_path, monkeypatch):
    """The two checks above report nothing when they parse nothing: they must still see a wrong name."""
    monkeypatch.setattr(__import__(__name__), "rel", lambda p: p.name)
    bad = tmp_path / "bad.md"
    bad.write_text('| `doc.no_such_method()` | `a.compare(b, author, t, granularty="word")` |\n'
                   "| `rdocx compare A B --nope -o X` | `rdocx nosuch F` | `rdocx comment list/reply --bogus` |\n"
                   "| `rpptx replace F -p OLD -v NEW --expect N -o OUT` | `doc.compare(b, a, t, granularity=\"word\")` |\n")
    assert [p.split(": ", 1)[1].split(":")[0] for p in python_problems(bad)] == [
        "no no_such_method in the pinned build", "compare() takes no keyword granularty"]
    assert [p.split(": ", 1)[1].split(":")[0] for p in cli_problems(bad)] == [
        "rdocx compare has no --nope", "no command rdocx nosuch", "rdocx comment list has no --bogus"]


# ---------------------------------------------------------------- gaps
def listed_gaps(path):
    return re.findall(r"^\|[^|]*\[([a-z0-9]+(?:-[a-z0-9]+)*)\]", path.read_text(), re.M)


def test_gap_pages_list_the_gaps_of_the_registry():
    pages = sorted(SKILLS.glob("*/references/gaps.md"))
    listed = [k for p in pages for k in listed_gaps(p)]
    assert sorted(listed) == sorted(GAPS), "each key of tests/gaps.py is listed once in a gaps.md, and no other"
    for page in pages:
        if not listed_gaps(page):
            assert "No gap is open" in page.read_text(), rel(page)


def test_every_gap_has_a_test():
    tests = "".join(p.read_text() for p in (ROOT / "tests").glob("*.py") if p.name != "gaps.py")
    assert [k for k in GAPS if f'"{k}"' not in tests] == []


def test_no_gap_is_claimed_outside_the_registry():
    """A limitation called a gap anywhere in the skills names a key of tests/gaps.py, so that a test decides when
    it closes; a bare "(gap)" is how a stale claim survived a pin that had fixed it."""
    problems = []
    for path in sorted(SKILLS.rglob("*")):
        if path.suffix not in (".md", ".py") or path.name == "gaps.md":
            continue
        for n, line in enumerate(path.read_text().splitlines(), 1):
            if re.search(r"\(gaps?\)", line):
                problems.append(f"{rel(path)}:{n}: a gap without its key")
            problems += [f"{rel(path)}:{n}: no gap {k} in tests/gaps.py"
                         for k in re.findall(r"\bgap ([a-z0-9]+(?:-[a-z0-9]+)+)", line) if k not in GAPS]
    assert problems == []


# ---------------------------------------------------------------- size
@pytest.mark.parametrize("path", sorted(SKILLS.glob("*/SKILL.md")), ids=lambda p: p.parent.name)
def test_skill_md_stays_under_200_lines(path):
    assert len(path.read_text().splitlines()) < 200
