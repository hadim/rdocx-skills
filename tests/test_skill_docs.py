"""The skills against the pinned build: every Python name and keyword argument and every CLI command and flag they
cite exists, every public Python name, CLI command and flag of the build is cited or excluded with a reason, their gap
pages match tests/gaps.py, and each SKILL.md stays under 200 lines. A pin that renames, drops or adds an API fails
here even when no recipe block runs it (the recipes themselves run in test_docs_snippets.py).

Positional parameters are written as short placeholders in the skills (`bi`, `rid`, `r, c`) and are not checked:
a positional call works whatever the parameter is called. Keyword arguments are checked, since a wrong one raises."""
import ast
import builtins
import functools
import importlib
import pkgutil
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
    # .py and .pyi only: "*.py*" also matches the extension module on Windows (.pyd), a binary
    sources = [f for m in (rdocx, rpptx) for f in Path(m.__file__).parent.rglob("*.py*") if f.suffix in (".py", ".pyi")]
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


# ---------------------------------------------------------------- the other way: every API is documented or excluded
# Every public class of rdocx and rpptx (all their modules), its public methods and properties, and every CLI command
# and flag must be cited in code in that skill's SKILL.md or references, or be listed here with the reason it is left
# out. Not scanned: names with a leading underscore (dunder, internal), members a class inherits from a builtin base
# (int, Exception), and ALL_CAPS enum values (covered by their enum). A class reachable under several names (python-pptx
# aliases such as MSO_FILL_TYPE for MSO_FILL) is covered when one of them is cited. A command group is covered by its
# subcommands; -h/--help and -V/--version are not scanned. A test fails on a new name a pin brings, and on an entry here
# that is no longer needed.
NOT_DOCUMENTED = {
    # sequences returned by attributes the skills document (`doc.paragraphs`, `row.cells`...): indexed, iterated and
    # measured with len(), never named or built
    "rdocx.CellCollection": "the type of `row.cells`, used as a sequence",
    "rdocx.CellParagraphCollection": "the type of `cell.paragraphs`, used as a sequence",
    "rdocx.ParagraphCollection": "the type of `doc.paragraphs`, used as a sequence",
    "rdocx.RowCollection": "the type of `table.rows`, used as a sequence",
    "rdocx.RunCollection": "the type of `paragraph.runs`, used as a sequence",
    "rdocx.TableCollection": "the type of `doc.tables`, used as a sequence",
    "rpptx.AdjustmentCollection": "the type of `shape.adjustments`, used as a sequence",
    "rpptx.ColumnCollection": "the type of `table.columns`, documented by its members",
    "rpptx.ParagraphCollection": "the type of `text_frame.paragraphs`, used as a sequence",
    "rpptx.PlaceholderCollection": "the type of `shapes.placeholders`, used as a sequence",
    "rpptx.RowCollection": "the type of `table.rows`, documented by its members",
    "rpptx.RunCollection": "the type of `paragraph.runs`, used as a sequence",
    "rpptx.SlideCollection": "the type of `prs.slides`, documented by its members",
    "rpptx.SlideLayoutCollection": "the type of `prs.slide_layouts`, documented by its members",
    # handles reached through a documented attribute, which the skills describe by their members
    "rpptx.Background": "the type of `slide.background`, documented by its members",
    "rpptx.Cell": "the type of `table.cell(r, c)`, documented by its members",
    "rpptx.ColorFormat": "the type of `fore_color` and `line.color`, documented by `rgb`",
    "rpptx.Column": "the type of `table.columns[k]`, documented by its members",
    "rpptx.Font": "the type of `run.font` and `paragraph.font`, documented by its members",
    "rpptx.Hyperlink": "the type of `run.hyperlink`, documented by `address`",
    "rpptx.Image": "the type of `shape.image`, documented by its members",
    "rpptx.Row": "the type of `table.rows[k]`, documented by its members",
    "rpptx.ShapeClickAction": "the type of `shape.click_action`, documented by its members",
    "rpptx.ShapeHyperlink": "the type of `click_action.hyperlink`, documented by `address`",
    "rpptx.SlideLayout": "the type of `prs.slide_layouts[k]` and `slide.slide_layout`, documented by its use",
}
SKILL_OF = {"rdocx": "docx", "rpptx": "pptx"}


@functools.lru_cache(maxsize=None)
def skill_code(skill):
    """The code of a skill's SKILL.md and references: inline spans (also those that wrap a line) and fenced blocks."""
    out = []
    for path in [SKILLS / skill / "SKILL.md", *sorted((SKILLS / skill / "references").glob("*.md"))]:
        for k, part in enumerate(re.split(r"^```.*$", path.read_text(), flags=re.M)):
            out += [part] if k % 2 else re.findall(r"`([^`]+)`", part)
    return "\n".join(out)


def python_surface(module):
    """{key: (names that cite it)}: "pkg.Class" for each class, "pkg.Class.member" for its public members."""
    classes = {}
    for info in [None, *pkgutil.walk_packages(module.__path__, module.__name__ + ".")]:
        mod = importlib.import_module(info.name) if info else module
        for name, obj in vars(mod).items():
            if (not name.startswith("_") and isinstance(obj, type) and obj not in vars(builtins).values()
                    and obj.__module__.split(".")[0] in (module.__name__, "builtins")):
                classes.setdefault(obj, set()).add(name)
    surface = {}
    for obj, names in classes.items():
        base = [b for b in obj.__mro__ if b in vars(builtins).values()]
        surface[f"{module.__name__}.{obj.__name__}"] = names
        for member in dir(obj):
            if not (member.startswith("_") or member.isupper() or any(hasattr(b, member) for b in base)):
                surface[f"{module.__name__}.{obj.__name__}.{member}"] = {member}
    return surface


def undocumented_python(surface, code, excluded):
    words = set(re.findall(r"[A-Za-z_]\w*", code))
    return sorted(k for k, names in surface.items() if not names & words and k not in excluded)


@functools.lru_cache(maxsize=None)
def cli_surface(tool):
    """{(tool, command, ...): [flag aliases such as ("-o", "--output")]} for every command and subcommand."""
    surface = {}

    def walk(command):
        text = cli_help(*command) or ""
        flags = []
        for line in text.split("Options:", 1)[1].splitlines() if "Options:" in text else []:
            m = re.match(r"\s+(-\w)?(?:, )?(--[\w-]+)?", line)
            names = tuple(x for x in m.groups() if x) if m else ()
            if names and not set(names) & {"-h", "--help", "-V", "--version"}:
                flags.append(names)
        surface[command] = flags
        for sub in sorted(subcommands(text)):
            walk(command + (sub,))
    walk((tool,))
    return surface


def cli_citations(code, tool, surface):
    """{command: flags} for every invocation of `tool` in the code (`a/b` alternatives and `[...]` options read)."""
    lines = []
    for line in code.splitlines():
        if lines and re.match(r"\s+[\[-]", line):          # an option list continued on the next line
            lines[-1] += " " + line
        else:
            lines.append(line)
    cited = {}
    for line in lines:
        for m in re.finditer(rf"(?:^|[\s/(]){tool}((?:\s+[^\s#`]+)*)", line):
            tokens = m.group(1).replace("[", " ").replace("]", " ").split()
            paths, i = [(tool,)], 0
            while i < len(tokens):
                deeper = [p + (w,) for p in paths for w in tokens[i].split("/") if p + (w,) in surface]
                if not deeper:
                    break
                paths, i = deeper, i + 1
            for p in paths:
                cited.setdefault(p, set()).update(f for t in tokens[i:] for f in t.split("/") if f.startswith("-"))
    return cited


def undocumented_cli(surface, cited, excluded):
    missing = []
    for command, flags in surface.items():
        name = " ".join(command)
        if len(command) > 1 and not any(c[:len(command)] == command for c in cited) and name not in excluded:
            missing.append(name)
        missing += [f"{name} {f[-1]}" for f in flags                # keyed by the long name
                    if not set(f) & cited.get(command, set()) and f"{name} {f[-1]}" not in excluded]
    return missing


@pytest.mark.parametrize("module", [rdocx, rpptx], ids=lambda m: m.__name__)
def test_every_python_api_is_documented_or_excluded(module):
    code = skill_code(SKILL_OF[module.__name__])
    assert undocumented_python(python_surface(module), code, NOT_DOCUMENTED) == []


@pytest.mark.parametrize("tool", ["rdocx", "rpptx"])
def test_every_cli_command_and_flag_is_documented_or_excluded(tool):
    surface = cli_surface(tool)
    assert undocumented_cli(surface, cli_citations(skill_code(SKILL_OF[tool]), tool, surface), NOT_DOCUMENTED) == []


def test_every_exclusion_is_still_needed():
    """An entry names a scanned name that the skills still do not cite: drop it once the name is documented or gone."""
    missing = set()
    for tool in SKILL_OF:
        code, surface = skill_code(SKILL_OF[tool]), cli_surface(tool)
        missing.update(undocumented_python(python_surface(importlib.import_module(tool)), code, {}))
        missing.update(undocumented_cli(surface, cli_citations(code, tool, surface), {}))
    assert sorted(set(NOT_DOCUMENTED) - missing) == []


def test_the_reverse_checks_catch_an_undocumented_name():
    """A fake class, member, command and flag that no skill cites are reported; an excluded one is not."""
    surface = {**python_surface(rdocx), "rdocx.NoSuchClass": {"NoSuchClass"}, "rdocx.Document.no_such_member": {"no_such_member"}}
    code = skill_code("docx")
    assert undocumented_python(surface, code, NOT_DOCUMENTED) == ["rdocx.Document.no_such_member", "rdocx.NoSuchClass"]
    assert undocumented_python(surface, code, {**NOT_DOCUMENTED, "rdocx.NoSuchClass": "x"}) == ["rdocx.Document.no_such_member"]
    cli = {**cli_surface("rdocx"), ("rdocx", "nosuch"): [("--zap",)]}
    cli[("rdocx", "text")] = cli[("rdocx", "text")] + [("-Z", "--zebra")]
    cited = cli_citations(code, "rdocx", cli)
    assert undocumented_cli(cli, cited, NOT_DOCUMENTED) == ["rdocx text --zebra", "rdocx nosuch", "rdocx nosuch --zap"]


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
