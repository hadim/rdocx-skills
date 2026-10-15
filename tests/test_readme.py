"""The two copies of the README: plugin/README.md is what the plugin ships and the plugin directory lists, README.md
is the repository's page. They differ only in where the pictures are, relative to each file."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def test_the_repository_readme_is_the_plugin_readme_with_its_pictures_under_plugin():
    plugin = (ROOT / "plugin" / "README.md").read_text(encoding="utf-8")
    assert "](plugin/" not in plugin
    assert (ROOT / "README.md").read_text(encoding="utf-8") == plugin.replace("](assets/", "](plugin/assets/"), (
        "edit plugin/README.md, then: sed 's|](assets/|](plugin/assets/|g' plugin/README.md > README.md")
