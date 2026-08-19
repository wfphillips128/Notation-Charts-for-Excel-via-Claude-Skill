"""Where things are, when the caller does not say.

Every default here is the **current working directory**, not a path derived
from this file. That is deliberate. A skill is installed under
``~/.claude/skills/`` and a repository is cloned wherever the user likes; in
both cases writing build output next to the source would scatter artefacts
through somebody's configuration or their checkout. Running the scripts from
the directory you want the output in is the behaviour everybody already
expects from a command-line tool.

Each default can be overridden by an environment variable, which is what makes
the same scripts usable from a scheduled job or a different checkout without
editing them.

    IBCS_BUILD           where workbooks, SVGs and images are written
    IBCS_TEMPLATE_REFS   the IBCS(R) reference renders, for comparison only

**The reference renders are not distributed with this project.** They are the
IBCS Institute's own images and are used here only to measure a recreation
against its source. ``compare_render.py`` and ``extract_palette.py`` are the
only two scripts that need them, and both say so when the directory is absent
rather than failing obscurely.
"""

from __future__ import annotations

import os
from pathlib import Path


def build_dir(name: str = "build") -> Path:
    """Where output goes. ``$IBCS_BUILD``, else ``./build``."""
    override = os.environ.get("IBCS_BUILD")
    return Path(override) if override else Path.cwd() / name


def refs_dir() -> Path:
    """The IBCS(R) reference renders. ``$IBCS_TEMPLATE_REFS``, else
    ``./template-refs``. Not shipped - see the module docstring."""
    override = os.environ.get("IBCS_TEMPLATE_REFS")
    return Path(override) if override else Path.cwd() / "template-refs"


def require_refs() -> Path:
    """The reference directory, or a message saying why there isn't one."""
    path = refs_dir()
    if not path.is_dir():
        raise SystemExit(
            "error: no reference renders at %s.\n"
            "  These are the IBCS(R) published template images and are not\n"
            "  distributed with this project. Point IBCS_TEMPLATE_REFS at your\n"
            "  own copy, or skip the comparison scripts - nothing else needs\n"
            "  them." % path)
    return path


def panel_charts_scripts() -> Path | None:
    """Where the companion ``panel-charts`` skill's scripts are, if anywhere.

    C13 is a grid of small multiples and Excel has no panel chart type, so that
    one sheet is drawn by the companion skill rather than reinvented here. It is
    looked for in three places, in this order, and the first that exists wins:

      1. ``$PANEL_CHARTS`` - a checkout anywhere;
      2. a sibling of this skill, which is how ``~/.claude/skills`` lays out;
      3. that same sibling under the user's home, for a project-local install.

    Returns None rather than raising. Sixteen of the seventeen templates do not
    need it, so a missing companion must cost exactly one sheet - and pushing a
    directory that does not exist onto ``sys.path`` to find that out is how an
    import error ends up blaming the wrong module.
    """
    here = Path(__file__).resolve()
    for candidate in (os.environ.get("PANEL_CHARTS"),
                      here.parents[2] / "panel-charts" / "scripts",
                      Path.home() / ".claude" / "skills" / "panel-charts"
                      / "scripts"):
        if candidate and Path(candidate).is_dir():
            return Path(candidate)
    return None
