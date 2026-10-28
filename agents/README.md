# Agent skills for docxsphinx

This directory holds skills for coding agents. A skill is a folder whose `SKILL.md` an agent
(Claude Code, or any tool that reads the `SKILL.md` format) loads on demand when the task
matches. They are not part of the installed Python package.

| Skill | Audience | Loaded when |
|---|---|---|
| [`docxsphinx-user`](skills/docxsphinx-user/SKILL.md) | People **using** docxsphinx in their own Sphinx project | Setting up `conf.py`, templates and styles, multiple outputs, `docx2md` / `docx2rst`, or working out why something is missing from the `.docx` |
| [`docxsphinx-developer`](skills/docxsphinx-developer/SKILL.md) | People **changing** docxsphinx | Editing `src/docxsphinx/`, adding tests, examples or goldens, debugging the writer, releasing |

## Layout

```
agents/
└── skills/
    ├── docxsphinx-user/
    │   ├── SKILL.md                         entry point: front-matter + core guidance
    │   └── references/feature-support.md    full what-renders / what-drops matrix
    └── docxsphinx-developer/
        ├── SKILL.md
        └── references/ooxml-and-python-docx.md
```

Each skill is a directory. The `name` and `description` in the front-matter of `SKILL.md` are all
an agent sees before it decides to load the skill, so the description has to say when to use
it. Detail that's needed only sometimes lives under `references/`, and `SKILL.md` links to it by
relative path, so it costs nothing until it's opened. To extend a skill, add a file under
`references/` (or `scripts/` for runnable helpers) and link it from `SKILL.md`. Keep `SKILL.md`
itself short. Add a new sibling directory only for a new audience or task.

## Installing

Claude Code finds skills in `~/.claude/skills/<name>/` (every project) or
`<project>/.claude/skills/<name>/` (one project). Symlink from a checkout so the skills stay
current:

```bash
# from a docxsphinx checkout, for all your projects
mkdir -p ~/.claude/skills
ln -s "$PWD/agents/skills/docxsphinx-user"      ~/.claude/skills/docxsphinx-user
ln -s "$PWD/agents/skills/docxsphinx-developer" ~/.claude/skills/docxsphinx-developer

# or copy the user skill into a single Sphinx project that builds .docx
cp -r /path/to/docxsphinx/agents/skills/docxsphinx-user  <project>/.claude/skills/
```

For other agents, point them at the `SKILL.md`, or paste its body into their instruction file
(for example `AGENTS.md`).

## Keeping them accurate

The skills describe the code at a particular version (2.1 as of writing). If a change affects
what renders, what a config value does, a Make target, a test tier or the release process, update
the skill that covers it in the same change. The developer skill's "Known gaps" list and the
user skill's feature matrix are the parts most likely to go stale.
