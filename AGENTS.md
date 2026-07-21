# Instructions for AI coding agents

## Project purpose

This repository contains small, practical tools that automate repetitive consultant work. Keep each tool easy to understand, safe to run, and useful to someone who has never written code.

## Working principles

- Preserve client confidentiality. Never commit workbooks, client names, credentials, access tokens, or real client data.
- Treat input files as read-only. A tool must write to a new file or folder and must not overwrite an input unless the user explicitly requests that behavior.
- Prefer clear command-line options over hard-coded machine-specific paths when improving a tool.
- Use plain language in messages and documentation. Explain every command a beginner must type.
- Keep tools small and independent. Avoid adding a framework unless it provides a clear benefit.
- Use `pathlib.Path` for file paths so tools remain portable.
- Report what the tool is doing, where output was written, and how many files or rows were processed.
- Fail with a helpful message when a file, folder, worksheet, or column cannot be found.

## Code quality

- Support currently maintained Python 3 versions; use Python 3.11 or newer for local development.
- Add type hints and short docstrings to reusable functions.
- Put executable behavior behind a `main()` function and an `if __name__ == "__main__":` guard.
- Keep third-party dependencies minimal and record them in `requirements.txt`.
- Format Python consistently and use descriptive names rather than abbreviations.
- Do not silently discard data. Preserve column order and clearly document which worksheets or rows a tool reads.

## Testing and verification

- Add or update automated tests when behavior changes.
- Use synthetic workbooks only; test fixtures must not contain client data.
- Test empty inputs, missing columns, duplicate names, blank values, unsafe filename characters, and existing output files where relevant.
- Before committing, run the relevant tool against a small synthetic workbook and confirm that the output opens successfully.
- Run `python -m compileall .` as a basic syntax check.

## Documentation

- Keep `README.md` accurate whenever commands, filenames, dependencies, defaults, or outputs change.
- Put beginner instructions first, including installation, exact commands, expected output, and common errors.
- Include examples that use generic names such as `input.xlsx`; never use a real client or engagement name.

## Git workflow

- Make focused commits: one logical change per commit.
- Write a concise subject followed by a detailed body explaining what changed, why, and how it was verified.
- Do not commit generated Excel files, virtual environments, caches, secrets, or temporary databases.
- Review `git status` and the staged diff before every commit.

