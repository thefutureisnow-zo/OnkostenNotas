# Git Workflow: Branches + Pull Requests

**All fixes and new features must be developed on a separate branch and merged via pull request.**
Never commit directly to `main`.

## Creating a branch

```bash
git checkout -b fix/description-of-fix   # bug fix
git checkout -b feat/description          # new feature
```

## After making changes

1. Run `python3 -m pytest -v` -- all tests must pass.
2. Commit on the branch.
3. Push and open a PR:

```bash
git push -u origin <branch-name>
gh pr create --title "..." --body "..."
```

## PR validation (sub-agent)

After opening a PR, always launch a validation sub-agent with the following instructions:

> "Review the open PR at [PR URL]. Check out the branch, run `python3 -m pytest -v`, confirm all
> tests pass, read the changed files, and report: (1) test results, (2) any logic issues or edge
> cases not covered by the tests, (3) whether the PR is safe to merge."

Only merge after the sub-agent confirms tests pass and raises no blocking issues.

## Merge

```bash
gh pr merge <PR-number> --squash --delete-branch
```
