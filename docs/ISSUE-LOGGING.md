# Issue Logging

Every bug fix or non-trivial feature must have a log file at:

```
issues/<issue-name>/<issue-name>.md
```

The file must cover four sections:

1. **The problem** -- what the user observed and what the actual error was
2. **How we spotted it** -- the quickest path to finding the root cause (commands run, tracebacks seen)
3. **Root cause** -- the underlying reason, not just the symptom
4. **Fix** -- what was changed and why, with the PR link

Keep it concise. The goal is that you or the user can re-read it months later and immediately understand
what happened and why the fix works.
