---
name: "commit-github"
description: "Commits and pushes project to GitHub remote. Invoke when user asks to commit, push, or submit code to GitHub."
---

# Commit & Push to GitHub

This skill commits all changes in the current project and pushes them to the configured GitHub remote repository.

## Target Remote

- URL: `https://github.com/chen0601-xin/chenxinyang.git`
- SSH port 22 is blocked in this environment; use HTTPS with `http.sslVerify=false` (already configured locally).
- SSH auth via `git@github.com:chen0601-xin/chenxinyang.git` will fail with "Connection refused" on port 22.

## Steps to Execute

1. **Check git status**: Run `git status` to see current state. If git is not initialized, run `git init` first.

2. **Configure remote**: If no remote is set, add it:
   ```
   git remote add origin https://github.com/chen0601-xin/chenxinyang.git
   ```
   If remote already exists but URL differs, update it:
   ```
   git remote set-url origin https://github.com/chen0601-xin/chenxinyang.git
   ```

3. **Add all files**: `git add .`

4. **Commit**: Use a descriptive commit message. If there are no prior commits, use an initial message like `chore: initial commit`. Subsequent commits should describe what changed.
   ```
   git commit -m "your commit message"
   ```

5. **Push**: Check current branch name first (`git branch --show-current`), then push. Default branch may be `master` or `main`.
   ```
   git push -u origin <branch-name>
   ```

## Notes

- Always check `git status` before committing to avoid committing unwanted files.
- If there are files that should NOT be committed (e.g. sensitive data, build artifacts), suggest creating a `.gitignore` file first.
- If push fails with a non-fast-forward error, inform the user that the remote has diverged and they may need to pull first.
