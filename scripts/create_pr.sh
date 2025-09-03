#!/usr/bin/env bash
set -euo pipefail
BRANCH=${1?"branch name"}
REMOTE=${2:-origin}
REPO=${3:-}

git push "$REMOTE" "$BRANCH"
if [ -n "$REPO" ]; then
  gh pr create --repo "$REPO" --head "$BRANCH" --base main --fill
else
  gh pr create --head "$BRANCH" --base main --fill
fi
