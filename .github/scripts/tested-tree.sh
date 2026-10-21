#!/bin/sh
# Whether ci.yml already passed on the tree of HEAD. A ci.yml run whose four suites passed uploads the artifact
# `tested-tree-<tree>` (job `tested`), the tree being the one it checked out (for a pull request, its merge into the
# base). A push to main that merges such a pull request, main unmoved, has that very tree: ci.yml runs nothing again
# and release.yml releases it at once. The mark counts only from a finished green run of ci.yml itself, on a push or
# a pull request of this repository: not from a fork's pull request, whose workflow the fork could change. Prints
# `tree=<sha>`, `run=<id>` (empty when none) and `passed=true|false` for $GITHUB_OUTPUT, and the run that passed on stderr. Needs git, gh, GH_TOKEN
# (actions: read) and GITHUB_REPOSITORY.
set -eu
tree=$(git rev-parse 'HEAD^{tree}')
runs=$(gh api "repos/$GITHUB_REPOSITORY/actions/artifacts?name=tested-tree-$tree&per_page=100" --jq \
  '.artifacts[] | select((.expired | not) and .workflow_run.head_repository_id == .workflow_run.repository_id)
   | .workflow_run.id')
run=""
for id in $runs; do  # the mark counts only from ci.yml itself, on a push or a pull request, and a run that passed
  ok=$(gh api "repos/$GITHUB_REPOSITORY/actions/runs/$id" --jq \
    '.path == ".github/workflows/ci.yml" and (.event == "push" or .event == "pull_request") and .conclusion == "success"')
  if [ "$ok" = true ]; then run=$id; break; fi
done
echo "tree=$tree"
echo "run=$run"
if [ -n "$run" ]; then
  echo "passed=true"
  echo "ci.yml run $run passed on tree $(printf %.12s "$tree")" >&2
else
  echo "passed=false"
fi
