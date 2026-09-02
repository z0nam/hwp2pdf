#!/bin/sh
# Warn when this checkout is behind origin.
#
# hwp2pdf is edited from two machines (a Mac and namun-ji), and starting work on
# a stale base already cost one 322-line hand merge. Claude Code runs this from
# SessionStart and before the first Edit/Write so `git pull` never gets skipped.
#
#   check_freshness.sh session   fetch, then report (SessionStart)
#   check_freshness.sh guard     report, refetching at most every 5 min (PreToolUse)
set -u

REPO=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd) || exit 0
cd "$REPO" || exit 0
GIT_DIR=$(git rev-parse --git-dir 2>/dev/null) || exit 0

MODE=${1:-session}
STAMP="$GIT_DIR/hwp2pdf-freshness"
now=$(date +%s)
last=0
[ -f "$STAMP" ] && last=$(cat "$STAMP" 2>/dev/null || echo 0)

# A fetch per edit would stall the session, so the guard reuses a recent one.
if [ "$MODE" = session ] || [ $((now - last)) -ge 300 ]; then
    git fetch --quiet --no-tags origin 2>/dev/null
    echo "$now" > "$STAMP" 2>/dev/null
fi

upstream=$(git rev-parse --abbrev-ref --symbolic-full-name '@{u}' 2>/dev/null) || exit 0
[ -n "$upstream" ] || exit 0
branch=$(git rev-parse --abbrev-ref HEAD 2>/dev/null)

counts=$(git rev-list --left-right --count "$upstream...HEAD" 2>/dev/null) || exit 0
behind=$(printf '%s' "$counts" | awk '{print $1+0}')
ahead=$(printf '%s' "$counts" | awk '{print $2+0}')
[ "$behind" -gt 0 ] || exit 0

if git diff --quiet 2>/dev/null && git diff --cached --quiet 2>/dev/null; then
    advice="Run 'git pull' before changing anything."
else
    advice="Commit or stash the working tree first, then 'git pull'."
fi
msg="hwp2pdf: $branch is $behind commit(s) behind $upstream (ahead $ahead). $advice The repo is also edited on the other machine, so a stale base means a hand merge later."

# printf, not a heredoc: the message is one line of ASCII with no JSON metacharacters.
if [ "$MODE" = guard ]; then
    printf '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"ask","permissionDecisionReason":"%s"}}\n' "$msg"
else
    printf '{"systemMessage":"%s","hookSpecificOutput":{"hookEventName":"SessionStart","additionalContext":"%s"}}\n' "$msg" "$msg"
fi
