#!/bin/bash
# Ralph Docker Loop
# Runs Claude in isolated container, backup/git runs on HOST

set -e

# Configuration
IMAGE_NAME="ralph-loop"
PROJECT_DIR="$(pwd)"
PROJECT_NAME=$(basename "$PROJECT_DIR")
BACKUP_ENABLED="${RALPH_BACKUP:-true}"
MODEL="${RALPH_MODEL:-sonnet}"

# Validate model against whitelist (security: prevents command injection)
validate_model() {
  local model="$1"
  case "$model" in
    opus|sonnet|haiku) return 0 ;;
    *)
      echo "Error: Invalid model '$model'. Allowed: opus, sonnet, haiku"
      exit 1
      ;;
  esac
}
validate_model "$MODEL"
PLAN_FILE="IMPLEMENTATION_PLAN.md"
LOG_FILE="ralph.log"
VBA_TEST_MARKER=".vba_tested_commit"
VBA_RESULTS_FILE="vba/tests/LAST_TEST_RUN.md"

# SAFETY: Verify PROJECT_DIR is safe to mount
if [ -z "$PROJECT_DIR" ] || [ "$PROJECT_DIR" = "/" ] || [ "$PROJECT_DIR" = "$HOME" ]; then
  echo "FATAL: Refusing to mount unsafe directory: $PROJECT_DIR"
  echo "Run this script from inside a project directory, not ~ or /"
  exit 1
fi

# Verify we're in a Ralph project
if [ ! -f "PROMPT_build.md" ] && [ ! -f "PROMPT_plan.md" ]; then
  echo "FATAL: Not a Ralph project directory (no PROMPT_*.md files)"
  echo "Run /setup-ralph first or cd into a Ralph project"
  exit 1
fi

# Load OAuth token (with security checks)
TOKEN_FILE="$HOME/.claude-oauth-token"
if [ -z "$CLAUDE_CODE_OAUTH_TOKEN" ]; then
  if [ -f "$TOKEN_FILE" ]; then
    # Security: Check file permissions (should be 600 or more restrictive)
    if [[ "$OSTYPE" == "darwin"* ]]; then
      TOKEN_PERMS=$(stat -f %Lp "$TOKEN_FILE" 2>/dev/null)
    else
      TOKEN_PERMS=$(stat -c %a "$TOKEN_FILE" 2>/dev/null)
    fi

    if [ -n "$TOKEN_PERMS" ] && [ "$((TOKEN_PERMS % 100))" -ne 0 ]; then
      echo "⚠️  Security warning: $TOKEN_FILE has insecure permissions ($TOKEN_PERMS)"
      echo "   Run: chmod 600 $TOKEN_FILE"
      echo ""
    fi

    CLAUDE_CODE_OAUTH_TOKEN=$(cat "$TOKEN_FILE")
  else
    echo "Error: No OAuth token found"
    echo "Run 'claude setup-token' and save to ~/.claude-oauth-token"
    echo "Then: chmod 600 ~/.claude-oauth-token"
    exit 1
  fi
fi

# ============================================================================
# DOCKER HEALTH CHECK
# ============================================================================
# WSL2 + Docker Desktop occasionally drops the Linux-side `docker` CLI from
# this distro (a WSL-integration hiccup, not a real engine crash -- Docker
# Desktop itself and the underlying WSL VMs stay up). Rather than stopping to
# ask the user to re-toggle Settings -> Resources -> WSL Integration each
# time, nudge Docker Desktop's own engine via its Windows-side CLI plugin
# (`docker.exe desktop restart`), which works even while the broken Linux
# symlink can't, then poll until the Linux CLI comes back.
WINDOWS_DOCKER_CLI="/mnt/c/Program Files/Docker/Docker/resources/bin/docker.exe"

ensure_docker_available() {
  if docker version >/dev/null 2>&1; then
    return 0
  fi

  echo "docker CLI unavailable -- attempting self-heal via Docker Desktop restart..."

  if [ ! -x "$WINDOWS_DOCKER_CLI" ]; then
    echo "FATAL: docker unavailable and no Windows-side Docker Desktop CLI found at:"
    echo "  $WINDOWS_DOCKER_CLI"
    echo "(expected only on WSL2 + Docker Desktop -- on native Linux/Mac, check the docker daemon directly)"
    return 1
  fi

  "$WINDOWS_DOCKER_CLI" desktop restart >/dev/null 2>&1 || true

  local attempt
  for attempt in $(seq 1 30); do
    sleep 2
    if docker version >/dev/null 2>&1; then
      echo "docker CLI recovered after Docker Desktop restart (waited ~$((attempt * 2))s)."
      return 0
    fi
  done

  echo "FATAL: docker still unavailable $((attempt * 2))s after a Docker Desktop restart attempt."
  echo "Check Docker Desktop -> Settings -> Resources -> WSL Integration manually."
  return 1
}

ensure_docker_available || exit 1

# Handle --build-image flag
if [ "$1" = "--build-image" ]; then
  echo "Building Docker image..."
  docker build -t "$IMAGE_NAME" .
  echo "Image built: $IMAGE_NAME"
  exit 0
fi

# Check if image exists
if ! docker image inspect "$IMAGE_NAME" &>/dev/null; then
  echo "Docker image not found. Building..."
  docker build -t "$IMAGE_NAME" .
fi

# Parse arguments
MODE="build"
LIMIT=""
while [[ $# -gt 0 ]]; do
  case $1 in
    plan) MODE="plan"; shift ;;
    [0-9]*) LIMIT=$1; shift ;;
    --model) MODEL=$2; validate_model "$MODEL"; shift 2 ;;
    *) shift ;;
  esac
done

# ============================================================================
# REMOTE BACKUP (runs on HOST with your gh auth)
# ============================================================================

setup_remote_backup() {
  if [ "$BACKUP_ENABLED" != "true" ]; then
    echo "Remote backup: disabled (set RALPH_BACKUP=true to enable)"
    return 0
  fi

  if [ ! -d ".git" ]; then
    echo "Initializing git..."
    git init
    git add -A
    git commit -m "Initial commit" 2>/dev/null || true
  fi

  if git remote get-url origin &>/dev/null; then
    echo "Remote backup: $(git remote get-url origin)"
    return 0
  fi

  if ! command -v gh &>/dev/null; then
    echo "Warning: gh CLI not found. Backup disabled."
    BACKUP_ENABLED="false"
    return 1
  fi

  if ! gh auth status &>/dev/null; then
    echo "Warning: gh not authenticated. Backup disabled."
    BACKUP_ENABLED="false"
    return 1
  fi

  local repo_name="${PROJECT_NAME}-ralph-backup"
  echo "Creating private backup: $repo_name"

  if gh repo create "$repo_name" --private --source=. --push 2>/dev/null; then
    echo "Remote backup: https://github.com/$(gh api user -q .login)/$repo_name"
  else
    echo "Warning: Could not create repo. Backup disabled."
    BACKUP_ENABLED="false"
  fi
}

push_to_backup() {
  if [ "$BACKUP_ENABLED" = "true" ]; then
    git add -A 2>/dev/null || true
    git diff --quiet HEAD 2>/dev/null || git commit -m "Auto-save after iteration $ITERATION" 2>/dev/null || true
    git push origin HEAD 2>/dev/null && echo "Pushed to backup" || echo "Push failed (continuing)"
  fi
}

# ============================================================================
# VBA TEST BRIDGE (runs on HOST -- real Office via COM, container can't do this)
# ============================================================================
# Ralph's container has no Windows/Office install, so VBA changes it commits
# are never actually executed, only read. This runs the real test harness
# (vba/tests/run_vba_tests.ps1) on the HOST via powershell.exe against every
# vba/*.bas change since the last tested commit, and commits the report back
# into the repo so the next iteration reads real pass/fail instead of
# "not executed in this environment."
#
# run_vba_tests.ps1 itself refuses to force-close Office if you have unsaved
# work open (asks Quit() nicely, which triggers Office's own save prompt, and
# exits 2 -- "skip" -- if that isn't resolved within its timeout). This
# function just needs to tell a skip apart from a real driver failure.
run_vba_test_bridge_if_needed() {
  local last_tested
  if [ -f "$VBA_TEST_MARKER" ]; then
    last_tested=$(cat "$VBA_TEST_MARKER")
  else
    last_tested=$(git rev-list --max-parents=0 HEAD | tail -1)
  fi

  if ! git cat-file -e "$last_tested" 2>/dev/null; then
    last_tested=$(git rev-list --max-parents=0 HEAD | tail -1)
  fi

  local changed
  changed=$(git diff --name-only "$last_tested" HEAD -- 'vba/*.bas' 2>/dev/null)
  if [ -z "$changed" ]; then
    return 0
  fi

  echo "VBA changed since last real-Office test run:"
  echo "$changed" | sed 's/^/  /'
  echo "Running vba/tests/run_vba_tests.ps1 on host..."

  local win_script
  win_script=$(wslpath -w "$PROJECT_DIR/vba/tests/run_vba_tests.ps1")
  local report rc
  report=$(powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$win_script" 2>&1)
  rc=$?

  if [ "$rc" -eq 2 ]; then
    echo "VBA test bridge skipped this iteration (Office in use / unresolved save prompt) -- will retry next commit."
    echo "$report" | tail -5
    return 0
  fi

  local sha ts
  sha=$(git rev-parse HEAD)
  ts=$(date '+%Y-%m-%d %H:%M:%S %z')

  {
    echo "## $ts -- commit $sha"
    echo
    [ "$rc" -ne 0 ] && echo "DRIVER FAILED (exit $rc):"
    echo '```'
    echo "$report"
    echo '```'
    echo
  } >> "$VBA_RESULTS_FILE"
  echo "$sha" > "$VBA_TEST_MARKER"

  git add "$VBA_RESULTS_FILE" "$VBA_TEST_MARKER"
  git commit -m "[vba-test-bridge] real-Office run against ${sha:0:7}" >/dev/null 2>&1 || true
  echo "VBA test bridge: results appended to $VBA_RESULTS_FILE"
}

# ============================================================================
# COMPLETION DETECTION (runs on HOST)
# ============================================================================

check_complete() {
  if [ ! -f "$PLAN_FILE" ]; then
    return 1
  fi

  local incomplete=$(grep -c '^\s*- \[ \]' "$PLAN_FILE" 2>/dev/null || echo "0")
  if [ "$incomplete" -eq 0 ]; then
    local completed=$(grep -c '^\s*- \[x\]' "$PLAN_FILE" 2>/dev/null || echo "0")
    [ "$completed" -gt 0 ] && return 0
  fi
  return 1
}

# ============================================================================
# MAIN
# ============================================================================

# Select prompt
if [ "$MODE" = "plan" ]; then
  PROMPT_FILE="PROMPT_plan.md"
  echo "Ralph Planning Mode (Docker)"
else
  PROMPT_FILE="PROMPT_build.md"
  echo "Ralph Building Mode (Docker)"
fi

if [ ! -f "$PROMPT_FILE" ]; then
  echo "Error: $PROMPT_FILE not found"
  exit 1
fi

echo "Project: $PROJECT_DIR"
echo "Model: $MODEL"
[ -n "$LIMIT" ] && echo "Limit: $LIMIT" || echo "Limit: until complete"
echo ""

setup_remote_backup
echo ""
echo "Starting loop..."
echo "---"

echo "=== Ralph Docker $(date '+%Y-%m-%d %H:%M:%S') ===" > "$LOG_FILE"

ITERATION=0
FAST_STREAK=0
while true; do
  ITERATION=$((ITERATION + 1))

  echo ""
  echo "Iteration $ITERATION - $(date '+%H:%M:%S')"

  # Check completion (build mode only)
  if [ "$MODE" = "build" ] && check_complete; then
    echo "ALL TASKS COMPLETE"
    push_to_backup
    exit 0
  fi

  # Check limit
  if [ -n "$LIMIT" ] && [ "$ITERATION" -gt "$LIMIT" ]; then
    echo "Reached limit ($LIMIT)"
    push_to_backup
    exit 0
  fi

  # Run single iteration in Docker
  # Container runs ONE iteration, then exits
  # Backup runs on HOST after container exits
  # Note: MODEL is already validated against whitelist above
  ITERATION_START=$(date +%s)
  if docker run --rm \
    --user "$(id -u):$(id -g)" \
    -v "$PROJECT_DIR:/workspace" \
    -w /workspace \
    -e "CLAUDE_CODE_OAUTH_TOKEN=$CLAUDE_CODE_OAUTH_TOKEN" \
    -e "HOME=/tmp" \
    "$IMAGE_NAME" \
    bash -c "cat '$PROMPT_FILE' | claude --model '$MODEL' -p --dangerously-skip-permissions --output-format text" \
    2>&1 | tee -a "$LOG_FILE"; then

    ITERATION_DURATION=$(( $(date +%s) - ITERATION_START ))
    echo "Iteration $ITERATION complete (${ITERATION_DURATION}s)"

    # A real iteration (study specs, investigate, implement, test, commit) takes
    # minutes. A run that "completes" in under 30s almost always means the underlying
    # `claude` call failed fast (auth/rate-limit/session-limit) but still exited 0 --
    # e.g. a session-limit rejection returns near-instantly. Without this check the
    # loop just keeps retrying every ~5s until it burns through the iteration cap,
    # silently wasting attempts instead of surfacing the real problem. (Hit 2026-07-24:
    # iterations 9-16 all completed within 45s total.)
    if [ "$ITERATION_DURATION" -lt 30 ]; then
      FAST_STREAK=$((FAST_STREAK + 1))
      if [ "$FAST_STREAK" -ge 3 ]; then
        echo ""
        echo "FATAL: $FAST_STREAK consecutive iterations completed in under 30s each."
        echo "This is not real work -- almost certainly a fast-failing claude call"
        echo "(auth, rate limit, or session/usage limit) that still exits 0. Check"
        echo "$LOG_FILE for the actual error, then re-run once resolved."
        push_to_backup
        exit 1
      fi
    else
      FAST_STREAK=0
    fi

    # Real-Office VBA test bridge ON HOST, then push (has gh auth)
    run_vba_test_bridge_if_needed
    push_to_backup
  else
    echo "Claude exited with error"
    push_to_backup
    exit 1
  fi

  sleep 1
done
