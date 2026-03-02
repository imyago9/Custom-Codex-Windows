# Codex Native Agent (WPF + VB.NET)

A native Windows desktop client for `codex app-server`, built with **VB.NET + WPF**.

This project is a full-featured Codex UI for local development workflows: authentication, thread/turn orchestration, approvals, transcript rendering, skills/apps mentions, Git operations, and local usage analytics.

## Screenshots

| New Thread (Light) | New Thread (Dark) |
| --- | --- |
| ![New thread (light mode)](_app_screenshots/new-thread%28light-mode%29.png) | ![New thread (dark mode)](_app_screenshots/new-thread%28dark-mode%29.png) |

| Chat In Progress (Light) | Chat In Progress (Dark) |
| --- | --- |
| ![Chat in progress (light mode)](_app_screenshots/chat-in-progress%28light-mode%29.png) | ![Chat in progress (dark mode)](_app_screenshots/chat-in-progress%28dark-mode%29.png) |

## Current Capabilities

### App Server Integration

- JSON-RPC transport over `codex app-server` (stdio)
- Full thread + turn lifecycle integration (`thread/*`, `turn/*`, `item/*`)
- Streaming transcript updates from server notifications
- Approval request handling and `request_user_input` dialog support
- Native support for skills (`skills/list`) and apps/connectors (`app/list`)

### Authentication and Session

- API key login
- ChatGPT browser login
- External token login
- Account/rate-limit reads
- Session watchdog, reconnect behavior, and diagnostics export

### Threads and Tabs

- Start, resume, read, fork, archive, and unarchive threads
- Thread grouping by project (`cwd`), with search/sort/filter controls
- Reopen and reconstruct thread runtime state reliably from snapshots + overlays
- Transcript tab strip with drag-and-drop reordering and drop indicator
- Runtime status dots in both thread list and tab chips (active, approval pending, pending updates)

### Turn Composer and Interaction

- `Enter` to send prompt (no `Ctrl+Enter` required)
- If a turn is active, `Enter` sends a steer input
- Model, effort, and approval controls in composer
- `$` token suggestions for skills/apps with keyboard and mouse selection
- Suggestion popup styled for light/dark themes with truncated text and descriptions

### Transcript UX

- Structured bubbles/rows for user, assistant, commands, file changes, plans, reasoning, and lifecycle markers
- Active turn status row with:
  - animated progress circle
  - elapsed timer (`mm:ss`)
  - live reasoning/working status text with animated ellipsis
- Persistent failed/interrupted/canceled prompt highlighting
- Improved lifecycle marker styling (`Turn started`, completion/failure states)
- Settings toggles for transcript lifecycle dots and reasoning bubble visibility
- Diff-related transcript entries hide noisy unified-diff header lines for readability

### Approval UX

- Queue-aware approval panel with active option selection
- Dot-based option highlighting (no chip chrome)
- Approval-needed indicators across thread list and tabs

### Built-in Git Panel

- Changes / History / Branches tabs
- Stage/unstage per change and stage/unstage all actions
- Commit composer under tab row with:
  - commit message box
  - amend checkbox
  - commit action
- Push action integrated with commit flow
- Primary action behavior similar to VS Code:
  - show `Push` when push is available and working tree is clean
  - otherwise show `Commit`
- Rich diff preview with:
  - resizable split layout
  - cleaner diff text (header lines hidden)
  - inline line editing workflow with save/cancel
- Diff/history/branch preview panes resize with panel width

### Metrics Dashboard

- Left-nav `Metrics` button opens a right-side metrics panel
- Metrics ingestion from local Codex data:
  - `.codex/history.jsonl`
  - `.codex/sessions`
  - `.codex/archived_sessions`
- Summary cards for sessions, prompts, projects, tokens, and tool calls
- Daily calendar heatmap with styled hover tooltips
- Month navigation (`prev`, `next`, `today`)
- Filter modes:
  - all activity
  - single day
  - last N days
- Project breakdown table
- Expanded model usage table (turn contexts, sessions, input/output tokens, tools, share)

### Sound and Feedback

- UI sound support with in-app toggle and volume setting
- Event sounds include:
  - `load-thread.mp3`
  - `turn-done.mp3`
  - `turn-failed.mp3`
  - `general-error.mp3`
  - `approval-needed.mp3`
  - `git-commit.mp3`
  - `git-push.mp3`
  - `hide-show-side-panels.mp3`
- Duplicate-suppression windows prevent repeated sound spam

### Appearance and Settings

- Light and dark themes
- Compact and comfortable density modes
- Transcript scale presets
- Thread filtering by working directory
- Persistent settings store for UI and behavior preferences

## Architecture

The app follows a pragmatic **MVVM + coordinator** architecture:

- ViewModels manage UI state (`ThreadsPanel`, `TurnComposer`, `TranscriptPanel`, `ApprovalPanel`, `SettingsPanel`, `SessionState`)
- Coordinators manage workflow orchestration (session, notifications, threads, turns, shell/approval flow)
- Services isolate RPC concerns (connection/auth/thread/turn/approval/skills-apps)
- `MainWindow` is the shell host and UI integration boundary
- UI is composed from focused user controls (sidebar/workspace/status/git/metrics panes)

See [`app-server.md`](app-server.md) for the protocol surface and method mapping used by this client.

## Requirements

- Windows 10/11
- .NET 10 SDK
- `codex` CLI installed and available on `PATH` (or configured via app settings)

## Run

From repo root:

```bash
dotnet run --project src/CodexNativeAgent/CodexNativeAgent.vbproj
```

## Test

```bash
dotnet test tests/CodexNativeAgent.Tests/CodexNativeAgent.Tests.csproj
```

Current test suite includes coordinator and view-model coverage for session, thread, transcript, and runtime-store behavior.

## Project Structure

- `src/CodexNativeAgent/Program.vb`: application entry point
- `src/CodexNativeAgent/Ui/MainWindow.xaml`: shell layout
- `src/CodexNativeAgent/Ui/MainWindow.xaml.vb`: shell host + shared UI integration
- `src/CodexNativeAgent/Ui/MainWindow.Session.vb`: session/auth lifecycle
- `src/CodexNativeAgent/Ui/MainWindow.Threads.vb`: thread workflows and list behaviors
- `src/CodexNativeAgent/Ui/MainWindow.Turns.vb`: turn lifecycle, streaming, approvals
- `src/CodexNativeAgent/Ui/MainWindow.TranscriptTabs.vb`: tab strip surfaces and reordering
- `src/CodexNativeAgent/Ui/MainWindow.SkillsApps.vb`: skills/apps catalog + mention integration
- `src/CodexNativeAgent/Ui/MainWindow.Metrics.vb`: metrics parsing, filtering, rendering
- `src/CodexNativeAgent/Ui/MainWindow.Audio.vb`: UI sound routing and suppression
- `src/CodexNativeAgent/Ui/SidebarPaneControl.xaml`: left navigation, threads, settings
- `src/CodexNativeAgent/Ui/WorkspacePaneControl.xaml`: transcript + composer workspace
- `src/CodexNativeAgent/Ui/GitPaneControl.xaml`: Git inspector panel
- `src/CodexNativeAgent/Ui/MetricsPaneControl.xaml`: metrics inspector panel
- `src/CodexNativeAgent/Ui/Coordinators/`: orchestration layer
- `src/CodexNativeAgent/Services/`: RPC service layer
- `src/CodexNativeAgent/AppServer/`: JSON-RPC client/contracts
- `tests/CodexNativeAgent.Tests/`: unit tests

## Public Repo Hygiene

This repo is prepared for public hosting:

- Build outputs and local IDE state are ignored via `.gitignore`
- Local runtime data that may contain private history/auth is ignored (`.codex`, `sessions`, `auth.json`, `history.jsonl`)
- Temporary local folders are ignored (`tmpbuild`, `.dotnet-cli`, etc.)

Before publishing:

1. Verify no credentials/tokens are in committed files.
2. Verify no local runtime/session exports are committed.
3. Rotate any key you suspect was previously exposed.
