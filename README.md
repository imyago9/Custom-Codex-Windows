# Codex Native Agent (WPF + VB.NET)

A native Windows desktop client for Codex App Server, built with **VB.NET + WPF**.

The app communicates with `codex app-server` over JSON-RPC (stdio) and includes native workflows for authentication, threads, turns, approvals, reconnect/watchdog behavior, diagnostics export, and a built-in Git inspector.

## Features

- Native WPF desktop UI with light/dark themes and density modes
- App Server process lifecycle + JSON-RPC transport over stdio
- Authentication flows
  - API key login
  - ChatGPT browser login
  - external token login
  - account/rate-limit reads
- Thread workflows
  - start/list/read/resume/fork/archive/unarchive
  - thread grouping by project folder (`cwd`)
  - folder collapse/expand, search, sort, filters
  - auto-load on selection with async loading + transcript loading indicator
- Turn workflows
  - start / steer / interrupt
  - streamed transcript + protocol handling
  - command output rows collapse by default while a turn is in progress
- Transcript UX
  - structured timeline/chat rendering (not raw text log only)
  - styled rows for assistant/user/command/file changes/plan/reasoning/errors
  - protocol log available in a dialog from Settings
- Approval workflows
  - command/file approvals
  - `request_user_input` dialog support
- Built-in Git inspector sidebar (resizable)
  - Changes / History / Branches tabs
  - diff preview with highlighted added/removed lines
  - branch/commit/file previews
  - open changed files in VS Code
- Reliability + diagnostics
  - watchdog monitoring
  - auto reconnect loop
  - diagnostics export (logs + config snapshot)

## Architecture (Current)

The app now uses a **pragmatic MVVM + coordinator** architecture (WPF-friendly hybrid):

- ViewModels own most UI state (`ThreadsPanel`, `TurnComposer`, `TranscriptPanel`, `ApprovalPanel`, `SettingsPanel`, `SessionState`)
- Coordinators own workflow orchestration (session/auth, notifications, threads, turns, shell commands)
- `MainWindow` acts primarily as the shell/view host and UI integration layer
- UI is split into user controls (sidebar, workspace, status bar, git pane)

This keeps behavior modular/testable while still allowing WPF code-behind for view-only concerns (focus, scrolling, popup placement, etc.).

## Requirements

- Windows 10/11
- .NET 10 SDK
- `codex` CLI installed and available on `PATH` (or configure full path in app settings)

## Run

From repo root:

```bash
dotnet run --project src/CodexNativeAgent/CodexNativeAgent.vbproj
```

## Test

```bash
dotnet test tests/CodexNativeAgent.Tests/CodexNativeAgent.Tests.csproj
```

## Project Structure

- `src/CodexNativeAgent/Program.vb`: app entry point
- `src/CodexNativeAgent/Ui/MainWindow.xaml`: WPF UI layout
- `src/CodexNativeAgent/Ui/MainWindow.xaml.vb`: shell host + shared UI integration
- `src/CodexNativeAgent/Ui/MainWindow.Threads.vb`: thread list/workflow UI integration
- `src/CodexNativeAgent/Ui/MainWindow.Session.vb`: session/auth/reconnect UI integration
- `src/CodexNativeAgent/Ui/MainWindow.Turns.vb`: turn/transcript/approval UI integration
- `src/CodexNativeAgent/Ui/MainWindow.Settings.vb`: settings persistence/appearance integration
- `src/CodexNativeAgent/Ui/SidebarPaneControl.xaml`: left sidebar (threads + settings)
- `src/CodexNativeAgent/Ui/WorkspacePaneControl.xaml`: transcript + composer workspace
- `src/CodexNativeAgent/Ui/StatusBarPaneControl.xaml`: bottom status bar
- `src/CodexNativeAgent/Ui/GitPaneControl.xaml`: right Git inspector sidebar
- `src/CodexNativeAgent/Ui/QuestionPromptWindow.xaml`: user-input prompt dialog
- `src/CodexNativeAgent/Ui/ViewModels/`: MVVM view models
- `src/CodexNativeAgent/Ui/Coordinators/`: workflow coordinators
- `src/CodexNativeAgent/AppServer/`: JSON-RPC client and helpers
- `src/CodexNativeAgent/Services/`: connection/auth/thread/turn/approval services
- `tests/CodexNativeAgent.Tests/`: unit tests (view models + coordinator slices)
- `scripts/migrate_wsl_codex_to_windows.sh`: optional migration helper for local Codex data

## Public Repo Hygiene

This repo is prepared for public hosting:

- Build outputs and local IDE state are ignored via `.gitignore`
- Local runtime data that may contain private history/auth is ignored (`.codex`, `sessions`, `auth.json`, `history.jsonl`)
- Temporary local folders are ignored (`tmpbuild`, `.dotnet-cli`, etc.)

Before publishing:

1. Verify no credentials/tokens are in committed files.
2. Verify no local runtime/session exports are committed.
3. Rotate any key you suspect was previously exposed.
