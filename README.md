# Codex Native Agent (WPF + VB.NET)

A native Windows desktop client for Codex App Server, built with **VB.NET + WPF**.

The app communicates with `codex app-server` over JSON-RPC (stdio), supports authentication, threads, turns, approvals, reconnect/watchdog behavior, and diagnostics export.

## Features

- Native WPF desktop UI
- App Server process lifecycle and JSON-RPC transport
- Authentication flows:
  - API key login
  - ChatGPT browser login
  - External ChatGPT auth tokens
- Thread and model workflows:
  - model listing
  - thread start/list/read/resume/fork/archive/unarchive
  - thread grouping by project folder (`cwd`) with search/sort
  - last-active tracking
- Turn workflows:
  - start/steer/interrupt
  - streamed transcript/protocol updates
- Approval workflows:
  - command/file approvals
  - `request_user_input` interactive dialog support
- Reliability:
  - watchdog monitoring
  - auto reconnect loop
  - diagnostics export (logs + config snapshot)

## Requirements

- Windows 10/11
- .NET 10 SDK
- `codex` CLI installed and available on `PATH` (or configure full path in app settings)

## Run

From repo root:

```bash
dotnet run --project src/CodexNativeAgent/CodexNativeAgent.vbproj
```

## Project Structure

- `src/CodexNativeAgent/Program.vb`: app entry point
- `src/CodexNativeAgent/Ui/MainWindow.xaml`: WPF UI layout
- `src/CodexNativeAgent/Ui/MainWindow.xaml.vb`: UI logic and workflow handling
- `src/CodexNativeAgent/Ui/QuestionPromptWindow.xaml`: user-input prompt dialog
- `src/CodexNativeAgent/AppServer/`: JSON-RPC client and helpers
- `src/CodexNativeAgent/Services/`: connection/auth/thread/turn/approval services
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
