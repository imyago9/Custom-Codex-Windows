# End-to-End App-Server Flow Stability Report and Rewrite Plan (First Pass)

## Scope Reviewed (Read Fully)

### Required by request
- `app-server.md`
- `src/CodexNativeAgent/AppServer/CodexAppServerClient.vb`
- `src/CodexNativeAgent/Ui/MainWindow.Session.vb`
- `src/CodexNativeAgent/Ui/Coordinators/SessionNotificationCoordinator.vb`
- `src/CodexNativeAgent/Ui/Coordinators/TurnFlowEventModels.vb`
- `src/CodexNativeAgent/Ui/Coordinators/TurnFlowRuntimeStore.vb`
- `src/CodexNativeAgent/Ui/MainWindow.Turns.vb`
- `src/CodexNativeAgent/Ui/ViewModels/TranscriptPanelViewModel.vb`
- `src/CodexNativeAgent/Ui/ViewModels/Transcript/TranscriptEntryModels.vb`
- `src/CodexNativeAgent/Ui/WorkspacePaneControl.xaml`
- `src/CodexNativeAgent/Ui/ViewModels/MainWindowViewModel.vb`
- `src/CodexNativeAgent/Ui/MainWindow.Threads.vb`
- `src/CodexNativeAgent/Ui/Coordinators/ThreadTranscriptSnapshotBuilder.vb`
- `src/CodexNativeAgent/Services/CodexApprovalService.vb`

### Additional supporting files reviewed to validate wire-level behavior
- `src/CodexNativeAgent/Services/CodexTurnService.vb`
- `src/CodexNativeAgent/Services/CodexThreadService.vb`
- `src/CodexNativeAgent/Services/CodexAccountService.vb`
- `src/CodexNativeAgent/Services/CodexConnectionService.vb`
- `src/CodexNativeAgent/Ui/Coordinators/TurnWorkflowCoordinator.vb`
- `src/CodexNativeAgent/Ui/Coordinators/ThreadWorkflowCoordinator.vb`
- `src/CodexNativeAgent/Ui/Coordinators/SessionCoordinator.vb`
- `src/CodexNativeAgent/Ui/StatusBarPaneControl.xaml`
- `src/CodexNativeAgent/Ui/ViewModels/TurnComposerViewModel.vb`
- `src/CodexNativeAgent/AppServer/JsonNodeUtils.vb`
- `src/CodexNativeAgent/Ui/Infrastructure/JsonNodeHelpers.vb`
- `src/CodexNativeAgent/Ui/MainWindow.xaml.vb` (targeted sections for wiring/state gating)

## Executive Summary

The end-to-end instability is primarily caused by **protocol drift** between this client and the current `app-server.md` contract, plus a few **state machine coupling issues** that amplify failures.

The most damaging issues are:
1. **Request enum/value mismatch** for `approvalPolicy` and `sandbox`.
2. **External ChatGPT token flow mismatch** (`chatgptAuthTokens`) between expected fields and sent fields.
3. **Server-request method mismatch** (`tool/requestUserInput` vs `item/tool/requestUserInput`).
4. **Event reducer drops events** when `threadId`/`turnId` are missing in payloads it currently assumes are mandatory.

These create exactly the symptoms you described: turns that appear stuck/inconsistent, approvals/tool flows failing, and behavior that feels unstable versus server expectations.

## Why the current flow is unstable (root-cause analysis)

## 1) Contract drift in request payload values

### Evidence
- UI uses legacy/hyphen values:
  - `TurnComposerViewModel`: `on-request`, `workspace-write` defaults (`src/CodexNativeAgent/Ui/ViewModels/TurnComposerViewModel.vb:13-14`).
  - Status bar options include `untrusted`, `on-failure`, `on-request`, `never`; sandbox options `workspace-write`, `read-only`, `danger-full-access` (`src/CodexNativeAgent/Ui/StatusBarPaneControl.xaml:21-35`).
- These are sent directly as wire values:
  - Thread APIs: `approvalPolicy`, `sandbox` (`src/CodexNativeAgent/Services/CodexThreadService.vb:251-260`).
  - Turn APIs: `approvalPolicy` (`src/CodexNativeAgent/Services/CodexTurnService.vb:38-40`).
- `app-server.md` documents camelCase values such as `onRequest`, `unlessTrusted`, and `workspaceWrite`.

### Impact
- Server-side validation failures or behavior deviations depending on app-server version.
- Non-deterministic UX if some values are tolerated and others are rejected.

## 2) External token auth flow does not match app-server contract

### Evidence
- `account/login/start` for `chatgptAuthTokens` currently sends:
  - `accessToken`, `chatgptAccountId`, optional `chatgptPlanType`
  - no `idToken` (`src/CodexNativeAgent/Services/CodexAccountService.vb:88-103`).
- Refresh response currently sends:
  - `accessToken`, `chatgptAccountId`, optional `chatgptPlanType`
  - no `idToken` (`src/CodexNativeAgent/Ui/MainWindow.Session.vb:965-974`).
- `app-server.md` documents `chatgptAuthTokens` login and refresh using `idToken` + `accessToken`.

### Impact
- External-token auth mode can fail to establish or fail to refresh.
- Repeated 401/refresh loops can manifest as unstable connection/session behavior.

## 3) Server request method mismatch for tool user-input prompts

### Evidence
- Coordinator handles:
  - `item/tool/requestUserInput`
  - `item/tool/call`
  (`src/CodexNativeAgent/Ui/Coordinators/TurnWorkflowCoordinator.vb:198-204`)
- Otherwise returns `-32601 Unsupported server request method` (`src/CodexNativeAgent/Ui/Coordinators/TurnWorkflowCoordinator.vb:208-212`).
- `app-server.md` documents `tool/requestUserInput`.

### Impact
- Tool interaction prompts can fail even though server is behaving correctly.
- User sees broken tool workflows and approval-like prompts not appearing correctly.

## 4) Event processing assumes fields that may not be guaranteed the way current code requires

### Evidence
- Reducer hard-skips turn lifecycle events without both `threadId` and `turnId`:
  - `TurnStarted`: skip (`src/CodexNativeAgent/Ui/Coordinators/TurnFlowRuntimeStore.vb:256-259`)
  - `TurnCompleted`: skip (`src/CodexNativeAgent/Ui/Coordinators/TurnFlowRuntimeStore.vb:278-281`)
- First-seen item is skipped if no `threadId`/`turnId`:
  - (`src/CodexNativeAgent/Ui/Coordinators/TurnFlowRuntimeStore.vb:743-746`)
- Notification pipeline stops on skipped events:
  - (`src/CodexNativeAgent/Ui/Coordinators/SessionNotificationCoordinator.vb:112-118`)
- Parser only looks for `threadId` in a small set of locations:
  - (`src/CodexNativeAgent/Ui/Coordinators/TurnFlowEventModels.vb:441-450`)

### Impact
- Missing item rendering and missing completion transitions.
- `_currentTurnId` can stay set if `turn/completed` gets skipped, leaving UI in false “turn active” state.
- Approval/UI timelines can desync from real server state.

## 5) Single global item keying and single-channel transcript state

### Evidence
- Runtime items keyed by `itemId` only (`src/CodexNativeAgent/Ui/Coordinators/TurnFlowRuntimeStore.vb:99`).
- No composite `(threadId, turnId, itemId)` keying for canonical identity.
- Transcript upserts by `item:{itemId}` (`src/CodexNativeAgent/Ui/ViewModels/TranscriptPanelViewModel.vb:624`).

### Impact
- Potential cross-turn collisions or accidental overwrites if item ids are reused.
- Hidden class of “wrong row updated” bugs that look intermittent.

## 6) Thread lifecycle/view model coupling adds race paths

### Evidence
- “New thread” is local draft only; server thread is created lazily on first prompt (`src/CodexNativeAgent/Ui/MainWindow.Threads.vb:38-49`, `590-603`).
- After first turn, client retries thread refresh to find created thread (`src/CodexNativeAgent/Ui/MainWindow.Threads.vb:1063-1083`).
- Selecting thread uses `thread/resume` first, then `thread/read` fallback for turns (`src/CodexNativeAgent/Ui/MainWindow.Threads.vb:190-197`).

### Impact
- Extra timing windows where UI state and server state diverge.
- More moving parts for list synchronization and active-thread switching.

## 7) Important notifications are currently ignored

### Evidence
- Coordinator only handles a subset of non-turn notifications (`thread/started`, account notifications, `model/rerouted`) and treats many others as unknown (`src/CodexNativeAgent/Ui/Coordinators/SessionNotificationCoordinator.vb:68-107`, `120-129`).
- `app-server.md` includes other lifecycle notifications (`thread/status/changed`, archive/unarchive events, sandbox setup completion, app updates, etc.).

### Impact
- Runtime status does not fully track server lifecycle.
- List/status UI can lag reality and appear unstable.

## 8) Observability gap: system messages suppressed from display transcript

### Evidence
- `ShouldSuppressSystemMessage` always returns `True` (`src/CodexNativeAgent/Ui/ViewModels/TranscriptPanelViewModel.vb:1635`).

### Impact
- User-facing transcript hides key lifecycle hints and failures.
- Debugging flow failures becomes harder, making instability feel worse.

## Contract Alignment Matrix (Current vs Required)

| Area | App-server contract | Current implementation | Severity |
|---|---|---|---|
| `approvalPolicy` values | camelCase policy enums (`onRequest`, `unlessTrusted`, `never`) | legacy/hyphen UI values forwarded directly | Critical |
| sandbox values | camelCase sandbox identifiers (`workspaceWrite`, etc.) | legacy/hyphen values forwarded directly | Critical |
| External token login | `idToken` + `accessToken` | `accessToken` + account/plan fields | Critical |
| External token refresh response | `idToken` + `accessToken` | `accessToken` + account/plan fields | Critical |
| Tool user input server request | `tool/requestUserInput` | expects `item/tool/requestUserInput` | Critical |
| Turn/item event scope handling | robust parsing per schema, no silent drops | hard-skips missing thread/turn scope | High |
| Runtime item identity | scoped item identity | global `itemId` identity | High |
| Notification coverage | full lifecycle coverage | partial coverage | Medium |
| New thread lifecycle | explicit thread lifecycle per protocol | local draft + delayed creation | Medium |
| Diagnostic visibility | meaningful user-visible lifecycle markers | system messages fully suppressed | Medium |

## File-by-file first-pass diagnosis and rewrite intent

### `CodexAppServerClient.vb`
- Keep transport core.
- Refactor for strict request/response error classification and cancellation semantics.
- Add explicit typed method wrappers or shared method constants to remove string drift.

### `MainWindow.Session.vb`
- Replace ad-hoc notification wiring with typed dispatch based on app-server contract map.
- Ensure auth refresh flows and account events are exact-match to current server methods.

### `SessionNotificationCoordinator.vb`
- Expand to full documented notification set for in-scope features.
- Remove silent behavior gaps; convert unknowns to explicit diagnostics with contract version tagging.
- Do not skip valid events due overly strict preconditions.

### `TurnFlowEventModels.vb`
- Replace loose parser with schema-driven event models.
- Promote exact payload shape handling for `turn/*`, `item/*`, and `error`.

### `TurnFlowRuntimeStore.vb`
- Rebuild as strict state machine keyed by `(threadId, turnId, itemId)`.
- Remove fragile assumptions that force skipping first-seen events.
- Ensure deterministic transitions for started/completed/failed/interrupted.

### `MainWindow.Turns.vb`
- Update server request handling for `tool/requestUserInput` and current server methods.
- Keep transcript render path single-source-of-truth from runtime state.

### `TranscriptPanelViewModel.vb`
- Keep rendering pipeline, but stop suppressing all system messages.
- Improve runtime upsert keys to include thread/turn scope.

### `TranscriptEntryModels.vb`
- Add fields needed for richer approval/network prompts and strict lifecycle badges.

### `WorkspacePaneControl.xaml`
- Wire UI affordances to new state model (approval context, turn status, diagnostics visibility).

### `MainWindowViewModel.vb`
- Add explicit active/viewing thread state if needed to avoid cross-thread contamination.

### `MainWindow.Threads.vb`
- Make thread lifecycle explicit and deterministic.
- Reduce race-prone delayed thread creation behavior.
- Avoid unnecessary `thread/resume` on purely historical read paths.

### `ThreadTranscriptSnapshotBuilder.vb`
- Use same runtime/event model as live stream to avoid drift between historical and live rendering.

### `CodexApprovalService.vb`
- Implement full decision matrix including `acceptWithExecpolicyAmendment`.
- Differentiate network approval context from generic command approval.

## Full rewrite plan (1:1 with `app-server.md`, no fallback behavior)

## Non-negotiable rewrite principles
1. **Schema-first**: generated schema/types are the source of truth.
2. **No alias fallback**: no legacy method/value aliases on wire.
3. **No silent skip**: invalid payloads become explicit diagnostics/errors.
4. **Single state machine**: one canonical runtime reducer for live and replay.
5. **Deterministic identity**: scope every runtime entity by thread/turn/item keys.

## Phase 0: Contract lock and prep
1. Pin target app-server contract version from this repository’s `app-server.md`.
2. Create method/value constant sets from spec sections.
3. Define a compatibility policy document (what is intentionally unsupported).

## Phase 1: Wire contract normalization
1. Normalize all method names to spec names.
2. Replace approval/sandbox policy values with contract enums and explicit UI label mapping.
3. Fix `chatgptAuthTokens` login/refresh payloads to exact spec fields.
4. Update tool server-request routing to `tool/requestUserInput`.

## Phase 2: Event model and reducer rewrite
1. Rebuild `TurnFlowEventModels` with strict schema-driven parse rules.
2. Rebuild `TurnFlowRuntimeStore` with scoped keys and deterministic transitions.
3. Handle full in-scope notification set (thread, turn, item, approvals, auth updates).
4. Remove “drop on missing scope” behavior; use schema-required context only.

## Phase 3: MainWindow integration rewrite
1. Refactor `MainWindow.Session` notification path to typed event dispatcher.
2. Refactor `MainWindow.Turns` item/approval handling to consume new reducer output only.
3. Refactor `MainWindow.Threads` lifecycle so thread creation/resume/read flows are explicit and predictable.
4. Ensure UI enable/disable states are derived from canonical runtime state.

## Phase 4: Transcript and snapshot unification
1. Refactor `TranscriptPanelViewModel` runtime upsert keys to scoped identity.
2. Remove unconditional suppression of system entries.
3. Refactor `ThreadTranscriptSnapshotBuilder` to replay through the same reducer/event model used live.

## Phase 5: Approval flow hardening
1. Implement all documented approval decisions including exec policy amendment payload shape.
2. Render network approval context distinctly when provided.
3. Ensure approval lifecycle is tied to `threadId` + `turnId` + `itemId`.

## Phase 6: Test harness and verification
1. Add parser tests for every in-scope method from `app-server.md` examples.
2. Add reducer sequence tests for:
   - normal turn lifecycle
   - interrupt lifecycle
   - failure lifecycle
   - command/file approvals
   - tool user input request flow
3. Add integration tests for:
   - connect/initialize
   - thread start/resume/read/list
   - turn start/steer/interrupt
   - auth modes (api key, chatgpt, external tokens)
4. Add transcript consistency tests (live stream vs snapshot replay parity).

## Phase 7: 

## Acceptance criteria for “100%, 1:1” target
1. Every request/notification method used by this client matches documented method names exactly.
2. Every enum-like wire value emitted by client matches documented value set exactly.
3. No event necessary for turn completion/transcript can be dropped silently.
4. External token auth mode works with documented login + refresh payloads.
5. Tool request-user-input flow works with documented server request method.
6. Thread/turn/item runtime state is deterministic under replay and live stream.

## Refactor execution order across requested files
1. `TurnFlowEventModels.vb`
2. `TurnFlowRuntimeStore.vb`
3. `SessionNotificationCoordinator.vb`
4. `CodexApprovalService.vb`
5. `CodexAppServerClient.vb`
6. `MainWindow.Session.vb`
7. `MainWindow.Turns.vb`
8. `MainWindow.Threads.vb`
9. `ThreadTranscriptSnapshotBuilder.vb`
10. `TranscriptPanelViewModel.vb`
11. `TranscriptEntryModels.vb`
12. `MainWindowViewModel.vb`
13. `WorkspacePaneControl.xaml`

## Risks and mitigations

### Risk
- Behavior changes may expose assumptions currently hidden by silent skips.

### Mitigation
- Ship with verbose diagnostics mode and deterministic replay tests before UI polish.

### Risk
- Existing saved settings contain legacy enum values.

### Mitigation
- One-time settings migration in-memory at startup with explicit user-visible confirmation.

### Risk
- Timeline UI may change noticeably once all system/runtime markers appear.

### Mitigation
- Keep visual design, but classify markers and allow filtering by type.

## Immediate next step

Begin Phase 1 with a mechanical contract-alignment patch set:
1. Normalize method name constants and value enums.
2. Fix auth token payload shapes.
3. Fix tool request method names.
4. Add strict assertion logging for any unexpected wire shape.

## Phase 2 Execution Log (Current Pass)

### Implemented
- Rebuilt `TurnFlowEventModels` to parse and model all currently documented server-initiated communication classes used by this app surface:
  - Thread lifecycle notifications (`thread/started`, `thread/archived`, `thread/unarchived`, `thread/status/changed`)
  - Turn lifecycle + item lifecycle + item deltas
  - Account/auth notifications (`account/login/completed`, `account/updated`, `account/rateLimits/updated`)
  - Model reroute notification
  - App + MCP + sandbox + fuzzy search notifications (`app/list/updated`, `mcpServer/oauthLogin/completed`, `windowsSandbox/setupCompleted`, fuzzy session events)
  - Approval server requests
  - Generic notification/request event wrappers for any non-modeled server methods so communication is still handled and logged.
- Rebuilt `TurnFlowRuntimeStore` reducer dispatch to consume the expanded event model and avoid unknown-event drops.
- Refactored item identity to scoped runtime keys (`threadId:turnId:itemId`) in reducer/runtime item ordering to prevent cross-turn collisions.
- Removed first-seen item hard-fail path on missing scope by introducing scope inference logic (existing item map, thread latest turn, single active turn, and deterministic inferred placeholders).
- Updated `SessionNotificationCoordinator` to route all notifications through parser + reducer first, then apply typed UI side-effects from parsed events (including thread/account/model/system notifications) and generic-method diagnostics.
- Updated transcript runtime upsert keying to use scoped item identity (fallback derives from `threadId:turnId:itemId`).

### Validation
- Build succeeds (`CodexNativeAgent.dll`) with alternate outdir:
  - `/p:OutDir='C:\Users\yagof\pycharmprojects\custom-codex-windows\tmp-build\' /p:UseAppHost=false`

### Remaining for full Phase 2 hardening
- Add sequence/reducer tests for new inference paths and generic-notification coverage.
- Tighten per-method reducer state updates for advanced notifications that are currently diagnostics-only (for example richer thread status materialization).

## Phase 3 Execution Log (Current Pass)

### Implemented
- Refactored `MainWindow.Session` notification handling to consume typed dispatch objects from `SessionNotificationCoordinator` instead of passing large UI callback sets directly.
- Added typed dispatch surfaces in `SessionNotificationCoordinator`:
  - `DispatchNotification(...)`
  - `DispatchServerRequest(...)`
  - `DispatchApprovalResolution(...)`
- Consolidated notification/server-request/approval side-effect application in `MainWindow.Session`:
  - protocol/debug stream updates
  - runtime diagnostics
  - runtime item upserts
  - turn lifecycle markers
  - turn metadata markers
  - token-usage widget updates
  - auth refresh/rate-limit/login-id effects
- Refactored `MainWindow.Turns` server-request path to consume reducer-driven dispatch output (`DispatchServerRequest`) before workflow command routing.
- Refactored thread selection load lifecycle in `MainWindow.Threads` to explicit steps:
  - `ResumeThreadForSelectionAsync`
  - `ReadThreadForSelectionAsync`
  - `LoadThreadObjectForSelectionAsync` (`resume -> read(includeTurns)` fallback)
- Added runtime-store query helpers in `TurnFlowRuntimeStore` (`HasActiveTurn`, `GetActiveTurnId`, `GetLatestTurnId`, `GetTurnState`) and wired `MainWindow` control-state gating to runtime turn activity.
- Added runtime turn synchronization helper in `MainWindow.Session` so `_currentTurnId` is reconciled against canonical runtime state after notifications/server requests/approval resolution and during session-state sync.

### Validation
- Build succeeds (`CodexNativeAgent.dll`) with alternate outdir:
  - `powershell.exe -NoProfile -Command "dotnet build src/CodexNativeAgent/CodexNativeAgent.vbproj /p:OutDir='C:\Users\yagof\pycharmprojects\custom-codex-windows\tmp-build\' /p:UseAppHost=false"`
- Regression tests pass (`CodexNativeAgent.Tests`):
  - `powershell.exe -NoProfile -Command "dotnet test tests/CodexNativeAgent.Tests/CodexNativeAgent.Tests.csproj -c Release"`
  - Added coverage for:
    - runtime active-turn query lifecycle (`HasActiveTurn`, `GetActiveTurnId`, `GetLatestTurnId`)
    - preferred active-turn resolution behavior
    - typed notification dispatch handling and unknown-method dedupe behavior
    - typed server-request + approval-resolution runtime item updates (pending approval counts)

### Remaining for full Phase 3 completion
- Add targeted integration tests for thread selection load race scenarios (rapid selection changes, cancellation during `resume/read` fallback).

## Phase 4 Execution Log (Current Pass)

### Implemented
- Removed unconditional system-message suppression in `TranscriptPanelViewModel`; system entries now render whenever message text is non-empty.
- Refactored `ThreadTranscriptSnapshotBuilder` replay path to drive runtime state through parser notifications (`TurnFlowEventParser.ParseNotification(...)` + `TurnFlowRuntimeStore.Reduce(...)`) instead of constructing event classes directly.
- Refactored snapshot display materialization to use runtime-style upsert semantics for turn lifecycle/metadata/item entries:
  - Turn lifecycle markers now resolve to final turn state per turn key.
  - Metadata entries upsert by `threadId + turnId + kind`.
  - Item entries upsert by scoped runtime identity key.
- Kept snapshot descriptor construction aligned with the same descriptor builders used for live runtime entries (`BuildRuntimeItemDescriptorForSnapshot`, turn marker/metadata builders).

### Validation
- Build succeeds:
  - `powershell.exe -NoProfile -Command "dotnet build src/CodexNativeAgent/CodexNativeAgent.vbproj /p:OutDir='C:\Users\yagof\pycharmprojects\custom-codex-windows\tmp-build\' /p:UseAppHost=false"`
- Test suite passes:
  - `powershell.exe -NoProfile -Command "dotnet test tests/CodexNativeAgent.Tests/CodexNativeAgent.Tests.csproj -c Release"`
  - `Passed: 35, Failed: 0`
- Added Phase 4 regression coverage for:
  - system message visibility in transcript panel
  - snapshot turn lifecycle final-state materialization
  - snapshot scoped identity behavior when item IDs repeat across turns
