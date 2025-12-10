# Exchange ActiveSync Test Suite – Modern Dependency Plan

## Objectives
- Replace `Microsoft.Protocols.TestTools`/related legacy PTF dependencies with drop-in, .NET 8–friendly packages.
- Avoid touching existing test code; only adjust project references/csproj files to point at the new framework.
- Keep the current solution structure; add shim projects into the existing solution so swapping references is a csproj change.
- Start by unblocking the `Common` project so higher-layer adapters/tests can follow.

## Work Plan
1. **Catalog PTF Usage (Now)**
   - Map every PTF type/member referenced in the repo (e.g., `ITestSite`, `ManagedAdapterBase`, `TestClassBase`, `Site.Assert`, `Site.Log`, `CaptureRequirement*`, `LogEntryKind`, `GetAdapter<T>()`, ptfconfig sinks).
   - Produce a minimal API compatibility list needed for `Common` first, then adapters/test suites.
2. **Design Shim Architecture (in existing solution)**
   - Add new shim projects into `ExchangeServerEASProtocolTestSuites.sln` targeting `net8.0`; do not create a new solution file.
   - Projects:
     - `Microsoft.Protocols.TestTools.Abstractions` (interfaces/enums used by tests).
     - `Microsoft.Protocols.TestTools` (implementations of asserts/logging/config/adapters).
     - Optional `Microsoft.Protocols.TestTools.VSTS` placeholder if referenced.
   - Keep namespaces/type names identical so existing code builds unchanged once csproj references are switched.
3. **Config System Replacement**
   - Implement `ITestSite.Properties` with XML loader for `.ptfconfig`/`.deployment.ptfconfig`, supporting property substitution `[OtherProp]` and merge of common/global files (mirroring current `Common.Merge*` expectations).
   - Expose strongly typed helpers where needed (e.g., `DefaultProtocolDocShortName`).
4. **Assertions & Logging**
   - Provide `IAssert` and `ILog` implementations with methods used in repo (`Fail`, `AreEqual`, `IsNotNull`, `IsTrue`, `IsNull`, etc.) plus `CaptureRequirement*` variants.
   - Map `LogEntryKind` to modern logging (e.g., `ILogger`) and support optional structured message formatting.
5. **Test Runtime Abstractions**
   - Reimplement lightweight `TestClassBase` with `Site` property, `TestInitialize`/`TestCleanup` hooks, and `GetAdapter<T>()`.
   - Build `ManagedAdapterBase` that stores `Site`, provides `Initialize(ITestSite)`, and shared `Site` accessor used by adapters.
   - Provide DI/container to satisfy `GetAdapter<T>()` calls (factory registration per protocol adapter).
6. **Common Project Unblock**
- Adjust `Common/Common.csproj` references to the shim projects (no code edits).
- Ensure `ActiveSyncClient`/schema validation code compiles against shim (`ITestSite`, `Assert`, logging).
- Stub any unimplemented pieces (e.g., requirement tracking storage) with no-op or telemetry that won’t break behavior.
7. **Adapter/Test Suite Enablement**
- Incrementally switch adapters/test projects (e.g., `MS-ASAIRS`, `MS-ASCMD`, etc.) by changing csproj references to the shim, filling in missing APIs discovered during builds (especially `CaptureRequirement*` overloads). Do not touch existing code files.
- Provide basic implementations for VSTS-specific logging sinks referenced in `.ptfconfig` (Beacon) or add safe no-op fallbacks.
8. **Validation & Toggle**
   - Add a build configuration (or separate solution) that targets .NET 8 using the shim to ensure clean builds.
   - Document the swap procedure: update csproj references, expected config files, and any runtime environment changes.
9. **Hardening**
   - Backfill unit/smoke tests for shim components (config parsing, assertions/logging) to prevent regressions.
   - Plan future improvements (modern logging sinks, better DI, async support) without affecting compatibility.

## Immediate Next Steps
- Finish the PTF API inventory with focus on `Common` dependencies.
- Scaffold the shim projects/solution targeting `net8.0`.
- Implement minimal `ITestSite`, `Assert`, `LogEntryKind`, and config loader sufficient for `Common` to compile and run basic flows.
