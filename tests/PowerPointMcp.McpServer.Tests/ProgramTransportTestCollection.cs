// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.PowerPointMcp.McpServer.Tests;

/// <summary>
/// Collection definition for tests that use Program.ConfigureTestTransport().
/// These tests MUST run sequentially because the in-memory MCP host uses a shared static
/// transport hook (see Program's TestTransportLock) — two tests configuring transport
/// concurrently would throw "Test transport is already configured".
/// </summary>
/// <remarks>
/// Any test that uses Program.ConfigureTestTransport() or otherwise drives the real MCP host
/// must join this collection so the shared transport state stays serialized. Round-trip tests
/// additionally touch real PowerPoint COM, which is itself unsafe to run concurrently (Ripley's
/// charter: maxParallelThreads: 1).
/// </remarks>
[CollectionDefinition("ProgramTransport")]
#pragma warning disable CA1711 // xUnit collection definition requires class name ending in 'Collection' by convention
public class ProgramTransportTestCollection
#pragma warning restore CA1711
{
    // This class has no code - it's a marker for xUnit collection definition.
}
