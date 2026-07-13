using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Sbroenne.PowerPointMcp.Generators.Common;

namespace Sbroenne.PowerPointMcp.Generators.Mcp;

/// <summary>
/// Generates one action-dispatch MCP tool per [ServiceCategory]/[McpTool]-annotated Core
/// interface (e.g. a single "chart" tool with an <c>action</c> enum parameter, instead of
/// separate "add_chart"/"get_chart_data" tools) — matching mcp-server-excel's and this repo's
/// own CLI generator's action-dispatch shape (see
/// <c>ExcelMcp.Generators.Mcp.McpToolGenerator</c> and <c>PowerPointMcp.Generators.Cli.CliSettingsGenerator</c>).
/// </summary>
/// <remarks>
/// Discovers [ServiceCategory] interfaces from referenced assemblies (Core), same pattern as
/// <c>CliSettingsGenerator</c>, since this generator runs inside the McpServer project which
/// references Core as a compiled assembly, not as source. Categories whose
/// <c>McpToolAttribute.SkipMcpToolGeneration</c> flag is set are skipped — used for Presentation,
/// whose session-lifecycle MCP tool (the single hand-written "presentation" action-dispatch
/// tool, mirroring Excel's ExcelFileTool.cs) needs an OPTIONAL session_id (create/open establish
/// a session rather than requiring one), which this generator's fixed, non-nullable session_id
/// parameter shape does not support.
///
/// PowerPoint Core has no enum/TimeSpan/FileOrValue/IProgress&lt;T&gt; parameters (only plain
/// primitives plus IReadOnlyList&lt;string&gt;/IReadOnlyList&lt;double&gt; for Chart), so this
/// port is a deliberately simplified subset of Excel's generator: it emits every exposed
/// parameter as a plain nullable type passed straight through to the already-generated
/// <c>ServiceRegistry.{Category}.RouteAction</c> method (MCP JSON-RPC natively supports array
/// parameters, so no JSON-string encoding is needed here, unlike the CLI generator's
/// string-flag surface).
/// </remarks>
[Generator]
public sealed class McpToolGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        context.RegisterSourceOutput(context.CompilationProvider,
            static (spc, compilation) =>
            {
                var services = DiscoverServices(compilation);
                foreach (var info in services)
                {
                    var code = GenerateMcpTool(info);
                    spc.AddSource($"McpTool.{info.CategoryPascal}.g.cs", SourceText.From(code, Encoding.UTF8));
                }
            });
    }

    /// <summary>
    /// Discovers [ServiceCategory]+[McpTool] interfaces from referenced assemblies, skipping any
    /// category flagged with <c>SkipMcpToolGeneration = true</c>.
    /// </summary>
    private static List<ServiceInfo> DiscoverServices(Compilation compilation)
    {
        var result = new List<ServiceInfo>();

        foreach (var reference in compilation.References)
        {
            if (compilation.GetAssemblyOrModuleSymbol(reference) is not IAssemblySymbol assembly)
                continue;

            foreach (var type in GetAllTypes(assembly.GlobalNamespace))
            {
                if (type.TypeKind != TypeKind.Interface)
                    continue;

                var hasServiceCategory = type.GetAttributes().Any(a =>
                    a.AttributeClass?.Name == "ServiceCategoryAttribute" &&
                    a.AttributeClass?.ContainingNamespace?.ToDisplayString() == "Sbroenne.PowerPointMcp.Core.Attributes");
                if (!hasServiceCategory)
                    continue;

                var mcpToolAttr = type.GetAttributes().FirstOrDefault(a => a.AttributeClass?.Name == "McpToolAttribute");
                if (mcpToolAttr is null)
                    continue; // No MCP surface intended for this category.

                var skip = mcpToolAttr.NamedArguments.Any(na => na.Key == "SkipMcpToolGeneration" && na.Value.Value is true);
                if (skip)
                    continue;

                var info = ServiceInfoExtractor.ExtractServiceInfo(type);
                if (info is null || info.Methods.Count == 0)
                    continue;

                result.Add(info);
            }
        }

        return result.OrderBy(i => i.CategoryPascal, StringComparer.Ordinal).ToList();
    }

    private static IEnumerable<INamedTypeSymbol> GetAllTypes(INamespaceSymbol ns)
    {
        foreach (var type in ns.GetTypeMembers())
            yield return type;
        foreach (var child in ns.GetNamespaceMembers())
            foreach (var type in GetAllTypes(child))
                yield return type;
    }

    /// <summary>
    /// All unique exposed parameters across every action in the category, each forced nullable
    /// (every action uses a different subset, so the MCP surface must accept "not supplied" for
    /// any of them) — mirrors <c>ServiceRegistryGenerator</c>'s private helper of the same shape.
    /// </summary>
    private static List<ExposedParameter> GetNullableExposedParameters(ServiceInfo info)
    {
        var parameters = ServiceInfoExtractor.GetAllExposedParameters(info);
        foreach (var p in parameters)
        {
            if (!p.TypeName.EndsWith("?", StringComparison.Ordinal))
                p.TypeName += "?";
        }
        return parameters;
    }

    private static string EscapeDescription(string? text)
    {
        var value = (text ?? string.Empty)
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\r", "")
            .Replace("\n", " ")
            .Trim();
        return value;
    }

    private static string GenerateMcpTool(ServiceInfo info)
    {
        var exposedParams = GetNullableExposedParameters(info);
        var toolDescription = EscapeDescription(
            info.McpToolDescription ?? info.XmlDocSummary ?? $"{info.CategoryPascal} operations.");
        var actionList = string.Join(", ", info.Methods.Select(m => m.ActionName));

        var sb = new StringBuilder();
        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member");
        sb.AppendLine();
        sb.AppendLine("using System.ComponentModel;");
        sb.AppendLine("using ModelContextProtocol.Server;");
        sb.AppendLine("using Sbroenne.PowerPointMcp.Generated;");
        sb.AppendLine("using Sbroenne.PowerPointMcp.McpServer.Infrastructure;");
        sb.AppendLine("using Sbroenne.PowerPointMcp.Service;");
        sb.AppendLine();
        sb.AppendLine("namespace Sbroenne.PowerPointMcp.McpServer.Tools;");
        sb.AppendLine();
        sb.AppendLine("/// <summary>");
        sb.AppendLine($"/// Generated action-dispatch MCP tool for {info.Category} operations.");
        sb.AppendLine("/// </summary>");
        sb.AppendLine("[McpServerToolType]");
        sb.AppendLine($"public static class PowerPoint{info.CategoryPascal}Tool");
        sb.AppendLine("{");
        sb.AppendLine($"    [McpServerTool(Name = \"{info.McpToolName}\")]");
        sb.AppendLine($"    [Description(\"{toolDescription} Actions: {actionList}.\")]");
        sb.AppendLine($"    public static string PowerPoint{info.CategoryPascal}(");
        sb.AppendLine($"        [Description(\"The action to perform. One of: {actionList}.\")] {info.CategoryPascal}Action action,");
        sb.AppendLine("        [Description(\"The session id returned by the presentation tool's action=open or action=create.\")] string session_id,");

        if (exposedParams.Count == 0)
        {
            sb.AppendLine("        PowerPointMcpService service)");
        }
        else
        {
            sb.AppendLine("        PowerPointMcpService service,");
            for (int i = 0; i < exposedParams.Count; i++)
            {
                var p = exposedParams[i];
                var snakeName = StringHelper.ToSnakeCase(p.Name);
                var description = EscapeDescription(p.DescriptionWithRequired ?? StringHelper.GetParameterDescription(p.Name));
                var suffix = i < exposedParams.Count - 1 ? "," : ")";
                sb.AppendLine($"        [Description(\"{description}\")] {p.TypeName} {snakeName} = null{suffix}");
            }
        }

        sb.AppendLine("    {");
        sb.AppendLine("        return PowerPointToolsBase.ExecuteToolAction(");
        sb.AppendLine($"            \"{info.McpToolName}\",");
        sb.AppendLine($"            ServiceRegistry.{info.CategoryPascal}.ToActionString(action),");
        sb.AppendLine($"            () => ServiceRegistry.{info.CategoryPascal}.RouteAction(");
        sb.AppendLine("                action,");
        sb.AppendLine("                session_id,");

        var forwardLine = "                (command, sid, args) => ServiceBridge.ForwardToService(service, command, sid, args)";
        sb.AppendLine(exposedParams.Count > 0 ? forwardLine + "," : forwardLine + "));");

        for (int i = 0; i < exposedParams.Count; i++)
        {
            var p = exposedParams[i];
            var snakeName = StringHelper.ToSnakeCase(p.Name);
            var suffix = i < exposedParams.Count - 1 ? "," : "));";
            sb.AppendLine($"                {p.Name}: {snakeName}{suffix}");
        }

        sb.AppendLine("    }");
        sb.AppendLine("}");

        return sb.ToString();
    }
}
