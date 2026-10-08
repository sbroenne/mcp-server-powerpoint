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
/// tool, mirroring Excel's ExcelFileTool.cs) needs an OPTIONAL presentation_session_id (create/open establish
/// a session rather than requiring one), which this generator's fixed, non-nullable presentation_session_id
/// parameter shape does not support.
///
/// The generator emits the union of action parameters and passes them through to the generated
/// <c>ServiceRegistry.{Category}.RouteAction</c> method (MCP JSON-RPC natively supports array
/// parameters, so no JSON-string encoding is needed here, unlike the CLI generator's
/// string-flag surface).
/// </remarks>
[Generator]
public sealed class McpToolGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var documentation = context.AdditionalTextsProvider
            .Where(file => file.Path.EndsWith("Sbroenne.PowerPointMcp.Core.xml", StringComparison.OrdinalIgnoreCase))
            .Select((file, cancellationToken) => file.GetText(cancellationToken)?.ToString())
            .Collect();

        context.RegisterSourceOutput(context.CompilationProvider.Combine(documentation),
            static (spc, input) =>
            {
                var compilation = input.Left;
                var coreReference = compilation.References.OfType<PortableExecutableReference>()
                    .FirstOrDefault(reference =>
                        compilation.GetAssemblyOrModuleSymbol(reference) is IAssemblySymbol assembly &&
                        assembly.Name == "Sbroenne.PowerPointMcp.Core");
                if (coreReference is not null)
                {
                    if (input.Right.Length != 1 || string.IsNullOrWhiteSpace(input.Right[0]))
                    {
                        spc.ReportDiagnostic(Diagnostic.Create(
                            new DiagnosticDescriptor("PPTMCP001", "Missing Core documentation",
                                "MCP generation requires the Core XML documentation as an AdditionalFile",
                                "PowerPointMcp.Generation", DiagnosticSeverity.Error, isEnabledByDefault: true),
                            Location.None));
                        return;
                    }

                    compilation = compilation.ReplaceReference(coreReference,
                        ((AssemblyMetadata)coreReference.GetMetadata()).GetReference(
                            documentation: new CoreDocumentationProvider(input.Right[0]!),
                            aliases: coreReference.Properties.Aliases,
                            embedInteropTypes: coreReference.Properties.EmbedInteropTypes,
                            filePath: coreReference.FilePath));
                }
                var services = DiscoverServices(compilation);
                foreach (var info in services)
                {
                    var code = GenerateMcpTool(info);
                    spc.AddSource($"McpTool.{info.CategoryPascal}.g.cs", SourceText.From(code, Encoding.UTF8));

                    if (info.McpReadOnlyActions.Count > 0)
                    {
                        var readOnlyCode = GenerateReadOnlyMcpTool(info);
                        spc.AddSource($"McpTool.{info.CategoryPascal}.Read.g.cs", SourceText.From(readOnlyCode, Encoding.UTF8));
                    }
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
    private static List<ExposedParameter> GetNullableExposedParameters(
        ServiceInfo info,
        IReadOnlyList<MethodInfo> methods)
    {
        var parameters = ServiceInfoExtractor.GetAllExposedParameters(info, methods);
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
        return GenerateMcpTool(
            info,
            info.Methods,
            info.McpToolName,
            $"{info.CategoryPascal}ToolOutputSchema",
            readOnly: false);
    }

    private static string GenerateReadOnlyMcpTool(ServiceInfo info)
    {
        var methods = info.Methods
            .Where(method => info.McpReadOnlyActions.Contains(method.ActionName, StringComparer.OrdinalIgnoreCase))
            .ToArray();

        return GenerateMcpTool(
            info,
            methods,
            $"{info.McpToolName}_read",
            $"{info.CategoryPascal}ReadToolOutputSchema",
            readOnly: true);
    }

    private static string GenerateMcpTool(
        ServiceInfo info,
        IReadOnlyList<MethodInfo> methods,
        string toolName,
        string outputSchemaName,
        bool readOnly)
    {
        var exposedParams = GetNullableExposedParameters(info, methods);
        var baseDescription = info.McpToolDescription ?? info.XmlDocSummary ?? $"{info.CategoryPascal} operations.";
        var toolDescription = EscapeDescription(readOnly
            ? $"Read-only {info.CategoryPascal} inspection actions that do not change the presentation."
            : baseDescription);
        var actionList = string.Join(", ", methods.Select(m => m.ActionName));
        var toolSuffix = readOnly ? "Read" : string.Empty;
        var actionTypeName = $"{info.CategoryPascal}{toolSuffix}Action";
        var toolClassName = $"PowerPoint{info.CategoryPascal}{toolSuffix}Tool";
        var methodName = $"PowerPoint{info.CategoryPascal}{toolSuffix}";
        var title = readOnly
            ? $"{info.McpToolTitle ?? $"PowerPoint {info.CategoryPascal} Operations"} (Read Only)"
            : info.McpToolTitle ?? $"PowerPoint {info.CategoryPascal} Operations";
        var destructive = readOnly ? "false" : info.McpToolDestructive ? "true" : "false";
        var readOnlyAnnotation = readOnly ? ", ReadOnly = true" : string.Empty;
        var category = info.McpToolCategory ?? "content";
        var actionExpression = readOnly
            ? $"ServiceRegistry.{info.CategoryPascal}.ToActionString(System.Enum.Parse<{info.CategoryPascal}Action>(action.ToString()))"
            : $"ServiceRegistry.{info.CategoryPascal}.ToActionString(action)";

        var sb = new StringBuilder();
        sb.AppendLine("// <auto-generated />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member");
        sb.AppendLine();
        sb.AppendLine("using System.ComponentModel;");
        sb.AppendLine("using System.Text.Json.Serialization;");
        sb.AppendLine("using System.Threading;");
        sb.AppendLine("using ModelContextProtocol.Protocol;");
        sb.AppendLine("using ModelContextProtocol.Server;");
        sb.AppendLine("using Sbroenne.PowerPointMcp.Generated;");
        sb.AppendLine("using Sbroenne.PowerPointMcp.McpServer.Infrastructure;");
        sb.AppendLine("using Sbroenne.PowerPointMcp.Service;");
        sb.AppendLine();
        sb.AppendLine("namespace Sbroenne.PowerPointMcp.McpServer.Tools;");
        sb.AppendLine();
        if (readOnly)
        {
            GenerateReadOnlyActionEnum(sb, actionTypeName, methods);
        }
        sb.AppendLine("/// <summary>");
        sb.AppendLine($"/// Generated action-dispatch MCP tool for {info.Category} {(readOnly ? "read-only " : string.Empty)}operations.");
        sb.AppendLine("/// </summary>");
        sb.AppendLine("[McpServerToolType]");
        sb.AppendLine($"public static class {toolClassName}");
        sb.AppendLine("{");
        sb.AppendLine($"    [McpServerTool(Name = \"{toolName}\", Title = \"{title}\", Destructive = {destructive}{readOnlyAnnotation}, UseStructuredContent = true, OutputSchemaType = typeof({outputSchemaName}))]");
        sb.AppendLine($"    [McpMeta(\"category\", \"{category}\")]");
        sb.AppendLine($"    [McpMeta(\"requiresSession\", {(!info.NoSession).ToString().ToLowerInvariant()})]");
        sb.AppendLine($"    [Description(\"{toolDescription} Actions: {actionList}.\")]");
        sb.AppendLine($"    public static Task<CallToolResult> {methodName}(");
        sb.AppendLine($"        [Description(\"The action to perform. One of: {actionList}.\")] {actionTypeName} action,");
        sb.AppendLine("        [Description(\"The session id returned by the presentation tool's action=open or action=create.\")] string presentation_session_id,");

        if (exposedParams.Count == 0)
        {
            sb.AppendLine("        PowerPointMcpService service,");
            sb.AppendLine("        CancellationToken cancellationToken = default)");
        }
        else
        {
            sb.AppendLine("        PowerPointMcpService service,");
            for (int i = 0; i < exposedParams.Count; i++)
            {
                var p = exposedParams[i];
                var snakeName = StringHelper.ToSnakeCase(p.Name);
                var actionDescriptions = methods
                    .SelectMany(method => method.Parameters
                        .Where(parameter => string.Equals(
                            parameter.ExposedName ?? parameter.Name, p.Name, StringComparison.OrdinalIgnoreCase))
                        .Select(parameter => new { method.ActionName, Description = parameter.XmlDocDescription }))
                    .Where(parameter => !string.IsNullOrWhiteSpace(parameter.Description))
                    .GroupBy(parameter => parameter.Description!, StringComparer.Ordinal)
                    .ToArray();
                string explanation;
                if (actionDescriptions.Length == 1)
                    explanation = actionDescriptions[0].Key;
                else if (actionDescriptions.Length > 1)
                    explanation = string.Join(" ", actionDescriptions.Select(group =>
                        $"{string.Join(", ", group.Select(parameter => parameter.ActionName))}: {group.Key}"));
                else
                {
                    explanation = string.Join(" ", methods
                        .Where(method => method.Parameters.Any(parameter =>
                            string.Equals(parameter.ExposedName ?? parameter.Name, p.Name, StringComparison.OrdinalIgnoreCase)))
                        .Select(method => string.IsNullOrWhiteSpace(method.XmlDocSummary)
                            ? method.ActionName
                            : $"{method.ActionName}: {method.XmlDocSummary}"));
                }
                var applicability = p.DescriptionWithRequired?.Substring(p.Description?.Length ?? 0);
                var description = EscapeDescription($"{explanation} {applicability}".Trim());
                sb.AppendLine($"        [Description(\"{description}\")] {p.TypeName} {snakeName} = null,");
            }
            sb.AppendLine("        CancellationToken cancellationToken = default)");
        }

        sb.AppendLine("    {");
        sb.AppendLine("        return PowerPointToolsBase.ExecuteToolActionAsync(");
        sb.AppendLine($"            \"{toolName}\",");
        sb.AppendLine($"            {actionExpression},");
        sb.AppendLine($"            () => ServiceRegistry.{info.CategoryPascal}.RouteAction(");
        sb.AppendLine($"                {(readOnly ? $"System.Enum.Parse<{info.CategoryPascal}Action>(action.ToString())" : "action")},");
        sb.AppendLine("                presentation_session_id,");

        var forwardLine = "                (command, sid, args) => ServiceBridge.ForwardToServiceAsync(service, command, sid, args, cancellationToken)";
        sb.AppendLine(exposedParams.Count > 0
            ? forwardLine + ","
            : forwardLine + "), cancellationToken, service.Sessions, presentation_session_id);");

        for (int i = 0; i < exposedParams.Count; i++)
        {
            var p = exposedParams[i];
            var snakeName = StringHelper.ToSnakeCase(p.Name);
            var suffix = i < exposedParams.Count - 1 ? "," : "), cancellationToken, service.Sessions, presentation_session_id);";
            sb.AppendLine($"                {p.Name}: {snakeName}{suffix}");
        }
        sb.AppendLine("    }");
        sb.AppendLine("}");
        sb.AppendLine();
        GenerateOutputSchemaClass(sb, outputSchemaName, info, methods);

        return sb.ToString();
    }

    private static void GenerateReadOnlyActionEnum(
        StringBuilder sb,
        string actionTypeName,
        IReadOnlyList<MethodInfo> methods)
    {
        sb.AppendLine($"[JsonConverter(typeof(JsonStringEnumConverter<{actionTypeName}>))]");
        sb.AppendLine($"public enum {actionTypeName}");
        sb.AppendLine("{");
        for (var i = 0; i < methods.Count; i++)
        {
            var comma = i < methods.Count - 1 ? "," : string.Empty;
            sb.AppendLine($"    [JsonStringEnumMemberName(\"{methods[i].ActionName}\")]");
            sb.AppendLine($"    {methods[i].MethodName}{comma}");
        }
        sb.AppendLine("}");
        sb.AppendLine();
    }

    private static void GenerateOutputSchemaClass(
        StringBuilder sb,
        string schemaName,
        ServiceInfo info,
        IReadOnlyList<MethodInfo> methods)
    {
        sb.AppendLine($"internal sealed class {schemaName}");
        sb.AppendLine("{");
        foreach (var property in GetOutputSchemaProperties(info, methods))
        {
            sb.AppendLine("    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]");
            sb.AppendLine($"    public {property.TypeName} {property.Name} {{ get; set; }}");
        }
        sb.AppendLine("}");
    }

    private static OutputSchemaProperty[] GetOutputSchemaProperties(
        ServiceInfo info,
        IReadOnlyList<MethodInfo> methods)
    {
        var properties = new Dictionary<string, OutputSchemaProperty>(StringComparer.Ordinal);
        foreach (var method in methods)
        {
            if (method.ReturnTypeSymbol is not INamedTypeSymbol returnType)
                continue;

            for (var type = returnType; type is not null; type = type.BaseType)
            {
                foreach (var property in type.GetMembers().OfType<IPropertySymbol>())
                {
                    if (property.IsStatic || property.IsIndexer ||
                        property.DeclaredAccessibility != Accessibility.Public || property.GetMethod is null)
                        continue;

                    var typeName = GetOptionalSchemaTypeName(property.Type);
                    if (properties.TryGetValue(property.Name, out var existing) && existing.TypeName != typeName)
                        properties[property.Name] = new(property.Name, "System.Text.Json.JsonElement?");
                    else
                        properties[property.Name] = new(property.Name, typeName);
                }
            }
        }

        return properties.Values
            .OrderBy(property => property.Name == "Success" ? 0 : 1)
            .ThenBy(property => property.Name, StringComparer.Ordinal)
            .ToArray();
    }

    private static string GetOptionalSchemaTypeName(ITypeSymbol type)
    {
        var typeName = TypeNameHelper.GetTypeName(type, type.NullableAnnotation);
        return typeName.EndsWith("?", StringComparison.Ordinal) ? typeName : typeName + "?";
    }

    private sealed class OutputSchemaProperty(string name, string typeName)
    {
        public string Name { get; } = name;
        public string TypeName { get; } = typeName;
    }
}
