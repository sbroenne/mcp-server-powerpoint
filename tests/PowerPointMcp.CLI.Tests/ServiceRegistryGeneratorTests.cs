using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Sbroenne.PowerPointMcp.Core.Attributes;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Generators;
using Spectre.Console.Cli;

namespace Sbroenne.PowerPointMcp.CLI.Tests;

public sealed class ServiceRegistryGeneratorTests
{
    [Theory]
    [InlineData("int GetCount();", 0)]
    [InlineData("int GetCount(IPresentationBatch batch);", 0)]
    [InlineData("int GetCount(IPresentationBatch batch); int Add(IPresentationBatch batch, int value);", 1)]
    public void GeneratedRoutes_CompileWithAndWithoutExposedParameters(string methods, int exposedParameterCount)
    {
        var source = $$"""
            global using System;
            global using System.Linq;
            using Sbroenne.PowerPointMcp.Core.Attributes;
            using Sbroenne.PowerPointMcp.ComInterop.Session;

            namespace GeneratorTests;

            [ServiceCategory("sample", "Sample")]
            public interface ISampleCommands
            {
                {{methods}}
            }
            """;
        var platformAssemblies = Assert.IsType<string>(AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES"));
        var references = platformAssemblies.Split(Path.PathSeparator)
            .Concat([
                typeof(ServiceCategoryAttribute).Assembly.Location,
                typeof(IPresentationBatch).Assembly.Location,
                typeof(CommandSettings).Assembly.Location
            ])
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(path => MetadataReference.CreateFromFile(path));
        var compilation = CSharpCompilation.Create(
            "GeneratedRoutes",
            [CSharpSyntaxTree.ParseText(source)],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        GeneratorDriver driver = CSharpGeneratorDriver.Create(new ServiceRegistryGenerator());
        driver = driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out var diagnostics);

        Assert.Empty(diagnostics);
        Assert.Contains(
            driver.GetRunResult().GeneratedTrees,
            tree => tree.FilePath.EndsWith("ServiceRegistry.Sample.g.cs", StringComparison.Ordinal));
        using var assembly = new MemoryStream();
        var result = output.Emit(assembly);
        Assert.True(
            result.Success,
            string.Join(Environment.NewLine, result.Diagnostics.Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error)));

        var category = output.GetTypeByMetadataName("Sbroenne.PowerPointMcp.Generated.ServiceRegistry+Sample");
        Assert.NotNull(category);
        var actionRoute = Assert.Single(category.GetMembers("RouteAction").OfType<IMethodSymbol>());
        var cliRoute = Assert.Single(category.GetMembers("RouteCliArgs").OfType<IMethodSymbol>());
        Assert.Equal(3 + exposedParameterCount, actionRoute.Parameters.Length);
        Assert.Equal(1 + exposedParameterCount, cliRoute.Parameters.Length);
    }
}
