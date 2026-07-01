using Sbroenne.PowerPointMcp.Core.Presentation;

if (args.Length < 2 || args[0] != "create")
{
    Console.WriteLine("Usage: powerpointcli create <path-to-new.pptx>");
    Console.WriteLine();
    Console.WriteLine("This is a minimal placeholder CLI proving the ComInterop -> Core -> CLI");
    Console.WriteLine("vertical slice end to end. Only the 'create' presentation-lifecycle command");
    Console.WriteLine("is implemented so far — see plan.md for the full command roadmap (slides,");
    Console.WriteLine("shapes, text, tables, charts, images, notes, layouts, export/QA).");
    return 1;
}

var commands = new PresentationCommands();

try
{
    var result = commands.Create(args[1]);
    if (result.Success)
    {
        Console.WriteLine($"Created: {result.PresentationPath}");
        return 0;
    }

    Console.Error.WriteLine($"Error: {result.ErrorMessage}");
    return 1;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}
