using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Slide;

void PrintUsage()
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  powerpointcli create <path-to-new.pptx>");
    Console.WriteLine("  powerpointcli slide add-blank <path-to.pptx>");
    Console.WriteLine("  powerpointcli slide count <path-to.pptx>");
    Console.WriteLine("  powerpointcli slide delete <path-to.pptx> <slideIndex>");
    Console.WriteLine();
    Console.WriteLine("This is a minimal placeholder CLI proving the ComInterop -> Core -> CLI");
    Console.WriteLine("vertical slice end to end — not the real Generators-based CLI used by");
    Console.WriteLine("mcp-server-excel. See plan.md for the full command roadmap (shapes, text,");
    Console.WriteLine("tables, charts, images, notes, layouts, export/QA).");
}

if (args.Length == 0)
{
    PrintUsage();
    return 1;
}

var presentationCommands = new PresentationCommands();
var slideCommands = new SlideCommands();

try
{
    switch (args[0])
    {
        case "create" when args.Length >= 2:
        {
            var result = presentationCommands.Create(args[1]);
            if (!result.Success)
            {
                Console.Error.WriteLine($"Error: {result.ErrorMessage}");
                return 1;
            }
            Console.WriteLine($"Created: {result.PresentationPath}");
            return 0;
        }

        case "slide" when args.Length >= 3 && args[1] == "add-blank":
        {
            using var batch = PresentationSession.BeginBatch(args[2]);
            var result = slideCommands.AddBlank(batch);
            if (!result.Success)
            {
                Console.Error.WriteLine($"Error: {result.ErrorMessage}");
                return 1;
            }
            presentationCommands.Save(batch);
            Console.WriteLine($"Added slide {result.SlideIndex}. Total slides: {result.SlideCount}");
            return 0;
        }

        case "slide" when args.Length >= 3 && args[1] == "count":
        {
            using var batch = PresentationSession.BeginBatch(args[2]);
            var result = slideCommands.GetCount(batch);
            if (!result.Success)
            {
                Console.Error.WriteLine($"Error: {result.ErrorMessage}");
                return 1;
            }
            Console.WriteLine($"{result.SlideCount}");
            return 0;
        }

        case "slide" when args.Length >= 4 && args[1] == "delete":
        {
            if (!int.TryParse(args[3], out int slideIndex))
            {
                Console.Error.WriteLine($"Error: '{args[3]}' is not a valid slide index.");
                return 1;
            }

            using var batch = PresentationSession.BeginBatch(args[2]);
            var result = slideCommands.Delete(batch, slideIndex);
            if (!result.Success)
            {
                Console.Error.WriteLine($"Error: {result.ErrorMessage}");
                return 1;
            }
            presentationCommands.Save(batch);
            Console.WriteLine($"Deleted slide {slideIndex}. Total slides: {result.SlideCount}");
            return 0;
        }

        default:
            PrintUsage();
            return 1;
    }
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}
