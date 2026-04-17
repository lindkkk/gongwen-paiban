using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniMaxAIDocx.Core;

if (args.Length == 0)
{
    Console.WriteLine("Usage: dotnet run --project MiniMaxAIDocx.Cli -- <command> [options]");
    Console.WriteLine();
    Console.WriteLine("Commands:");
    Console.WriteLine("  format <input.docx> [-o <output.docx>]  - 排版公文文档");
    Console.WriteLine("  convert <input.doc> [-o <output.docx>]  - 转换DOC为DOCX");
    Console.WriteLine("  preview <input.docx>                    - 预览文档内容");
    return;
}

var command = args[0];

try
{
    switch (command)
    {
        case "format":
            await FormatDocument(args.Skip(1).ToArray());
            break;
        case "convert":
            await ConvertDocument(args.Skip(1).ToArray());
            break;
        case "preview":
            await PreviewDocument(args.Skip(1).ToArray());
            break;
        case "test-marker":
            TestMarker(args.Skip(1).ToArray());
            break;
        default:
            Console.WriteLine($"Unknown command: {command}");
            return;
    }
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"Error: {ex.Message}");
    Console.ResetColor();
    Environment.Exit(1);
}

async Task FormatDocument(string[] args)
{
    string inputPath = "";
    string? outputPath = null;
    var options = new FormatOptions();

    for (int i = 0; i < args.Length; i++)
    {
        string a = args[i];
        string? next() => (i + 1 < args.Length) ? args[++i] : null;
        if (a == "-o" || a == "--output")          { outputPath = next(); }
        else if (a == "--source")                  { options.Source = next() ?? "auto"; }
        else if (a == "--h1-marker")               { options.H1MarkerSample = next(); }
        else if (a == "--h2-marker")               { options.H2MarkerSample = next(); }
        else if (a == "--h3-marker")               { options.H3MarkerSample = next(); }
        else if (!a.StartsWith("-"))               { inputPath = a; }
    }

    if (string.IsNullOrEmpty(inputPath))
    {
        Console.WriteLine("Error: Input file path is required");
        Environment.Exit(1);
        return;
    }

    if (string.IsNullOrEmpty(outputPath))
        outputPath = Path.ChangeExtension(inputPath, ".formatted.docx");

    Console.WriteLine($"Formatting document: {inputPath}");
    Console.WriteLine($"Output: {outputPath}");
    Console.WriteLine($"Source hint: {options.Source}");
    if (!string.IsNullOrEmpty(options.H1MarkerSample)) Console.WriteLine($"H1 marker : {options.H1MarkerSample}");
    if (!string.IsNullOrEmpty(options.H2MarkerSample)) Console.WriteLine($"H2 marker : {options.H2MarkerSample}");
    if (!string.IsNullOrEmpty(options.H3MarkerSample)) Console.WriteLine($"H3 marker : {options.H3MarkerSample}");

    File.Copy(inputPath, outputPath, true);

    using var doc = WordprocessingDocument.Open(outputPath, true);
    var formatter = new GongWenFormatter();
    formatter.FormatDocument(doc, options);
    doc.Save();

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Document formatted successfully!");
    Console.ResetColor();
}

async Task ConvertDocument(string[] args)
{
    string inputPath = "";
    string? outputPath = null;

    for (int i = 0; i < args.Length; i++)
    {
        if (args[i] == "-o" && i + 1 < args.Length)
        {
            outputPath = args[i + 1];
            i++;
        }
        else if (!args[i].StartsWith("-"))
        {
            inputPath = args[i];
        }
    }

    if (string.IsNullOrEmpty(inputPath))
    {
        Console.WriteLine("Error: Input file path is required");
        return;
    }

    // 如果没有指定输出，生成默认路径
    if (string.IsNullOrEmpty(outputPath))
    {
        outputPath = Path.ChangeExtension(inputPath, ".docx");
    }

    Console.WriteLine($"Converting document: {inputPath}");
    Console.WriteLine($"Output: {outputPath}");

    // DOC转DOCX需要使用旧版格式
    // 这里我们使用一个简化的方法 - 实际上需要用Word或者更复杂的库
    Console.WriteLine("Note: Full DOC to DOCX conversion requires Microsoft Word or a specialized library.");
    Console.WriteLine("Please save your document as .docx format before using the format command.");

    await Task.CompletedTask;
}

void TestMarker(string[] args)
{
    if (args.Length == 0)
    {
        Console.WriteLine("Usage: test-marker <sample> [<text-to-match>...]");
        return;
    }
    var regex = MarkerPatternInferrer.Infer(args[0]);
    Console.WriteLine($"sample:  {args[0]}");
    Console.WriteLine($"regex :  {regex?.ToString() ?? "(null)"}");
    for (int i = 1; i < args.Length; i++)
    {
        bool m = regex != null && regex.IsMatch(args[i]);
        Console.WriteLine($"  match '{args[i]}'  → {m}");
    }
}

async Task PreviewDocument(string[] args)
{
    if (args.Length == 0)
    {
        Console.WriteLine("Error: Input file path is required");
        return;
    }

    var inputPath = args[0];

    if (!File.Exists(inputPath))
    {
        Console.WriteLine($"Error: File not found: {inputPath}");
        return;
    }

    Console.WriteLine($"Previewing document: {inputPath}");
    Console.WriteLine(new string('=', 60));

    using var doc = WordprocessingDocument.Open(inputPath, false);
    var body = doc.MainDocumentPart?.Document?.Body;

    if (body == null)
    {
        Console.WriteLine("Error: Unable to read document body");
        return;
    }

    int paraNum = 0;
    foreach (var para in body.Elements<Paragraph>())
    {
        paraNum++;
        var text = para.InnerText.Trim();
        if (!string.IsNullOrEmpty(text))
        {
            var preview = text.Length > 80 ? text[..80] + "..." : text;
            Console.WriteLine($"[{paraNum}] {preview}");
        }
    }

    await Task.CompletedTask;
}
