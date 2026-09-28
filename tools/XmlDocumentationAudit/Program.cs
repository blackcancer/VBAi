using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Xml.Linq;

if (args.Length != 1)
{
    Console.Error.WriteLine("Usage: XmlDocumentationAudit <source-directory>");
    return 2;
}

var root = Path.GetFullPath(args[0]);
var files = Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)
    .Where(path => !path.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar)
        && !path.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar))
    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
    .ToArray();
var missing = new List<(string File, int Line, string Kind, string Name, string[] Tags)>();
var kinds = new SortedDictionary<string, int>(StringComparer.Ordinal);
var parseErrors = new List<(string File, string Message)>();

foreach (var file in files)
{
    var tree = CSharpSyntaxTree.ParseText(File.ReadAllText(file), new CSharpParseOptions(LanguageVersion.Latest));
    foreach (var diagnostic in tree.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error))
        parseErrors.Add((Path.GetRelativePath(root, file), diagnostic.ToString()));

    foreach (var node in tree.GetRoot().DescendantNodes().Where(IsDeclaration))
    {
        IEnumerable<(SyntaxNode Node, string Name)> entries = node is FieldDeclarationSyntax field
            ? field.Declaration.Variables.Select(variable => ((SyntaxNode)node, variable.Identifier.ValueText))
            : new[] { (node, Name(node)) };
        foreach (var entry in entries)
        {
            var kind = entry.Node.Kind().ToString();
            kinds[kind] = kinds.GetValueOrDefault(kind) + 1;
            var tags = MissingTags(entry.Node);
            if (tags.Length > 0)
                missing.Add((Path.GetRelativePath(root, file),
                    tree.GetLineSpan(entry.Node.Span).StartLinePosition.Line + 1,
                    kind, entry.Name, tags));
        }
    }
}

var total = kinds.Values.Sum();
Console.WriteLine($"Source files: {files.Length}");
Console.WriteLine($"Syntax errors: {parseErrors.Count}");
Console.WriteLine($"Documentable declarations: {total}");
Console.WriteLine($"Covered declarations: {total - missing.Count}");
Console.WriteLine($"Missing or invalid: {missing.Count}");
foreach (var pair in kinds)
    Console.WriteLine($"  {pair.Key}: {pair.Value}");
foreach (var error in parseErrors)
    Console.WriteLine($"SYNTAX {error.File}: {error.Message}");
foreach (var entry in missing)
    Console.WriteLine($"MISSING {entry.File}:{entry.Line} {entry.Kind} {entry.Name} [{string.Join(", ", entry.Tags)}]");
return missing.Count == 0 && parseErrors.Count == 0 ? 0 : 1;

static bool IsDeclaration(SyntaxNode node) => node is BaseTypeDeclarationSyntax or DelegateDeclarationSyntax
    or MethodDeclarationSyntax or ConstructorDeclarationSyntax or DestructorDeclarationSyntax
    or OperatorDeclarationSyntax or ConversionOperatorDeclarationSyntax or PropertyDeclarationSyntax
    or IndexerDeclarationSyntax or EventDeclarationSyntax or FieldDeclarationSyntax or EnumMemberDeclarationSyntax;

static string[] MissingTags(SyntaxNode node)
{
    var trivia = node.GetLeadingTrivia().Where(item =>
        item.IsKind(SyntaxKind.SingleLineDocumentationCommentTrivia)
        || item.IsKind(SyntaxKind.MultiLineDocumentationCommentTrivia)).ToArray();
    if (trivia.Length == 0)
        return ["summary"];

    var source = string.Concat(trivia.Select(item => item.ToFullString()
        .Replace("///", "").Replace("/**", "").Replace("*/", "")));
    XElement xml;
    try { xml = XElement.Parse("<root>" + source + "</root>"); }
    catch { return ["valid XML"]; }

    var result = new List<string>();
    foreach (var tag in new[] { "summary", "returns", "value" })
        if (xml.Elements(tag).Count() > 1)
            result.Add("duplicate " + tag);
    if (!HasText(xml, "summary"))
        result.Add("summary");

    IEnumerable<string> typeParameters = node switch
    {
        TypeDeclarationSyntax type => type.TypeParameterList?.Parameters.Select(parameter => parameter.Identifier.ValueText) ?? [],
        DelegateDeclarationSyntax method => method.TypeParameterList?.Parameters.Select(parameter => parameter.Identifier.ValueText) ?? [],
        MethodDeclarationSyntax method => method.TypeParameterList?.Parameters.Select(parameter => parameter.Identifier.ValueText) ?? [],
        _ => []
    };
    var typeParameterNames = typeParameters.ToArray();
    foreach (var parameter in typeParameterNames)
        if (!HasNamedText(xml, "typeparam", parameter))
            result.Add("typeparam " + parameter);
    foreach (var item in xml.Elements("typeparam"))
    {
        var name = (string?)item.Attribute("name");
        if (name == null || !typeParameterNames.Contains(name, StringComparer.Ordinal))
            result.Add("unexpected typeparam " + (name ?? "(sans nom)"));
        if (xml.Elements("typeparam").Count(other => (string?)other.Attribute("name") == name) > 1)
            result.Add("duplicate typeparam " + name);
    }

    IEnumerable<string> parameters = node switch
    {
        MethodDeclarationSyntax method => method.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText),
        ConstructorDeclarationSyntax method => method.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText),
        DelegateDeclarationSyntax method => method.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText),
        IndexerDeclarationSyntax indexer => indexer.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText),
        _ => []
    };
    var parameterNames = parameters.ToArray();
    foreach (var parameter in parameterNames)
        if (!HasNamedText(xml, "param", parameter))
            result.Add("param " + parameter);
    foreach (var item in xml.Elements("param"))
    {
        var name = (string?)item.Attribute("name");
        if (name == null || !parameterNames.Contains(name, StringComparer.Ordinal))
            result.Add("unexpected param " + (name ?? "(sans nom)"));
        if (xml.Elements("param").Count(other => (string?)other.Attribute("name") == name) > 1)
            result.Add("duplicate param " + name);
    }

    var hasReturn = node switch
    {
        MethodDeclarationSyntax method => method.ReturnType.ToString() != "void",
        DelegateDeclarationSyntax method => method.ReturnType.ToString() != "void",
        OperatorDeclarationSyntax or ConversionOperatorDeclarationSyntax => true,
        _ => false
    };
    if (hasReturn && !HasText(xml, "returns"))
        result.Add("returns");
    if (!hasReturn && xml.Elements("returns").Any())
        result.Add("unexpected returns");
    if (node is PropertyDeclarationSyntax or IndexerDeclarationSyntax && !HasText(xml, "value"))
        result.Add("value");
    if (node is not PropertyDeclarationSyntax and not IndexerDeclarationSyntax && xml.Elements("value").Any())
        result.Add("unexpected value");
    return [.. result];
}

static bool HasText(XElement xml, string element) =>
    xml.Elements(element).Any(item => !string.IsNullOrWhiteSpace(item.Value));

static bool HasNamedText(XElement xml, string element, string name) =>
    xml.Elements(element).Any(item => (string?)item.Attribute("name") == name && !string.IsNullOrWhiteSpace(item.Value));

static string Name(SyntaxNode node) => node switch
{
    BaseTypeDeclarationSyntax type => type.Identifier.ValueText,
    DelegateDeclarationSyntax method => method.Identifier.ValueText,
    MethodDeclarationSyntax method => method.Identifier.ValueText,
    BaseMethodDeclarationSyntax method => method switch
    {
        ConstructorDeclarationSyntax constructor => constructor.Identifier.ValueText,
        DestructorDeclarationSyntax destructor => destructor.Identifier.ValueText,
        _ => "operator"
    },
    PropertyDeclarationSyntax property => property.Identifier.ValueText,
    IndexerDeclarationSyntax => "this[]",
    EventDeclarationSyntax eventDeclaration => eventDeclaration.Identifier.ValueText,
    FieldDeclarationSyntax field => string.Join(",", field.Declaration.Variables.Select(variable => variable.Identifier.ValueText)),
    EnumMemberDeclarationSyntax member => member.Identifier.ValueText,
    _ => node.Kind().ToString()
};
