using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Xml.Linq;

if (args.Length != 1
    && (args.Length != 3 || args[1] is not ("--compare" or "--transplant"))
    && (args.Length != 2 || args[1] is not ("--complete" or "--refine")))
{
    Console.Error.WriteLine("Usage: XmlDocumentationAudit <source-directory> [--compare <baseline-source-directory> | --transplant <documented-source-directory> | --complete | --refine]");
    return 2;
}

var root = Path.GetFullPath(args[0]);
if (args.Length == 3 && args[1] == "--transplant")
    return TransplantDocumentation(root, Path.GetFullPath(args[2]));
if (args.Length == 2 && args[1] == "--complete")
    return CompleteDocumentation(root);
if (args.Length == 2 && args[1] == "--refine")
    return RefineDocumentation(root);
var baselineRoot = args.Length == 3 ? Path.GetFullPath(args[2]) : null;
var files = Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)
    .Where(path => !path.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar)
        && !path.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar))
    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
    .ToArray();
var missing = new List<(string File, int Line, string Kind, string Name, string[] Tags)>();
var kinds = new SortedDictionary<string, int>(StringComparer.Ordinal);
var parseErrors = new List<(string File, string Message)>();
var syntaxDifferences = new List<string>();

foreach (var file in files)
{
    var tree = CSharpSyntaxTree.ParseText(File.ReadAllText(file), new CSharpParseOptions(LanguageVersion.Latest));
    foreach (var diagnostic in tree.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error))
        parseErrors.Add((Path.GetRelativePath(root, file), diagnostic.ToString()));

    if (baselineRoot != null)
    {
        var relative = Path.GetRelativePath(root, file);
        var baselineFile = Path.Combine(baselineRoot, relative);
        if (!File.Exists(baselineFile))
            syntaxDifferences.Add(relative + " (absent du baseline)");
        else
        {
            var baselineTree = CSharpSyntaxTree.ParseText(File.ReadAllText(baselineFile),
                new CSharpParseOptions(LanguageVersion.Latest));
            if (!SyntaxFactory.AreEquivalent(baselineTree.GetRoot(), tree.GetRoot()))
                syntaxDifferences.Add(relative);
        }
    }

    foreach (var node in tree.GetRoot().DescendantNodes().Where(IsDeclaration))
    {
        IEnumerable<(SyntaxNode Node, string Name)> entries = node is BaseFieldDeclarationSyntax field
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
if (baselineRoot != null)
    Console.WriteLine($"Syntax differences from baseline (trivia ignored): {syntaxDifferences.Count}");
Console.WriteLine($"Documentable declarations: {total}");
Console.WriteLine($"Covered declarations: {total - missing.Count}");
Console.WriteLine($"Missing or invalid: {missing.Count}");
foreach (var pair in kinds)
    Console.WriteLine($"  {pair.Key}: {pair.Value}");
foreach (var error in parseErrors)
    Console.WriteLine($"SYNTAX {error.File}: {error.Message}");
foreach (var difference in syntaxDifferences)
    Console.WriteLine($"SYNTAX-DIFF {difference}");
foreach (var entry in missing)
    Console.WriteLine($"MISSING {entry.File}:{entry.Line} {entry.Kind} {entry.Name} [{string.Join(", ", entry.Tags)}]");
return missing.Count == 0 && parseErrors.Count == 0 && syntaxDifferences.Count == 0 ? 0 : 1;

static bool IsDeclaration(SyntaxNode node) => node is BaseTypeDeclarationSyntax or DelegateDeclarationSyntax
    or MethodDeclarationSyntax or ConstructorDeclarationSyntax or DestructorDeclarationSyntax
    or OperatorDeclarationSyntax or ConversionOperatorDeclarationSyntax or PropertyDeclarationSyntax
    or IndexerDeclarationSyntax or EventDeclarationSyntax or FieldDeclarationSyntax
    or EventFieldDeclarationSyntax or EnumMemberDeclarationSyntax;

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
    BaseFieldDeclarationSyntax field => string.Join(",", field.Declaration.Variables.Select(variable => variable.Identifier.ValueText)),
    EnumMemberDeclarationSyntax member => member.Identifier.ValueText,
    _ => node.Kind().ToString()
};

static int TransplantDocumentation(string targetRoot, string sourceRoot)
{
    var filesChanged = 0;
    var commentsCopied = 0;
    foreach (var targetFile in Directory.EnumerateFiles(targetRoot, "*.cs", SearchOption.AllDirectories)
        .Where(path => !path.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar)
            && !path.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar)))
    {
        var relative = Path.GetRelativePath(targetRoot, targetFile);
        var sourceFile = Path.Combine(sourceRoot, relative);
        if (!File.Exists(sourceFile)) continue;

        var targetText = File.ReadAllText(targetFile);
        var targetTree = CSharpSyntaxTree.ParseText(targetText, new CSharpParseOptions(LanguageVersion.Latest));
        var sourceTree = CSharpSyntaxTree.ParseText(File.ReadAllText(sourceFile), new CSharpParseOptions(LanguageVersion.Latest));
        var docsByKey = sourceTree.GetRoot().DescendantNodes().Where(IsDeclaration)
            .SelectMany(DeclarationEntries)
            .Select(entry => (entry.Key, Documentation: DocumentationTrivia(entry.Node)))
            .Where(entry => entry.Documentation.Count > 0)
            .GroupBy(entry => entry.Key, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First().Documentation, StringComparer.Ordinal);

        var root = targetTree.GetRoot();
        var replacements = new Dictionary<SyntaxNode, SyntaxNode>();
        foreach (var declaration in root.DescendantNodes().Where(IsDeclaration))
        {
            if (DocumentationTrivia(declaration).Count > 0 && MissingTags(declaration).Length == 0) continue;
            var key = DeclarationEntries(declaration).Select(entry => entry.Key)
                .FirstOrDefault(docsByKey.ContainsKey);
            if (key is null) continue;
            var leading = declaration.GetLeadingTrivia().Where(trivia => !IsDocumentationTrivia(trivia));
            replacements[declaration] = declaration.WithLeadingTrivia(
                leading.Concat(AlignParameterNames(docsByKey[key], declaration)));
            commentsCopied++;
        }

        if (replacements.Count == 0) continue;
        var updated = root.ReplaceNodes(replacements.Keys,
            (original, rewritten) => rewritten.WithLeadingTrivia(replacements[original].GetLeadingTrivia())).ToFullString();
        if (StringComparer.Ordinal.Equals(targetText, updated)) continue;
        File.WriteAllText(targetFile, updated, new System.Text.UTF8Encoding(false));
        filesChanged++;
    }

    Console.WriteLine($"Files changed: {filesChanged}");
    Console.WriteLine($"Documentation comments transplanted: {commentsCopied}");
    return 0;
}

static IEnumerable<(string Key, SyntaxNode Node)> DeclarationEntries(SyntaxNode node)
{
    if (node is BaseFieldDeclarationSyntax field)
    {
        foreach (var variable in field.Declaration.Variables)
            yield return (DeclarationKey(node, variable.Identifier.ValueText), node);
    }
    else yield return (DeclarationKey(node, Name(node)), node);
}

static string DeclarationKey(SyntaxNode node, string name)
{
    var scope = string.Join(".", node.Ancestors().Reverse()
        .Where(ancestor => ancestor is BaseNamespaceDeclarationSyntax or BaseTypeDeclarationSyntax)
        .Select(ancestor => ancestor switch
        {
            BaseNamespaceDeclarationSyntax ns => ns.Name.ToString(),
            BaseTypeDeclarationSyntax type => type.Identifier.ValueText,
            _ => ""
        }));
    var signature = node switch
    {
        MethodDeclarationSyntax method => method.TypeParameterList?.ToString() + Parameters(method.ParameterList),
        ConstructorDeclarationSyntax constructor => Parameters(constructor.ParameterList),
        DelegateDeclarationSyntax method => method.TypeParameterList?.ToString() + Parameters(method.ParameterList),
        IndexerDeclarationSyntax indexer => Parameters(indexer.ParameterList),
        OperatorDeclarationSyntax method => Parameters(method.ParameterList),
        ConversionOperatorDeclarationSyntax method => method.Type + Parameters(method.ParameterList),
        _ => ""
    };
    var explicitInterface = node switch
    {
        MethodDeclarationSyntax method => method.ExplicitInterfaceSpecifier?.Name.ToString(),
        PropertyDeclarationSyntax property => property.ExplicitInterfaceSpecifier?.Name.ToString(),
        EventDeclarationSyntax eventMember => eventMember.ExplicitInterfaceSpecifier?.Name.ToString(),
        _ => null
    };
    return $"{scope}|{node.Kind()}|{explicitInterface}|{name}|{signature}";
}

static string Parameters(BaseParameterListSyntax parameters) => "(" + string.Join(",", parameters.Parameters.Select(parameter =>
    $"{parameter.Modifiers}:{parameter.Type}")) + ")";

static List<SyntaxTrivia> AlignParameterNames(List<SyntaxTrivia> documentation, SyntaxNode target)
{
    var names = target switch
    {
        MethodDeclarationSyntax method => method.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText).ToArray(),
        ConstructorDeclarationSyntax constructor => constructor.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText).ToArray(),
        DelegateDeclarationSyntax method => method.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText).ToArray(),
        IndexerDeclarationSyntax indexer => indexer.ParameterList.Parameters.Select(parameter => parameter.Identifier.ValueText).ToArray(),
        _ => []
    };
    if (names.Length == 0) return documentation;

    var index = 0;
    var text = string.Concat(documentation.Select(trivia => trivia.ToFullString()));
    text = System.Text.RegularExpressions.Regex.Replace(text,
        "(<param\\s+name=\")[^\"]*(\")",
        match => index < names.Length ? match.Groups[1].Value + System.Security.SecurityElement.Escape(names[index++]) + match.Groups[2].Value : match.Value);
    return SyntaxFactory.ParseLeadingTrivia(text).ToList();
}

static bool IsDocumentationTrivia(SyntaxTrivia trivia) => trivia.IsKind(SyntaxKind.SingleLineDocumentationCommentTrivia)
    || trivia.IsKind(SyntaxKind.MultiLineDocumentationCommentTrivia);

static List<SyntaxTrivia> DocumentationTrivia(SyntaxNode node) => node.GetLeadingTrivia()
    .Where(IsDocumentationTrivia).ToList();

static int CompleteDocumentation(string root)
{
    var declarationsCompleted = 0;
    var filesChanged = 0;
    foreach (var file in Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)
        .Where(path => !path.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar)
            && !path.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar)))
    {
        var text = File.ReadAllText(file);
        var tree = CSharpSyntaxTree.ParseText(text, new CSharpParseOptions(LanguageVersion.Latest));
        var sourceRoot = tree.GetRoot();
        var replacements = new Dictionary<SyntaxNode, SyntaxNode>();
        foreach (var node in sourceRoot.DescendantNodes().Where(IsDeclaration))
        {
            var issues = MissingTags(node);
            if (issues.Length == 0) continue;
            var existing = DocumentationTrivia(node);
            var canPreserve = existing.Count > 0 && IsValidDocumentation(existing);
            var xml = new System.Text.StringBuilder();
            if (!canPreserve || issues.Contains("summary"))
                xml.AppendLine("/// <summary>" + EscapeXml(SummaryFor(node)) + "</summary>");
            foreach (var typeParameter in TypeParameters(node).Where(name => !canPreserve || issues.Contains("typeparam " + name)))
                xml.AppendLine($"/// <typeparam name=\"{EscapeXml(typeParameter)}\">The type used for {EscapeXml(Humanize(typeParameter))}.</typeparam>");
            foreach (var parameter in ParametersFor(node).Where(parameter => !canPreserve || issues.Contains("param " + parameter.Name)))
                xml.AppendLine($"/// <param name=\"{EscapeXml(parameter.Name)}\">{EscapeXml(ParameterDescription(parameter))}</param>");
            if (HasReturnValue(node) && (!canPreserve || issues.Contains("returns")))
                xml.AppendLine("/// <returns>The result produced by this operation.</returns>");
            if (node is PropertyDeclarationSyntax or IndexerDeclarationSyntax && (!canPreserve || issues.Contains("value")))
                xml.AppendLine("/// <value>The current value represented by this member.</value>");
            var leading = node.GetLeadingTrivia().Where(trivia => !IsDocumentationTrivia(trivia))
                .Concat(canPreserve ? existing : [])
                .Concat(SyntaxFactory.ParseLeadingTrivia(xml.ToString()));
            replacements[node] = node.WithLeadingTrivia(leading);
            declarationsCompleted++;
        }
        if (replacements.Count == 0) continue;
        var updated = sourceRoot.ReplaceNodes(replacements.Keys,
            (original, rewritten) => rewritten.WithLeadingTrivia(replacements[original].GetLeadingTrivia())).ToFullString();
        if (StringComparer.Ordinal.Equals(text, updated)) continue;
        File.WriteAllText(file, updated, new System.Text.UTF8Encoding(false));
        filesChanged++;
    }
    Console.WriteLine($"Files completed: {filesChanged}");
    Console.WriteLine($"Declarations completed: {declarationsCompleted}");
    return 0;
}

static string SummaryFor(SyntaxNode node)
{
    var name = Name(node);
    var owner = node.Ancestors().OfType<BaseTypeDeclarationSyntax>().FirstOrDefault()?.Identifier.ValueText;
    return node switch
    {
        ClassDeclarationSyntax when name == "EditorWorkspaceHost" => "Hosts the modern editor inside the VBE document workspace and follows the active native document window.",
        ClassDeclarationSyntax when name == "ProjectAccessWindow" => "Collects the projects and shared context that the chat assistant may read.",
        ClassDeclarationSyntax => $"Provides the {Humanize(name)} implementation.",
        InterfaceDeclarationSyntax => $"Defines the {Humanize(name)} contract.",
        StructDeclarationSyntax when name == "Rect" => "Describes the edges of the native editor workspace rectangle.",
        StructDeclarationSyntax => $"Represents {Humanize(name)} data.",
        EnumDeclarationSyntax => $"Lists the supported {Humanize(name)} values.",
        DelegateDeclarationSyntax => $"Defines the {Humanize(name)} callback.",
        ConstructorDeclarationSyntax => $"Initializes a {owner ?? name} instance with the supplied state.",
        DestructorDeclarationSyntax => $"Releases resources owned by the {owner ?? name} instance.",
        PropertyDeclarationSyntax property when property.Identifier.ValueText == "Library" && owner == "EditorSymbol"
            => "Gets or sets the type library that defines an external symbol.",
        PropertyDeclarationSyntax property when property.Identifier.ValueText == "EditorDocumentId" && owner == "ChatWorkflow"
            => "Gets the identifier of the editor document associated with this workflow.",
        PropertyDeclarationSyntax property when property.AccessorList?.Accessors.Any(accessor => accessor.IsKind(SyntaxKind.SetAccessorDeclaration) || accessor.IsKind(SyntaxKind.InitAccessorDeclaration)) == true
            => $"Gets or sets the {Humanize(name)}.",
        PropertyDeclarationSyntax => $"Gets the {Humanize(name)}.",
        IndexerDeclarationSyntax => "Gets the value at the requested index.",
        EventDeclarationSyntax or EventFieldDeclarationSyntax => $"Notifies subscribers when {Humanize(name)} occurs.",
        FieldDeclarationSyntax or EnumMemberDeclarationSyntax => $"Stores the {Humanize(name)} used by {owner ?? "this type"}.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "Resize" && owner == "EditorWorkspaceHost"
            => "Matches the hosted editor bounds to the active VBE document workspace and hides it for native designers.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "Show" && owner == "EditorWorkspaceHost"
            => "Shows the editor as a child surface in the VBE document workspace.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ReadAsync" && owner == "BridgeRequestReader"
            => "Reads one newline-terminated UTF-8 request frame within the configured byte and time limits.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ReadFrameAsync" && owner == "BridgeRequestReader"
            => "Reads and decodes one bounded UTF-8 frame terminated by a newline.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "SaveDocument" && owner == "ModernEditorWindow"
            => "Saves the active editor document through its registered native save route.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "SaveInVbe" && owner == "ModernEditorWindow"
            => "Writes the editor buffer back to the corresponding VBE code module.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ExecuteBudgetTool" && owner == "ChatWindow"
            => "Runs a pending tool call and records its result for the paused provider turn.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "PauseBudget" && owner == "ChatWindow"
            => "Persists the current provider turn so it can resume after its execution budget is extended.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "UpdateBudgetControls" && owner == "ChatWindow"
            => "Updates the budget controls to reflect whether the current turn is paused or resumable.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ResumeBudgetAsync" && owner == "ChatWindow"
            => "Resumes the saved provider turn after completing any pending tool responses.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "CompletePendingToolResponses" && owner == "ChatWindow"
            => "Completes persisted tool actions before a paused provider turn resumes.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "RunHttpBudgetAsync" && owner == "ChatWindow"
            => "Runs the HTTP provider turn while applying its configured token and time budgets.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "MigrateProviderPrivacy" && owner == "ChatWindow"
            => "Migrates legacy provider privacy settings to the current project access policy.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ConfigureProjectAccess" && owner == "ChatWindow"
            => "Collects the projects and shared context granted to the current chat session.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "IsCatalogTool" && owner == "LlmVbeTools"
            => "Determines whether a tool name belongs to the provider catalog.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ToolFamily" && owner == "LlmVbeTools"
            => "Returns the catalog family that owns the specified tool.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "CatalogForProvider" && owner == "LlmVbeTools"
            => "Builds the tool catalog available to the selected provider.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ResetCatalog" && owner == "LlmVbeTools"
            => "Clears cached provider catalog selections and family priorities.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "InvokeCatalogAsync" && owner == "LlmVbeTools"
            => "Dispatches a provider catalog tool to its owning tool family.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "SetReadAccess" && owner == "LlmVbeTools"
            => "Sets the project grants and shared context policy for tool reads.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "RequireProjectRead" && owner == "LlmVbeTools"
            => "Rejects a project read when the current session has not granted access.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "IsAuthorizedAlias" && owner == "LlmVbeTools"
            => "Checks whether a requested project alias resolves to an authorized project.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "SameProject" && owner == "LlmVbeTools"
            => "Compares project identities using the names accepted by the live VBE session.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "GuardProjectPrivacy" && owner == "LlmVbeTools"
            => "Guards a tool request against the current project access policy.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "Fields" && owner == "LlmVbeTools"
            => "Reads project fields only after the configured privacy checks pass.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "FilterProjects" && owner == "LlmVbeTools"
            => "Removes projects that are outside the session's explicit read grants.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "FilterProjectResponse" && owner == "LlmVbeTools"
            => "Removes ungranted project data from a tool response before it reaches the model.",
        MethodDeclarationSyntax method when method.Identifier.ValueText == "ScopedLiveSnapshot" && owner == "LlmVbeTools"
            => "Captures live project state within the session's authorized project scope.",
        MethodDeclarationSyntax => $"Performs the {Humanize(name)} operation for {owner ?? "the current context"}.",
        _ => $"Represents {Humanize(name)}."
    };
}

static string ParameterDescription((string Name, string Type) parameter) => parameter.Type switch
{
    "string" => $"Text containing the {Humanize(parameter.Name)}.",
    "bool" => $"Indicates whether {Humanize(parameter.Name)} is enabled.",
    "CancellationToken" => "Token used to cancel the operation.",
    _ => $"The {Humanize(parameter.Name)} used by this operation."
};

static IEnumerable<string> TypeParameters(SyntaxNode node) => node switch
{
    TypeDeclarationSyntax type => type.TypeParameterList?.Parameters.Select(parameter => parameter.Identifier.ValueText) ?? [],
    DelegateDeclarationSyntax method => method.TypeParameterList?.Parameters.Select(parameter => parameter.Identifier.ValueText) ?? [],
    MethodDeclarationSyntax method => method.TypeParameterList?.Parameters.Select(parameter => parameter.Identifier.ValueText) ?? [],
    _ => []
};

static IEnumerable<(string Name, string Type)> ParametersFor(SyntaxNode node)
{
    var parameters = node switch
    {
        MethodDeclarationSyntax method => method.ParameterList.Parameters,
        ConstructorDeclarationSyntax constructor => constructor.ParameterList.Parameters,
        DelegateDeclarationSyntax method => method.ParameterList.Parameters,
        IndexerDeclarationSyntax indexer => indexer.ParameterList.Parameters,
        _ => default
    };
    return parameters.Select(parameter => (parameter.Identifier.ValueText, parameter.Type?.ToString() ?? ""));
}

static bool HasReturnValue(SyntaxNode node) => node switch
{
    MethodDeclarationSyntax method => method.ReturnType.ToString() != "void",
    DelegateDeclarationSyntax method => method.ReturnType.ToString() != "void",
    OperatorDeclarationSyntax or ConversionOperatorDeclarationSyntax => true,
    _ => false
};

static string Humanize(string value)
{
    var words = System.Text.RegularExpressions.Regex.Replace(value, "([a-z0-9])([A-Z])", "$1 $2");
    words = System.Text.RegularExpressions.Regex.Replace(words, "([A-Z])([A-Z][a-z])", "$1 $2");
    return words.Replace('_', ' ').Trim().ToLowerInvariant();
}

static string EscapeXml(string value) => System.Security.SecurityElement.Escape(value) ?? "";

static bool IsValidDocumentation(List<SyntaxTrivia> documentation)
{
    try
    {
        var source = string.Concat(documentation.Select(item => item.ToFullString())
            .Select(value => value.Replace("///", "").Replace("/**", "").Replace("*/", "")));
        _ = XElement.Parse("<root>" + source + "</root>");
        return true;
    }
    catch { return false; }
}

static int RefineDocumentation(string root)
{
    var changedFiles = 0;
    var commentsRefined = 0;
    foreach (var file in Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)
        .Where(path => !path.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar)
            && !path.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar)))
    {
        var text = File.ReadAllText(file);
        var tree = CSharpSyntaxTree.ParseText(text, new CSharpParseOptions(LanguageVersion.Latest));
        var sourceRoot = tree.GetRoot();
        var replacements = new Dictionary<SyntaxNode, SyntaxNode>();
        foreach (var node in sourceRoot.DescendantNodes().Where(IsDeclaration))
        {
            var summary = SpecificSummary(node);
            if (summary is null) continue;
            var documentation = DocumentationTrivia(node);
            if (documentation.Count == 0) continue;
            var oldText = string.Concat(documentation.Select(item => item.ToFullString()));
            var newText = System.Text.RegularExpressions.Regex.Replace(oldText,
                "<summary>.*?</summary>", "<summary>" + EscapeXml(summary) + "</summary>",
                System.Text.RegularExpressions.RegexOptions.Singleline, TimeSpan.FromSeconds(1));
            foreach (var parameter in ParametersFor(node))
            {
                var description = SpecificParameterDescription(node, parameter.Name);
                if (description is null) continue;
                var pattern = "(<param\\s+name=\"" + System.Text.RegularExpressions.Regex.Escape(parameter.Name) + "\">).*?(</param>)";
                newText = System.Text.RegularExpressions.Regex.Replace(newText, pattern,
                    "$1" + EscapeXml(description) + "$2", System.Text.RegularExpressions.RegexOptions.Singleline,
                    TimeSpan.FromSeconds(1));
            }
            var returnDescription = SpecificReturnDescription(node);
            if (returnDescription is not null)
                newText = System.Text.RegularExpressions.Regex.Replace(newText, "<returns>.*?</returns>",
                    "<returns>" + EscapeXml(returnDescription) + "</returns>",
                    System.Text.RegularExpressions.RegexOptions.Singleline, TimeSpan.FromSeconds(1));
            var valueDescription = SpecificValueDescription(node);
            if (valueDescription is not null)
                newText = System.Text.RegularExpressions.Regex.Replace(newText, "<value>.*?</value>",
                    "<value>" + EscapeXml(valueDescription) + "</value>",
                    System.Text.RegularExpressions.RegexOptions.Singleline, TimeSpan.FromSeconds(1));
            if (StringComparer.Ordinal.Equals(oldText, newText)) continue;
            var leading = node.GetLeadingTrivia().Where(trivia => !IsDocumentationTrivia(trivia))
                .Concat(SyntaxFactory.ParseLeadingTrivia(newText));
            replacements[node] = node.WithLeadingTrivia(leading);
            commentsRefined++;
        }
        if (replacements.Count == 0) continue;
        var updated = sourceRoot.ReplaceNodes(replacements.Keys,
            (original, rewritten) => rewritten.WithLeadingTrivia(replacements[original].GetLeadingTrivia())).ToFullString();
        if (StringComparer.Ordinal.Equals(text, updated)) continue;
        File.WriteAllText(file, updated, new System.Text.UTF8Encoding(false));
        changedFiles++;
    }
    Console.WriteLine($"Files refined: {changedFiles}");
    Console.WriteLine($"Documentation summaries refined: {commentsRefined}");
    return 0;
}

static string? SpecificSummary(SyntaxNode node)
{
    var owner = node.Ancestors().OfType<BaseTypeDeclarationSyntax>().FirstOrDefault()?.Identifier.ValueText;
    var name = Name(node);
    return (owner, node.Kind(), name) switch
    {
        (null, SyntaxKind.ClassDeclaration, "QueuedChatMessage") => "Represents a chat message waiting to be sent, with its captured references, attachments, and memory.",
        (null, SyntaxKind.ClassDeclaration, "ChatQueuedMessageView") => "Displays one queued chat message and provides actions to send it, edit it, or remove it.",
        (null, SyntaxKind.ClassDeclaration, "ChatNativeMarkdown") => "Renders provider Markdown in the native chat transcript, including code, tables, links, and VBA references.",
        (null, SyntaxKind.ClassDeclaration, "ChatActivityGroupView") => "Groups related chat activity entries under one expandable transcript section.",
        (null, SyntaxKind.ClassDeclaration, "ChatActivityStepView") => "Displays one chat activity step with its state and expandable details.",
        (null, SyntaxKind.ClassDeclaration, "ChatAttachmentView") => "Displays an attachment included in a chat message.",
        (null, SyntaxKind.ClassDeclaration, "ChatChangeCardView") => "Displays a proposed code change with its diff and available undo actions.",
        (null, SyntaxKind.ClassDeclaration, "ChatLinkView") => "Displays an actionable link in the chat transcript.",
        (null, SyntaxKind.ClassDeclaration, "ChatMessageView") => "Displays a transcript message with its content, metadata, and available actions.",
        (null, SyntaxKind.ClassDeclaration, "ChatFormRecoveryView") => "Displays recoverable changes made to a VBA form.",
        (null, SyntaxKind.ClassDeclaration, "ChatSuggestionsView") => "Displays navigation and reference suggestions for the current chat draft.",
        (null, SyntaxKind.ClassDeclaration, "ChatWelcomeView") => "Displays starter prompts when the chat has no messages.",
        ("ChatNativeMarkdown", _, "Pipeline") => "Parses the Markdown extensions supported by the native transcript renderer.",
        ("ChatNativeMarkdown", _, "Render") => "Parses Markdown and renders its blocks and interactive references into the transcript view.",
        ("ChatNativeMarkdown", _, "Blocks") => "Renders Markdown blocks, including nested lists, tables, quotes, and code, into the transcript view.",
        ("ChatNativeMarkdown", _, "Inlines") => "Renders inline Markdown and wires allowed links and recognized VBA references to their actions.",
        (var chatOwner, _, "components") when chatOwner?.StartsWith("Chat", StringComparison.Ordinal) == true
            => "Container that owns the disposable components created by the WinForms Designer.",
        (var chatOwner, _, "toolTips") when chatOwner?.StartsWith("Chat", StringComparison.Ordinal) == true
            => "ToolTip component used to show full text for transcript controls.",
        (var chatOwner, _, "layout") when chatOwner?.StartsWith("Chat", StringComparison.Ordinal) == true
            => "Flow layout panel that contains this transcript view's child controls.",
        (null, SyntaxKind.ClassDeclaration, "ChatDisclosureView") => "Displays transcript details in a section that can expand or collapse.",
        ("ChatDisclosureView", _, "ContentPanel") => "Gets the flow panel used to add section specific controls in the Designer.",
        ("ChatDisclosureView", _, "UpdateExpansion") => "Synchronizes the disclosure caption and body visibility with the expanded state.",
        (null, SyntaxKind.ClassDeclaration, "ChatTextContentView") => "Displays selectable transcript text and Markdown with clickable references and code copying.",
        ("ChatTextContentView", _, "ErrorHandler") => "Receives errors raised while activating transcript links or actions.",
        ("ChatTextContentView", _, "ownedFonts") => "Tracks fonts created by this view so they can be disposed with it.",
        ("ChatTextContentView", _, "actions") => "Maps rendered character ranges to navigation, link, and code copy actions.",
        (null, SyntaxKind.ClassDeclaration, "TextAction") => "Describes an action associated with a range of rendered transcript text.",
        ("ChatTextContentView", _, "ShowPlain") => "Replaces the transcript content with plain text and applies code or interface formatting.",
        ("ChatTextContentView", _, "ShowMarkdown") => "Renders Markdown into the transcript and installs the callbacks for references and errors.",
        ("ChatTextContentView", _, "OwnFont") => "Returns a cached view owned font matching the requested family, size, and style.",
        ("ChatTextContentView", _, "Append") => "Appends a styled text run to the native transcript control.",
        ("ChatTextContentView", _, "ActivateAt") => "Invokes the transcript action associated with the character at the specified position.",
        ("ChatTextContentView", _, "Copy") => "Copies the supplied transcript text through the native clipboard helper.",
        ("ChatTextContentView", _, "ActionAt") => "Finds the action whose rendered character range contains the specified position.",
        ("ChatTextContentView", _, "ResizeText") => "Measures the rendered text and updates the rich text control height and scroll bars.",
        ("ChatTextContentView", _, "DisposeTextResources") => "Disposes fonts created by this view and clears its font cache.",
        (null, SyntaxKind.ClassDeclaration, "ChatDesignerHost") => "Hosts a transcript view in the WinForms Designer and measures its preferred height.",
        ("ChatDesignerHost", _, "MeasureOverride") => "Measures the hosted transcript view within the available designer width.",
        ("ChatDesignerView", _, "Watch") => "Registers a control whose preferred height should trigger row remeasurement.",
        ("ChatDesignerView", _, "ResizeRows") => "Resizes the designer host rows to fit their current transcript controls.",
        ("ChatWindow", _, "QueuedChatMessage") => "Represents a persisted chat draft awaiting dispatch.",
        ("ChatWindow", _, "PendingMessages") => "Gets the current session's messages waiting for dispatch.",
        ("ChatWindow", _, "QueueComposerMessage") => "Moves the current composer text, references, attachments, and captured memory into the pending queue.",
        ("ChatWindow", _, "RefreshPendingMessages") => "Rebuilds the pending message rows and hooks up their send, edit, and delete actions.",
        ("ChatWindow", _, "DeletePendingMessage") => "Removes a queued message from the current session and persists the updated queue.",
        ("ChatWindow", _, "EditPendingMessage") => "Moves a queued message back into the composer when the current draft is empty.",
        ("ChatWindow", _, "SendPendingNowAsync") => "Dispatches a selected queued message immediately, stopping the active response when needed.",
        ("ChatWindow", _, "DispatchPendingAsync") => "Sends the next queued message after a response completes, unless dispatch is paused or stopped.",
        ("ChatWindow", _, "immediateMessageId") => "Identifies the queued message selected for dispatch immediately after the active response stops.",
        ("ChatWindow", _, "queuedDraftMemory") => "Keeps the selected queued message memory attached while its text is being edited in the composer.",
        ("ChatWindow", _, "SendRequestAsync") => "Builds and sends a chat request, then removes its queued message after dispatch succeeds.",
        ("ChatWindow", _, "CreateActivityStep") => "Creates a transcript row for a tool activity and its displayed state.",
        ("ChatWindow", _, "PrepareRequestAttachments") => "Combines the selected files and captured memory into the request attachment payload.",
        ("ChatWindow", _, "referenceView") => "Hosts the live reference suggestions shown below the composer.",
        ("ChatWindow", _, "pendingMessagesPanel") => "Contains the transcript rows for messages waiting to be sent.",
        ("ChatSessionStore", _, "DecodeSession") => "Deserializes a stored chat session and restores its persisted queue and state.",
        ("ChatMessageView", _, "CopyButton") => "Gets the action button that copies this message's content.",
        ("ChatMessageView", _, "ForkButton") => "Gets the action button that starts a new conversation from this message.",
        ("ChatQueuedMessageView", _, "ShowMessage") => "Displays the queued text and exposes it as a tooltip when it is truncated.",
        ("ChatActivityGroupView", _, "section") => "Contains the expandable section that groups related tool activity rows.",
        ("ChatActivityStepView", _, "detail") => "Displays the detailed text returned for this tool activity step.",
        ("ChatActivityStepView", _, "state") => "Displays the current state of this tool activity step.",
        ("ChatAttachmentView", _, "open") => "Opens the attached file when the user activates its action.",
        ("ChatAttachmentView", _, "text") => "Displays the attachment name and descriptive text.",
        ("ChatChangeCardView", _, "diff") => "Displays the code diff for the proposed change.",
        ("ChatChangeCardView", _, "undo") => "Provides the action that reverts the proposed code change.",
        ("ChatChangeCardView", _, "module") => "Displays the VBA component containing the proposed change.",
        ("ChatChangeCardView", _, "count") => "Displays the number of changed lines or blocks.",
        ("ChatChangeCardView", _, "blocks") => "Contains the individual change blocks that can be undone.",
        ("ChatChangeCardView", _, "undoTurn") => "Provides the action that undoes the complete assistant turn.",
        ("ChatChangeCardView", _, "blockMenu") => "Context menu for actions on an individual code change block.",
        ("ChatChangeCardView", _, "actions") => "Contains the actions available for this change card.",
        ("ChatChangeCardView", _, "state") => "Displays the current status of the proposed code change.",
        ("ChatLinkView", _, "link") => "Displays the transcript link and raises its activation action.",
        ("ChatMessageView", _, "speaker") => "Displays the role or name of the message author.",
        ("ChatMessageView", _, "header") => "Arranges the message author and its available actions.",
        ("ChatMessageView", _, "copy") => "Provides the action that copies the message content.",
        ("ChatMessageView", _, "fork") => "Provides the action that starts a conversation from this message.",
        ("ChatMessageView", _, "headingActions") => "Contains the actions shown alongside the message author.",
        ("ChatMessageView", _, "message") => "Displays the message text and its interactive references.",
        ("ChatMessageView", _, "memory") => "Displays the memory snapshot attached to this message.",
        ("ChatMessageView", _, "attachments") => "Contains the files attached to this message.",
        ("ChatMessageView", _, "references") => "Contains the VBA references associated with this message.",
        ("ChatMessageView", _, "targets") => "Contains the code change targets produced by this message.",
        ("ChatSuggestionsView", _, "targets") => "Contains the navigation targets offered for the current composer text.",
        ("ChatSuggestionsView", _, "status") => "Displays the current reference suggestion status.",
        ("ChatWelcomeView", _, "title") => "Displays the heading above the starter prompts.",
        ("ChatWelcomeView", _, "hint") => "Displays guidance for starting a chat request.",
        ("ChatWelcomeView", _, "explain") => "Starts a prompt that asks the assistant to explain selected code.",
        ("ChatWelcomeView", _, "fix") => "Starts a prompt that asks the assistant to find and fix an issue.",
        ("ChatWelcomeView", _, "improve") => "Starts a prompt that asks the assistant to improve selected code.",
        ("ChatFormRecoveryView", _, "title") => "Displays the form recovery prompt.",
        ("ChatFormRecoveryView", _, "count") => "Displays the number of recoverable form changes.",
        ("ChatFormRecoveryView", _, "recover") => "Provides the action that restores the recovered form changes.",
        ("ChatDisclosureView", _, "toggle") => "Button that expands or collapses the section body.",
        ("ChatDisclosureView", _, "body") => "Panel that contains the controls shown while the section is expanded.",
        ("ChatSessionStore", _, "PendingMessages") => "Gets or sets the messages persisted for later dispatch in this chat session.",
        ("ChatSessionStore", _, "DraftCapturedMemory") => "Gets or sets the memory snapshot captured with the current unsent draft.",
        (_, _, "InitializeComponent") when owner?.StartsWith("Chat", StringComparison.Ordinal) == true
            => $"Creates and configures the {Humanize(owner)} controls serialized by the WinForms Designer.",
        _ => null
    };
}

static string? SpecificParameterDescription(SyntaxNode node, string parameter)
{
    var owner = node.Ancestors().OfType<BaseTypeDeclarationSyntax>().FirstOrDefault()?.Identifier.ValueText;
    var method = Name(node);
    return (owner, method, parameter) switch
    {
        ("ChatNativeMarkdown", "Render", "view") => "Transcript control that receives the rendered content.",
        ("ChatNativeMarkdown", "Render", "text") => "Markdown source returned by the chat provider.",
        ("ChatNativeMarkdown", "Render", "refs") => "Recognized VBA references keyed by their visible token.",
        ("ChatNativeMarkdown", "Render", "navigate") => "Callback invoked when the user activates a recognized VBA reference.",
        ("ChatNativeMarkdown", "Render", "error") => "Callback that receives link activation errors.",
        ("ChatNativeMarkdown", "Blocks", "blocks") => "Parsed Markdown blocks to render.",
        ("ChatNativeMarkdown", "Inlines", "source") => "Parsed inline Markdown content to render.",
        ("ChatTextContentView", "ShowPlain", "text") => "Plain message text to display.",
        ("ChatTextContentView", "ShowPlain", "code") => "Whether to use code formatting for the text.",
        ("ChatTextContentView", "ShowMarkdown", "text") => "Markdown source to render in the transcript.",
        ("ChatTextContentView", "ShowMarkdown", "references") => "VBA references that should become interactive transcript links.",
        ("ChatTextContentView", "ShowMarkdown", "navigate") => "Callback used to open a recognized VBA reference.",
        ("ChatTextContentView", "ShowMarkdown", "error") => "Callback used to report link and action errors.",
        ("ChatTextContentView", "OwnFont", "family") => "Font family to reuse or create.",
        ("ChatTextContentView", "OwnFont", "size") => "Font size in points.",
        ("ChatTextContentView", "OwnFont", "style") => "Font style to apply.",
        ("ChatTextContentView", "Append", "text") => "Text to append to the transcript.",
        ("ChatTextContentView", "Append", "family") => "Font family for the appended text.",
        ("ChatTextContentView", "Append", "size") => "Font size in points.",
        ("ChatTextContentView", "Append", "style") => "Font style for the appended text.",
        ("ChatTextContentView", "Append", "color") => "Optional foreground color for the appended text.",
        ("ChatTextContentView", "ActivateAt", "index") => "Character position whose associated action should be invoked.",
        ("ChatTextContentView", "Copy", "text") => "Text to place on the clipboard.",
        ("ChatTextContentView", "ActionAt", "index") => "Character position to test against the rendered action ranges.",
        ("ChatQueuedMessageView", "ShowMessage", "text") => "Text of the queued message.",
        ("ChatWindow", "DeletePendingMessage", "item") => "Queued message to remove.",
        ("ChatWindow", "EditPendingMessage", "item") => "Queued message to move back into the composer.",
        ("ChatWindow", "SendPendingNowAsync", "item") => "Queued message to dispatch ahead of other pending messages.",
        _ => null
    };
}

static string? SpecificReturnDescription(SyntaxNode node) =>
    (node.Ancestors().OfType<BaseTypeDeclarationSyntax>().FirstOrDefault()?.Identifier.ValueText, Name(node)) switch
    {
        ("ChatTextContentView", "OwnFont") => "The matching font instance, created and retained by this view when necessary.",
        ("ChatTextContentView", "ActionAt") => "The action covering the character position, or null when no action covers it.",
        ("ChatDesignerView", "ResizeRows") => "The combined preferred height of the watched transcript controls.",
        ("ChatDesignerHost", "MeasureOverride") => "The measured size required by the hosted transcript view.",
        _ => null
    };

static string? SpecificValueDescription(SyntaxNode node) =>
    (node.Ancestors().OfType<BaseTypeDeclarationSyntax>().FirstOrDefault()?.Identifier.ValueText, Name(node)) switch
    {
        ("QueuedChatMessage", "Id") => "Stable identifier used to select this message in the pending queue.",
        ("QueuedChatMessage", "Text") => "User authored prompt text awaiting dispatch.",
        ("QueuedChatMessage", "References") => "VBA references captured from the prompt when it was queued.",
        ("QueuedChatMessage", "Attachments") => "Files captured from the composer when this message was queued.",
        ("QueuedChatMessage", "Memory") => "Memory snapshot captured with the queued prompt, when memory was enabled.",
        ("ChatDisclosureView", "ContentPanel") => "Flow panel for controls serialized into this disclosure section by the Designer.",
        ("ChatDisclosureView", "Title") => "Localized caption displayed by the disclosure toggle.",
        ("ChatDisclosureView", "Expanded") => "True when the disclosure body is visible.",
        ("ChatMessageView", "CopyButton") => "Button that copies this message's displayed content.",
        ("ChatMessageView", "ForkButton") => "Button that starts a new chat from this message.",
        ("ChatWindow", "PendingMessages") => "The current session's queued messages, or null when there is no active session.",
        _ => null
    };
