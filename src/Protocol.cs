namespace CodexVBE
{
    public sealed class Request
    {
        public string Command { get; set; }
        public string Project { get; set; }
        public string Module { get; set; }
        public int StartLine { get; set; }
        public int Count { get; set; }
        public string ExpectedSha256 { get; set; }
        public string Text { get; set; }
        public string Query { get; set; }
        public string Action { get; set; }
        public int ControlId { get; set; }
        public string ControlCaption { get; set; }
        public int ExpectedMode { get; set; }
        public string Form { get; set; }
        public string Control { get; set; }
        public string ControlType { get; set; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string Caption { get; set; }
        public string ExpectedFormVersion { get; set; }
        public string Property { get; set; }
        public object Value { get; set; }
        public string Path { get; set; }
        public string Guid { get; set; }
        public int Major { get; set; }
        public int Minor { get; set; }
        public int Offset { get; set; }
        public int Limit { get; set; }
        public int TypeIndex { get; set; }
        public string TypeIdentity { get; set; }
        public string ExpectedReferencesVersion { get; set; }
        public string ParentPath { get; set; }
        public string ExpectedTreeVersion { get; set; }
        public string ControlPath { get; set; }
        public string ExpectedProjectVersion { get; set; }
        public string ExpectedComponentVersion { get; set; }
        public string NewName { get; set; }
        public string Procedure { get; set; }
        public string EventName { get; set; }
        public string ObjectName { get; set; }
        public int ProcKind { get; set; }
        public int? InsertIndex { get; set; }
        public bool WholeWord { get; set; }
        public bool MatchCase { get; set; }
        public bool PatternSearch { get; set; }
        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool FontBold { get; set; }
    }

    public sealed class Response
    {
        public bool Ok { get; set; }
        public string Error { get; set; }
        public object Data { get; set; }

        public static Response Success(object data) { return new Response { Ok = true, Data = data }; }
        public static Response Failure(string error) { return new Response { Ok = false, Error = error }; }
    }
}
