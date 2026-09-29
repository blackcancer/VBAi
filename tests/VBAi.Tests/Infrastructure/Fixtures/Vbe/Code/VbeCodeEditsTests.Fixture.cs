namespace VBAi.Tests.Unit
{
    public sealed partial class VbeCodeEditsTests
    {
        public sealed class ProcedureCatalog
        {
            public string Sha256 { get; set; }
            public ProcedureRow[] Procedures { get; set; }
        }
        public sealed class ProcedureRow
        {
            public string Name { get; set; }
            public int Kind { get; set; }
            public int BodyLine { get; set; }
            public int EndLine { get; set; }
        }
        public sealed class DesignState { public int Mode { get; set; } }
        /// <summary>Type VBIDE du composant relu avant le renommage de paramètre.</summary>
        public sealed class ComponentKind { public int Type { get; set; } }

    }
}
