using System.Collections.Generic;

namespace CodexVBE
{
    internal sealed class VbePropertyInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Kind { get; set; }
        public bool? ReadOnly { get; set; }
        public object Value { get; set; }
        public string Display { get; set; }
        public string Digest { get; set; }
        public string Error { get; set; }
        public int NumIndices { get; set; }
        public List<VbePropertyInfo> Members { get; set; }
    }
}
