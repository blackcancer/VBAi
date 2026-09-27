using System.Collections.Generic;

namespace CodexVBE
{
    internal sealed class VbePropertyInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Kind { get; set; }
        public bool? ReadOnly { get; set; }
        // COM descriptors can advertise a setter that fails at invocation
        // (observed for Label.Cancel). This is metadata, not runtime proof.
        public string SetterStatus
        {
            get { return ReadOnly == true ? "DescriptorReadOnly" :
                ReadOnly == false ? "DescriptorCandidateUnverified" : "Unknown"; }
        }
        public object Value { get; set; }
        public string Display { get; set; }
        public string Digest { get; set; }
        public string Error { get; set; }
        public int NumIndices { get; set; }
        public List<VbePropertyInfo> Members { get; set; }
    }
}
