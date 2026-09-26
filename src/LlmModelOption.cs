namespace CodexVBE
{
    internal sealed class LlmModelOption
    {
        public string Id { get; private set; }
        public string Label { get; private set; }
        public bool IsDefault { get; private set; }

        public LlmModelOption(string id, string label, bool isDefault = false)
        {
            Id = id;
            Label = string.IsNullOrWhiteSpace(label) ? id : label;
            IsDefault = isDefault;
        }

        public override string ToString() { return Label == Id ? Id : Label + " (" + Id + ")"; }
    }
}
