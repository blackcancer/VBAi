namespace CodexVBE
{
    internal sealed class LlmModelOption
    {
        public string Id { get; private set; }
        public string Label { get; private set; }
        public bool IsDefault { get; private set; }
        public string DefaultEffort { get; private set; }
        public LlmEffortOption[] Efforts { get; private set; }

        public LlmModelOption(string id, string label, bool isDefault = false,
            string defaultEffort = null, LlmEffortOption[] efforts = null)
        {
            Id = id;
            Label = string.IsNullOrWhiteSpace(label) ? id : label;
            IsDefault = isDefault;
            DefaultEffort = defaultEffort;
            Efforts = efforts ?? new LlmEffortOption[0];
        }

        public override string ToString() { return Label == Id ? Id : Label + " (" + Id + ")"; }
    }

    internal sealed class LlmEffortOption
    {
        public string Id { get; private set; }
        public string Description { get; private set; }
        public LlmEffortOption(string id, string description) { Id = id; Description = description; }
        public override string ToString() { return Id; }
    }
}
