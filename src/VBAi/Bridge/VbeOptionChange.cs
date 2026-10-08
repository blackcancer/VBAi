namespace VBAi
{
    /// <summary>One staged native Options edit within a single guarded commit.</summary>
    public sealed class VbeOptionChange
    {
        /// <summary>Exact native Options tab name returned by the inspection.</summary>
        /// <value>The observed native tab label.</value>
        public string Pane { get; set; }
        /// <summary>Exact supported control label within that tab.</summary>
        /// <value>The observed native control label.</value>
        public string Property { get; set; }
        /// <summary>Native choice, Boolean state or supported numeric preference.</summary>
        /// <value>A scalar validated against the supported preference and its native choices.</value>
        public object Value { get; set; }
        /// <summary>Optional exact Code Colors category for a palette edit.</summary>
        /// <value>The observed category label, or null to retain the originally viewed category.</value>
        public string Query { get; set; }
    }
}
