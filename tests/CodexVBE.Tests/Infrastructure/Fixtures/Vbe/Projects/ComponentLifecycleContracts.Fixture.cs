namespace CodexVBE.Tests.Unit
{
    /// <summary>Mutable native reference observations for lifecycle version checks.</summary>
    public sealed class LifecycleReference
    {
        public string GUID { get; set; }
        public int Major { get; set; }
        public int Minor { get; set; }
        public bool IsBroken { get; set; }
        public bool BuiltIn { get; set; }
    }
}