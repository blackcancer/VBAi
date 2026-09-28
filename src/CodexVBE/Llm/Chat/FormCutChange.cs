using System.Web.Script.Serialization;
namespace CodexVBE
{
    internal sealed class FormCutChange
    {
        public string Project { get; set; }
        public string Form { get; set; }
        public string ParentPath { get; set; }
        public int ControlCount { get; set; }
        public bool Restored { get; set; }
        public bool Attempted { get; set; }
        [ScriptIgnore] public string RecoveryId { get; set; }
        [ScriptIgnore] public LlmVbeTools Owner { get; set; }
    }
}
