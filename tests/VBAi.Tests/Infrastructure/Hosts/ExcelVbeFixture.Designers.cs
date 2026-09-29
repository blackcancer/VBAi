using System;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Affiche le concepteur et sa boîte à outils dans le seul classeur possédé.</summary>
        internal void FocusFormToolbox(string form)
        {
            object project = null, components = null, component = null, designer = null, window = null;
            try
            {
                project = ((dynamic)workbook).VBProject;
                components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form);
                designer = ((dynamic)component).Designer;
                window = ((dynamic)component).DesignerWindow();
                ((dynamic)window).Visible = true;
                ((dynamic)window).SetFocus();
                ((dynamic)designer).ShowToolbox = true;
            }
            finally { Release(window); Release(designer); Release(component); Release(components); Release(project); }
        }
        /// <summary>Ferme puis réouvre le seul classeur possédé et relit sa protection VBIDE.</summary>
        internal int ReopenAndReadProjectProtection(string path)
        {
            ((dynamic)workbook).Close(false); Release(workbook); workbook = null;
            workbook = ((dynamic)workbooks).Open(path, 0, false);
            object project = null;
            try { project = ((dynamic)workbook).VBProject; return Convert.ToInt32(((dynamic)project).Protection); }
            finally { Release(project); }
        }
        /// <summary>Active uniquement le volet de code d'un module du classeur possédé.</summary>
        internal void FocusCodeModule(string module)
        {
            object project = null, components = null, component = null, code = null, pane = null, window = null;
            try
            {
                project = ((dynamic)workbook).VBProject;
                components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(module);
                code = ((dynamic)component).CodeModule;
                pane = ((dynamic)code).CodePane;
                ((dynamic)pane).Show(); window = ((dynamic)pane).Window; ((dynamic)window).SetFocus();
            }
            finally { Release(window); Release(pane); Release(code); Release(component); Release(components); Release(project); }
        }
        /// <summary>Relit une dimension réelle du concepteur via le classeur possédé, indépendamment du pont.</summary>
        internal double ReadDesignerMetric(string form, string property)
        {
            object project = null, components = null, component = null, designer = null;
            try
            {
                project = ((dynamic)workbook).VBProject;
                components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form);
                designer = ((dynamic)component).Designer;
                if (property == "InsideWidth") return Convert.ToDouble(((dynamic)designer).InsideWidth);
                if (property == "InsideHeight") return Convert.ToDouble(((dynamic)designer).InsideHeight);
                if (property == "ScrollWidth") return Convert.ToDouble(((dynamic)designer).ScrollWidth);
                if (property == "ScrollHeight") return Convert.ToDouble(((dynamic)designer).ScrollHeight);
                throw new ArgumentException("Unsupported independent designer metric.");
            }
            finally { Release(designer); Release(component); Release(components); Release(project); }
        }
    }
}
