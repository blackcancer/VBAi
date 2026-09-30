using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace VBAi
{
    /// <summary>Complète l’adaptateur MSForms pour les écritures de propriétés scalaires du Designer.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Identifie les objets COM natifs sans remplacer le dispatch des propriétés.</summary>
        internal static Func<object, bool> NativeDesignerObject = Marshal.IsComObject;
                /// <summary>Utilise le dispatch typé pour les propriétés MSForms usuelles ; les doubles .NET conservent leurs descripteurs.</summary>
                /// <param name="target">Contrôle MSForms ou objet .NET à mettre à jour.</param>
                /// <param name="descriptor">Descripteur de la propriété à écrire.</param>
                /// <param name="value">Valeur convertie vers le type attendu par la propriété native.</param>
        private static void SetDesignerScalar(object target, PropertyDescriptor descriptor, object value)
        {
            if (!NativeDesignerObject(target)) { descriptor.SetValue(target, value); return; }
            dynamic native = target;
            switch (descriptor.Name.ToLowerInvariant())
            {
                case "left": native.Left = Convert.ToSingle(value); return;
                case "top": native.Top = Convert.ToSingle(value); return;
                case "width": native.Width = Convert.ToSingle(value); return;
                case "height": native.Height = Convert.ToSingle(value); return;
                case "scrollwidth": native.ScrollWidth = Convert.ToSingle(value); return;
                case "scrollheight": native.ScrollHeight = Convert.ToSingle(value); return;
                case "caption": native.Caption = Convert.ToString(value); return;
                case "name": native.Name = Convert.ToString(value); return;
                case "text": native.Text = Convert.ToString(value); return;
                case "tag": native.Tag = Convert.ToString(value); return;
                case "controltiptext": native.ControlTipText = Convert.ToString(value); return;
                case "enabled": native.Enabled = Convert.ToBoolean(value); return;
                case "visible": native.Visible = Convert.ToBoolean(value); return;
                case "locked": native.Locked = Convert.ToBoolean(value); return;
                case "tabstop": native.TabStop = Convert.ToBoolean(value); return;
                case "wordwrap": native.WordWrap = Convert.ToBoolean(value); return;
                case "autosize": native.AutoSize = Convert.ToBoolean(value); return;
                case "bold": native.Bold = Convert.ToBoolean(value); return;
                case "italic": native.Italic = Convert.ToBoolean(value); return;
                case "underline": native.Underline = Convert.ToBoolean(value); return;
                case "strikethrough": native.Strikethrough = Convert.ToBoolean(value); return;
                default: VbeScalarProperty.SetNative(target, descriptor.Name, value); return;
            }
        }
    }
}
