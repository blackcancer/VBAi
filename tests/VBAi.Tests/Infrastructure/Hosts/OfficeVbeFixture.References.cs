using System;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Balances one acquired COM reference without invalidating aliases held by the fixture or host.</summary>
        internal Action<object> ReleaseAcquiredReference = VbaTestComReferences.Release;

        /// <summary>Retains references during uncertain delivery; otherwise releases only this getter acquisition.</summary>
        private void Release(object item)
        {
            if (item == null) return;
            if (commandContainment.Pending || commandContainment.Uncertain)
            {
                retainedDiagnosticReferences.Add(item);
                return;
            }
            ReleaseAcquiredReference(item);
        }
    }
}
