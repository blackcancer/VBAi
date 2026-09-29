namespace VBAi.Tests.Unit
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Runtime.InteropServices;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed class OlePictureLoaderTests
    {
        [TestMethod]
        public void NativeOleImageLoadAndInvalidNativeResultsMatrix()
        {
            foreach (string path in new[] { null, " ", "relative.bmp" }) Assert.ThrowsException<ArgumentException>(() => OlePictureLoader.Load(path));
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ole-picture-" + Guid.NewGuid().ToString("N") + ".bmp");
            Assert.ThrowsException<FileNotFoundException>(() => OlePictureLoader.Load(file));
            var previous = OlePictureLoader.ReadPicture;
            try
            {
                using (var bitmap = new Bitmap(3, 2)) bitmap.Save(file, ImageFormat.Bmp);
                object picture = OlePictureLoader.Load(file);
                try { Assert.IsTrue(Marshal.IsComObject(picture)); Assert.IsFalse(string.IsNullOrWhiteSpace(OlePictureLoader.Fingerprint(picture))); }
                finally { Marshal.FinalReleaseComObject(picture); }
                OlePictureLoader.ReadPicture = (object path, out object result) => result = null;
                Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Load(file));
                OlePictureLoader.ReadPicture = (object path, out object result) => result = new object();
                Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Load(file));
                Assert.IsNull(OlePictureLoader.Fingerprint(null));
                Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Fingerprint(new object()));
            }
            finally { OlePictureLoader.ReadPicture = previous; File.Delete(file); }
        }
    }
}
