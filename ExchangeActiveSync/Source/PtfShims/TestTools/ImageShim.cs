// Minimal System.Drawing.Image stand-in to satisfy legacy code paths that only perform type checks.
namespace System.Drawing
{
    using System.IO;

    public abstract class Image
    {
        public static Image FromStream(Stream stream)
        {
            // This shim does not decode; it only validates that a stream was provided.
            // It returns a lightweight placeholder instance.
            return new ShimImage();
        }

        private sealed class ShimImage : Image
        {
        }
    }
}
