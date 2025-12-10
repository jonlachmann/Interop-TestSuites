namespace Microsoft.Protocols.TestTools.Logging
{
    using System;

    // Placeholder for legacy PTF log sink type referenced by ptfconfig files.
    public class BeaconLogSink
    {
        public BeaconLogSink()
        {
        }

        public void Write(string message)
        {
            Console.WriteLine($"[Beacon] {message}");
        }
    }
}
