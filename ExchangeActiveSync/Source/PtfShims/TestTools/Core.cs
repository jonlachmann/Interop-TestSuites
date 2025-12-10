namespace Microsoft.Protocols.TestTools
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public class TestAssertFailedException : Exception
    {
        public TestAssertFailedException(string message)
            : base(message)
        {
        }
    }

    internal class BasicAssert : IAssert
    {
        public void AreEqual<T>(T expected, T actual, string message, params object[] parameters)
        {
            if (!object.Equals(expected, actual))
            {
                throw new TestAssertFailedException(Format(message, parameters, $"Expected: {expected}, Actual: {actual}"));
            }
        }

        public void AreNotEqual<T>(T notExpected, T actual, string message, params object[] parameters)
        {
            if (object.Equals(notExpected, actual))
            {
                throw new TestAssertFailedException(Format(message, parameters, $"Did not expect: {notExpected}"));
            }
        }

        public void Inconclusive(string message, params object[] parameters)
        {
            throw new TestAssertFailedException("Inconclusive: " + Format(message, parameters, "Test marked inconclusive."));
        }

        public void Fail(string message, params object[] parameters)
        {
            throw new TestAssertFailedException(Format(message, parameters, "Assertion failed."));
        }

        public void IsFalse(bool condition, string message, params object[] parameters)
        {
            if (condition)
            {
                throw new TestAssertFailedException(Format(message, parameters, "Expected false but was true."));
            }
        }

        public void IsNotNull(object? value, string message, params object[] parameters)
        {
            if (value is null)
            {
                throw new TestAssertFailedException(Format(message, parameters, "Value was null."));
            }
        }

        public void IsNull(object? value, string message, params object[] parameters)
        {
            if (value is not null)
            {
                throw new TestAssertFailedException(Format(message, parameters, "Value was not null."));
            }
        }

        public void IsTrue(bool condition, string message, params object[] parameters)
        {
            if (!condition)
            {
                throw new TestAssertFailedException(Format(message, parameters, "Expected true but was false."));
            }
        }

        private static string Format(string message, object[] parameters, string fallback)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return fallback;
            }

            return parameters is { Length: > 0 } ? string.Format(message, parameters) : message;
        }
    }

    internal class BasicAssume : BasicAssert, IAssume
    {
    }

    internal class BasicLog : ILog
    {
        public void Add(LogEntryKind kind, string message, params object[] args)
        {
            string formatted = args is { Length: > 0 } ? string.Format(message, args) : message;
            Console.WriteLine($"[{kind}] {formatted}");
        }
    }

    public class TestSite : ITestSite
    {
        private readonly ConcurrentDictionary<Type, object> adapters = new();

        public TestSite()
        {
            this.Properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            this.TestProperties = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            this.Assert = new BasicAssert();
            this.Assume = new BasicAssume();
            this.Log = new BasicLog();
            this.LoadPtfConfigDefaults();
        }

        public IDictionary<string, string> Properties { get; }

        public IDictionary<string, object> TestProperties { get; }

        public string? DefaultProtocolDocShortName { get; set; }

        public ILog Log { get; }

        public IAssert Assert { get; }

        public IAssume Assume { get; }

        public T GetAdapter<T>() where T : class
        {
            return (T)this.adapters.GetOrAdd(typeof(T), CreateAdapter<T>);
        }

        public void CaptureRequirement(int requirementId, string description)
        {
            this.Log.Add(LogEntryKind.Comment, $"CaptureRequirement: R{requirementId} - {description}");
        }

        public void CaptureRequirement(int requirementId, string description, params object[] args)
        {
            this.CaptureRequirement(requirementId, Format(description, args));
        }

        public void CaptureRequirement(string protocol, int requirementId, string description)
        {
            this.Log.Add(LogEntryKind.Comment, $"[{protocol}] R{requirementId} - {description}");
        }

        public void CaptureRequirement(string protocol, int requirementId, string description, params object[] args)
        {
            this.CaptureRequirement(protocol, requirementId, Format(description, args));
        }

        public void CaptureRequirementIfAreEqual<T>(T expected, T actual, int requirementId, string description)
        {
            if (object.Equals(expected, actual))
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfAreEqual<T>(T expected, T actual, int requirementId, string description, params object[] args)
        {
            if (object.Equals(expected, actual))
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfAreEqual<T>(T expected, T actual, string protocol, int requirementId, string description)
        {
            if (object.Equals(expected, actual))
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfAreEqual<T>(T expected, T actual, string protocol, int requirementId, string description, params object[] args)
        {
            if (object.Equals(expected, actual))
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, int requirementId, string description)
        {
            if (!object.Equals(expected, actual))
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, int requirementId, string description, params object[] args)
        {
            if (!object.Equals(expected, actual))
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, string protocol, int requirementId, string description)
        {
            if (!object.Equals(expected, actual))
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, string protocol, int requirementId, string description, params object[] args)
        {
            if (!object.Equals(expected, actual))
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsNotNull(object? target, int requirementId, string description)
        {
            if (target is not null)
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfIsNotNull(object? target, int requirementId, string description, params object[] args)
        {
            if (target is not null)
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsNotNull(object? target, string protocol, int requirementId, string description)
        {
            if (target is not null)
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfIsNotNull(object? target, string protocol, int requirementId, string description, params object[] args)
        {
            if (target is not null)
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsNull(object? target, int requirementId, string description)
        {
            if (target is null)
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfIsNull(object? target, int requirementId, string description, params object[] args)
        {
            if (target is null)
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsNull(object? target, string protocol, int requirementId, string description)
        {
            if (target is null)
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfIsNull(object? target, string protocol, int requirementId, string description, params object[] args)
        {
            if (target is null)
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsTrue(bool condition, int requirementId, string description)
        {
            if (condition)
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfIsTrue(bool condition, int requirementId, string description, params object[] args)
        {
            if (condition)
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsTrue(bool condition, string protocol, int requirementId, string description)
        {
            if (condition)
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfIsTrue(bool condition, string protocol, int requirementId, string description, params object[] args)
        {
            if (condition)
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsFalse(bool condition, int requirementId, string description)
        {
            if (!condition)
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfIsFalse(bool condition, int requirementId, string description, params object[] args)
        {
            if (!condition)
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsFalse(bool condition, string protocol, int requirementId, string description)
        {
            if (!condition)
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfIsFalse(bool condition, string protocol, int requirementId, string description, params object[] args)
        {
            if (!condition)
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, int requirementId, string description)
        {
            if (value != null && expectedType.IsInstanceOfType(value))
            {
                this.CaptureRequirement(requirementId, description);
            }
        }

        public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, int requirementId, string description, params object[] args)
        {
            if (value != null && expectedType.IsInstanceOfType(value))
            {
                this.CaptureRequirement(requirementId, Format(description, args));
            }
        }

        public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, string protocol, int requirementId, string description)
        {
            if (value != null && expectedType.IsInstanceOfType(value))
            {
                this.CaptureRequirement(protocol, requirementId, description);
            }
        }

        public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, string protocol, int requirementId, string description, params object[] args)
        {
            if (value != null && expectedType.IsInstanceOfType(value))
            {
                this.CaptureRequirement(protocol, requirementId, Format(description, args));
            }
        }

        private object CreateAdapter<T>(Type _)
        {
            Type targetInterface = typeof(T);
            Assembly assembly = targetInterface.Assembly;
            Type? implementation = assembly
                .GetTypes()
                .FirstOrDefault(t => targetInterface.IsAssignableFrom(t) && t.IsClass && !t.IsAbstract);

            if (implementation == null)
            {
                throw new InvalidOperationException($"No adapter implementation found for {targetInterface.FullName}.");
            }

            object instance = Activator.CreateInstance(implementation)
                               ?? throw new InvalidOperationException($"Could not create adapter {implementation.FullName}.");

            if (instance is IAdapter adapter)
            {
                adapter.Initialize(this);
            }

            return instance;
        }

        private static string Format(string description, object[] args)
        {
            return args is { Length: > 0 } ? string.Format(description, args) : description;
        }

        private void LoadPtfConfigDefaults()
        {
            try
            {
                string? root = LocateSolutionRoot();
                if (root == null)
                {
                    return;
                }

                List<string> files = new();
                string commonConfig = Path.Combine(root, "Common", "ExchangeCommonConfiguration.deployment.ptfconfig");
                if (File.Exists(commonConfig))
                {
                    files.Add(commonConfig);
                }

                files.AddRange(Directory.EnumerateFiles(root, "*_TestSuite.ptfconfig", SearchOption.AllDirectories));
                files.AddRange(Directory.EnumerateFiles(root, "*_deployment.ptfconfig", SearchOption.AllDirectories));

                foreach (string file in files.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    LoadPtfConfigFile(file);
                }
            }
            catch
            {
                // Best-effort preload. Failures will surface later via explicit config validation.
            }
        }

        private static string? LocateSolutionRoot()
        {
            string? current = AppContext.BaseDirectory;
            while (!string.IsNullOrEmpty(current))
            {
                string candidate = Path.Combine(current, "ExchangeServerEASProtocolTestSuites.sln");
                if (File.Exists(candidate))
                {
                    return current;
                }

                current = Directory.GetParent(current)?.FullName;
            }

            return null;
        }

        private void LoadPtfConfigFile(string path)
        {
            try
            {
                XmlDocument doc = new();
                doc.Load(path);
                XmlNamespaceManager nsmgr = new(doc.NameTable);
                nsmgr.AddNamespace("tc", "http://schemas.microsoft.com/windows/ProtocolsTest/2007/07/TestConfig");

                XmlNodeList? properties = doc.DocumentElement?.SelectNodes("//tc:Property", nsmgr);
                if (properties == null)
                {
                    return;
                }

                foreach (XmlNode property in properties)
                {
                    if (property.Attributes?["name"] == null || property.Attributes?["value"] == null)
                    {
                        continue;
                    }

                    string name = property.Attributes["name"]!.Value;
                    string value = property.Attributes["value"]!.Value;
                    this.Properties[name] = value;
                }
            }
            catch
            {
                // Ignore malformed files during preload; explicit merge will validate.
            }
        }
    }

    public abstract class ManagedAdapterBase : IAdapter
    {
        protected ITestSite Site { get; private set; } = null!;

        public virtual void Initialize(ITestSite site)
        {
            this.Site = site;
        }
    }

    public class TestClassBase
    {
        public static ITestSite BaseTestSite { get; private set; } = new TestSite();

        protected ITestSite Site => BaseTestSite;

        public static void Initialize(object? testContext)
        {
            if (testContext != null)
            {
                TryPopulateTestName(testContext);
            }
        }

        public static void Cleanup()
        {
        }

        protected virtual void TestInitialize()
        {
        }

        protected virtual void TestCleanup()
        {
        }

        [TestInitialize]
        public void BaseTestInitialize()
        {
            this.TestInitialize();
        }

        [TestCleanup]
        public void BaseTestCleanup()
        {
            this.TestCleanup();
        }

        public T GetAdapter<T>() where T : class
        {
            return BaseTestSite.GetAdapter<T>();
        }

        private static void TryPopulateTestName(object testContext)
        {
            try
            {
                Type ctxType = testContext.GetType();
                PropertyInfo? nameProperty = ctxType.GetProperty("TestName") ?? ctxType.GetProperty("FullyQualifiedTestClassName");
                if (nameProperty != null)
                {
                    object? name = nameProperty.GetValue(testContext);
                    if (name != null)
                    {
                        BaseTestSite.TestProperties["CurrentTestCaseName"] = name.ToString() ?? string.Empty;
                    }
                }
            }
            catch
            {
                // Best-effort; ignore reflection issues
            }
        }
    }
}
