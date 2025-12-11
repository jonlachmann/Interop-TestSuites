namespace Microsoft.Protocols.TestTools;

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
        if (!Equals(expected, actual))
        {
            throw new TestAssertFailedException(Format(message, parameters, $"Expected: {expected}, Actual: {actual}"));
        }
    }

    public void AreNotEqual<T>(T notExpected, T actual, string message, params object[] parameters)
    {
        if (Equals(notExpected, actual))
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
        var formatted = args is { Length: > 0 } ? string.Format(message, args) : message;
        Console.WriteLine($"[{kind}] {formatted}");
    }
}

internal class PropertyBag : IDictionary<string, string>
{
    private readonly Dictionary<string, string> inner = new(StringComparer.OrdinalIgnoreCase);

    public string? this[string key]
    {
        get => inner.TryGetValue(key, out var value) ? value : null;
        set
        {
            if (value == null)
            {
                inner.Remove(key);
            }
            else
            {
                inner[key] = value;
            }
        }
    }

    public ICollection<string> Keys => inner.Keys;

    public ICollection<string> Values => inner.Values;

    public int Count => inner.Count;

    public bool IsReadOnly => false;

    public void Add(string key, string value) => inner.Add(key, value);

    public void Add(KeyValuePair<string, string> item) => inner.Add(item.Key, item.Value);

    public void Clear() => inner.Clear();

    public bool Contains(KeyValuePair<string, string> item) => inner.Contains(item);

    public bool ContainsKey(string key) => inner.ContainsKey(key);

    public void CopyTo(KeyValuePair<string, string>[] array, int arrayIndex) => ((IDictionary<string, string>)inner).CopyTo(array, arrayIndex);

    public IEnumerator<KeyValuePair<string, string>> GetEnumerator() => inner.GetEnumerator();

    public bool Remove(string key) => inner.Remove(key);

    public bool Remove(KeyValuePair<string, string> item) => inner.Remove(item.Key);

    public bool TryGetValue(string key, out string value) => inner.TryGetValue(key, out value!);

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => inner.GetEnumerator();
}

public class TestSite : ITestSite
{
    private readonly ConcurrentDictionary<Type, object> adapters = new();

    public TestSite()
    {
        Properties = new PropertyBag();
        TestProperties = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        Assert = new BasicAssert();
        Assume = new BasicAssume();
        Log = new BasicLog();
        LoadPtfConfigDefaults();
        EnsureSchemaFiles();
    }

    public IDictionary<string, string> Properties { get; }

    public IDictionary<string, object> TestProperties { get; }

    public string? DefaultProtocolDocShortName { get; set; }

    public ILog Log { get; }

    public IAssert Assert { get; }

    public IAssume Assume { get; }

    public T GetAdapter<T>() where T : class
    {
        return (T)adapters.GetOrAdd(typeof(T), CreateAdapter<T>);
    }

    public void CaptureRequirement(int requirementId, string description)
    {
        Log.Add(LogEntryKind.Comment, $"CaptureRequirement: R{requirementId} - {description}");
    }

    public void CaptureRequirement(int requirementId, string description, params object[] args)
    {
        CaptureRequirement(requirementId, Format(description, args));
    }

    public void CaptureRequirement(string protocol, int requirementId, string description)
    {
        Log.Add(LogEntryKind.Comment, $"[{protocol}] R{requirementId} - {description}");
    }

    public void CaptureRequirement(string protocol, int requirementId, string description, params object[] args)
    {
        CaptureRequirement(protocol, requirementId, Format(description, args));
    }

    public void CaptureRequirementIfAreEqual<T>(T expected, T actual, int requirementId, string description)
    {
        if (Equals(expected, actual))
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfAreEqual<T>(T expected, T actual, int requirementId, string description, params object[] args)
    {
        if (Equals(expected, actual))
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfAreEqual<T>(T expected, T actual, string protocol, int requirementId, string description)
    {
        if (Equals(expected, actual))
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfAreEqual<T>(T expected, T actual, string protocol, int requirementId, string description, params object[] args)
    {
        if (Equals(expected, actual))
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, int requirementId, string description)
    {
        if (!Equals(expected, actual))
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, int requirementId, string description, params object[] args)
    {
        if (!Equals(expected, actual))
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, string protocol, int requirementId, string description)
    {
        if (!Equals(expected, actual))
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, string protocol, int requirementId, string description, params object[] args)
    {
        if (!Equals(expected, actual))
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsNotNull(object? target, int requirementId, string description)
    {
        if (target is not null)
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfIsNotNull(object? target, int requirementId, string description, params object[] args)
    {
        if (target is not null)
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsNotNull(object? target, string protocol, int requirementId, string description)
    {
        if (target is not null)
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfIsNotNull(object? target, string protocol, int requirementId, string description, params object[] args)
    {
        if (target is not null)
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsNull(object? target, int requirementId, string description)
    {
        if (target is null)
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfIsNull(object? target, int requirementId, string description, params object[] args)
    {
        if (target is null)
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsNull(object? target, string protocol, int requirementId, string description)
    {
        if (target is null)
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfIsNull(object? target, string protocol, int requirementId, string description, params object[] args)
    {
        if (target is null)
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsTrue(bool condition, int requirementId, string description)
    {
        if (condition)
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfIsTrue(bool condition, int requirementId, string description, params object[] args)
    {
        if (condition)
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsTrue(bool condition, string protocol, int requirementId, string description)
    {
        if (condition)
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfIsTrue(bool condition, string protocol, int requirementId, string description, params object[] args)
    {
        if (condition)
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsFalse(bool condition, int requirementId, string description)
    {
        if (!condition)
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfIsFalse(bool condition, int requirementId, string description, params object[] args)
    {
        if (!condition)
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsFalse(bool condition, string protocol, int requirementId, string description)
    {
        if (!condition)
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfIsFalse(bool condition, string protocol, int requirementId, string description, params object[] args)
    {
        if (!condition)
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, int requirementId, string description)
    {
        if (value != null && expectedType.IsInstanceOfType(value))
        {
            CaptureRequirement(requirementId, description);
        }
    }

    public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, int requirementId, string description, params object[] args)
    {
        if (value != null && expectedType.IsInstanceOfType(value))
        {
            CaptureRequirement(requirementId, Format(description, args));
        }
    }

    public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, string protocol, int requirementId, string description)
    {
        if (value != null && expectedType.IsInstanceOfType(value))
        {
            CaptureRequirement(protocol, requirementId, description);
        }
    }

    public void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, string protocol, int requirementId, string description, params object[] args)
    {
        if (value != null && expectedType.IsInstanceOfType(value))
        {
            CaptureRequirement(protocol, requirementId, Format(description, args));
        }
    }

    private object CreateAdapter<T>(Type _)
    {
        var targetInterface = typeof(T);
        var assembly = targetInterface.Assembly;
        var implementation = assembly
            .GetTypes()
            .FirstOrDefault(t => targetInterface.IsAssignableFrom(t) && t.IsClass && !t.IsAbstract);

        if (implementation == null)
        {
            implementation = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a =>
                {
                    try
                    {
                        return a.GetTypes();
                    }
                    catch (ReflectionTypeLoadException ex)
                    {
                        return ex.Types.Where(t => t != null)!;
                    }
                })
                .FirstOrDefault(t => t != null && targetInterface.IsAssignableFrom(t) && t.IsClass && !t.IsAbstract);
        }

        if (implementation == null)
        {
            if (targetInterface.IsInterface && typeof(IAdapter).IsAssignableFrom(targetInterface))
            {
                return CreateNoOpAdapter(targetInterface);
            }

            throw new InvalidOperationException($"No adapter implementation found for {targetInterface.FullName}.");
        }

        var instance = Activator.CreateInstance(implementation)
                       ?? throw new InvalidOperationException($"Could not create adapter {implementation.FullName}.");

        if (instance is IAdapter adapter)
        {
            adapter.Initialize(this);
        }

        return instance;
    }

    private object CreateNoOpAdapter(Type adapterInterface)
    {
        var proxy = DispatchProxy.Create(adapterInterface, typeof(NoOpAdapterProxy))
                    ?? throw new InvalidOperationException($"Could not create proxy for {adapterInterface.FullName}.");

        if (proxy is NoOpAdapterProxy shim)
        {
            shim.Site = this;
        }

        return proxy;
    }

    private static string Format(string description, object[] args)
    {
        return args is { Length: > 0 } ? string.Format(description, args) : description;
    }

    private void EnsureSchemaFiles()
    {
        try
        {
            var schemaRoot = GetPtfConfigProbeRoots()
                .Select(r => Path.Combine(r, "Common", "ActiveSyncClient", "SchemaValidation"))
                .FirstOrDefault(Directory.Exists);

            if (schemaRoot == null)
            {
                return;
            }

            var targetDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
            foreach (var src in Directory.EnumerateFiles(schemaRoot, "*.xsd", SearchOption.AllDirectories))
            {
                var dest = Path.Combine(targetDir, Path.GetFileName(src));
                if (!File.Exists(dest))
                {
                    File.Copy(src, dest, overwrite: false);
                }
            }
        }
        catch
        {
            // Best-effort; schema validation will fail later with a clearer error if copies are unavailable.
        }
    }

    private void LoadPtfConfigDefaults()
    {
        try
        {
            var roots = GetPtfConfigProbeRoots().ToList();
            var suitePrefix = GuessSuitePrefix();

            var suiteConfigs = CollectSuiteConfigs(roots, suitePrefix).ToList();
            foreach (var config in suiteConfigs)
            {
                LoadPtfConfigFile(config, overwriteExisting: false);
            }

            ResolveCommonConfigPath(roots);

            if (Properties["CommonConfigurationFileName"] == null)
            {
                var commonConfig = FindConfigByName("ExchangeCommonConfiguration.deployment.ptfconfig", roots);
                Properties["CommonConfigurationFileName"] = commonConfig ?? "ExchangeCommonConfiguration.deployment.ptfconfig";
            }

            if (DefaultProtocolDocShortName == null)
            {
                var inferred = InferProtocolShortName(suitePrefix, suiteConfigs);
                if (!string.IsNullOrEmpty(inferred))
                {
                    DefaultProtocolDocShortName = inferred;
                }
            }
        }
        catch
        {
            // Best-effort preload. Failures will surface later via explicit config validation.
        }
    }

    private static IEnumerable<string> GetPtfConfigProbeRoots()
    {
        List<string> roots = new();

        var baseDir = AppContext.BaseDirectory;
        if (!string.IsNullOrEmpty(baseDir) && Directory.Exists(baseDir))
        {
            roots.Add(baseDir);
        }

        var currentDir = Directory.GetCurrentDirectory();
        if (!string.IsNullOrEmpty(currentDir) && Directory.Exists(currentDir) && !roots.Any(r => string.Equals(r, currentDir, StringComparison.OrdinalIgnoreCase)))
        {
            roots.Add(currentDir);
        }

        var solutionRoot = LocateSolutionRoot();
        if (solutionRoot != null && !roots.Any(r => string.Equals(r, solutionRoot, StringComparison.OrdinalIgnoreCase)))
        {
            roots.Add(solutionRoot);
        }

        return roots;
    }

    private static string? GuessSuitePrefix()
    {
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
        {
            var name = asm.GetName().Name;
            if (string.IsNullOrEmpty(name))
            {
                continue;
            }

            if (name.EndsWith("_TestSuite", StringComparison.OrdinalIgnoreCase))
            {
                return name[..^"_TestSuite".Length];
            }
        }

        return null;
    }

    private static IEnumerable<string> CollectSuiteConfigs(IEnumerable<string> roots, string? suitePrefix)
    {
        HashSet<string> results = new(StringComparer.OrdinalIgnoreCase);

        foreach (var root in roots.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var candidates = Directory.EnumerateFiles(root, "*_TestSuite*.ptfconfig", SearchOption.AllDirectories)
                .Where(f => !Path.GetFileName(f).Contains("_SHOULDMAY", StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(suitePrefix))
            {
                var prefixWithUnderscore = suitePrefix + "_";
                candidates = candidates.Where(f => Path.GetFileName(f).StartsWith(prefixWithUnderscore, StringComparison.OrdinalIgnoreCase));
            }

            foreach (var candidate in candidates.OrderBy(f => f.Length))
            {
                results.Add(candidate);
            }
        }

        if (results.Count == 0)
        {
            foreach (var root in roots.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                foreach (var candidate in Directory.EnumerateFiles(root, "*.ptfconfig", SearchOption.AllDirectories)
                             .Where(f => !Path.GetFileName(f).Contains("_SHOULDMAY", StringComparison.OrdinalIgnoreCase)))
                {
                    results.Add(candidate);
                }
            }
        }

        // Prioritize configs that sit closer to the base directory to mirror MSTest deployment behavior.
        return results.OrderBy(f => f.StartsWith(AppContext.BaseDirectory ?? string.Empty, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ThenBy(f => f, StringComparer.OrdinalIgnoreCase);
    }

    private static string? FindConfigByName(string fileName, IEnumerable<string> roots)
    {
        foreach (var root in roots)
        {
            var direct = Path.Combine(root, fileName);
            if (File.Exists(direct))
            {
                return direct;
            }

            var match = Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories).FirstOrDefault();
            if (match != null)
            {
                return match;
            }
        }

        return null;
    }

    private void ResolveCommonConfigPath(IEnumerable<string> roots)
    {
        var commonConfig = Properties["CommonConfigurationFileName"];
        if (string.IsNullOrWhiteSpace(commonConfig))
        {
            return;
        }

        if (Path.IsPathRooted(commonConfig) && File.Exists(commonConfig))
        {
            return;
        }

        var resolved = FindConfigByName(Path.GetFileName(commonConfig), roots);
        if (!string.IsNullOrEmpty(resolved))
        {
            Properties["CommonConfigurationFileName"] = resolved;
        }
    }

    private static string? InferProtocolShortName(string? suitePrefix, IEnumerable<string> loadedConfigs)
    {
        if (!string.IsNullOrEmpty(suitePrefix))
        {
            return suitePrefix;
        }

        var suiteConfig = loadedConfigs.FirstOrDefault(f => Path.GetFileName(f).Contains("_TestSuite", StringComparison.OrdinalIgnoreCase));
        if (suiteConfig != null)
        {
            var name = Path.GetFileNameWithoutExtension(suiteConfig);
            var underscoreIndex = name.IndexOf('_');
            if (underscoreIndex > 0)
            {
                return name[..underscoreIndex];
            }
        }

        return null;
    }

    private static string? LocateSolutionRoot()
    {
        var current = AppContext.BaseDirectory;
        while (!string.IsNullOrEmpty(current))
        {
            var candidate = Path.Combine(current, "ExchangeServerEASProtocolTestSuites.sln");
            if (File.Exists(candidate))
            {
                return current;
            }

            current = Directory.GetParent(current)?.FullName;
        }

        return null;
    }

    private void LoadPtfConfigFile(string path, bool overwriteExisting)
    {
        try
        {
            XmlDocument doc = new();
            doc.Load(path);
            XmlNamespaceManager nsmgr = new(doc.NameTable);
            nsmgr.AddNamespace("tc", "http://schemas.microsoft.com/windows/ProtocolsTest/2007/07/TestConfig");

            var properties = doc.DocumentElement?.SelectNodes("//tc:Property", nsmgr);
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

                var name = property.Attributes["name"]!.Value;
                var value = property.Attributes["value"]!.Value;
                if (overwriteExisting || Properties[name] == null)
                {
                    Properties[name] = value;
                }
            }
        }
        catch
        {
            // Ignore malformed files during preload; explicit merge will validate.
        }
    }
}

internal class NoOpAdapterProxy : DispatchProxy
{
    public TestSite? Site { get; set; }

    protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
    {
        if (targetMethod == null)
        {
            return null;
        }

        // Initialize is part of IAdapter; ignore to allow test execution to continue.
        if (targetMethod.Name.Equals("Initialize", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        Site?.Log.Add(LogEntryKind.Warning, $"No-op adapter invoked: {targetMethod.DeclaringType?.FullName}.{targetMethod.Name}");
        return GetDefaultValue(targetMethod.ReturnType);
    }

    private static object? GetDefaultValue(Type? returnType)
    {
        if (returnType == null || returnType == typeof(void))
        {
            return null;
        }

        if (returnType.IsValueType)
        {
            return Activator.CreateInstance(returnType);
        }

        return null;
    }
}

public abstract class ManagedAdapterBase : IAdapter
{
    protected ITestSite Site { get; private set; } = null!;

    public virtual void Initialize(ITestSite site)
    {
        Site = site;
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
        TestInitialize();
    }

    [TestCleanup]
    public void BaseTestCleanup()
    {
        TestCleanup();
    }

    public T GetAdapter<T>() where T : class
    {
        return BaseTestSite.GetAdapter<T>();
    }

    private static void TryPopulateTestName(object testContext)
    {
        try
        {
            var ctxType = testContext.GetType();
            var nameProperty = ctxType.GetProperty("TestName") ?? ctxType.GetProperty("FullyQualifiedTestClassName");
            if (nameProperty != null)
            {
                var name = nameProperty.GetValue(testContext);
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
