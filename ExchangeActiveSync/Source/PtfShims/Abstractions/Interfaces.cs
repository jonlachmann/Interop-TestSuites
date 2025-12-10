namespace Microsoft.Protocols.TestTools
{
    using System;
    using System.Collections.Generic;

    public enum LogEntryKind
    {
        Debug,
        Comment,
        Warning,
        Information,
        CheckFailed,
        Error,
        TestError
    }

    public interface ILog
    {
        void Add(LogEntryKind kind, string message, params object[] args);
    }

    public interface IAssert
    {
        void Fail(string message, params object[] parameters);
        void IsTrue(bool condition, string message, params object[] parameters);
        void IsFalse(bool condition, string message, params object[] parameters);
        void IsNull(object? value, string message, params object[] parameters);
        void IsNotNull(object? value, string message, params object[] parameters);
        void AreEqual<T>(T expected, T actual, string message, params object[] parameters);
        void AreNotEqual<T>(T notExpected, T actual, string message, params object[] parameters);
        void Inconclusive(string message, params object[] parameters);
    }

    public interface IAssume : IAssert
    {
    }

    public interface ITestSite
    {
        IDictionary<string, string> Properties { get; }

        IDictionary<string, object> TestProperties { get; }

        string? DefaultProtocolDocShortName { get; set; }

        ILog Log { get; }

        IAssert Assert { get; }

        IAssume Assume { get; }

        T GetAdapter<T>() where T : class;

        void CaptureRequirement(int requirementId, string description);
        void CaptureRequirement(int requirementId, string description, params object[] args);
        void CaptureRequirement(string protocol, int requirementId, string description);
        void CaptureRequirement(string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfIsTrue(bool condition, int requirementId, string description);
        void CaptureRequirementIfIsTrue(bool condition, int requirementId, string description, params object[] args);
        void CaptureRequirementIfIsTrue(bool condition, string protocol, int requirementId, string description);
        void CaptureRequirementIfIsTrue(bool condition, string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfIsNotNull(object? target, int requirementId, string description);
        void CaptureRequirementIfIsNotNull(object? target, int requirementId, string description, params object[] args);
        void CaptureRequirementIfIsNotNull(object? target, string protocol, int requirementId, string description);
        void CaptureRequirementIfIsNotNull(object? target, string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfIsNull(object? target, int requirementId, string description);
        void CaptureRequirementIfIsNull(object? target, int requirementId, string description, params object[] args);
        void CaptureRequirementIfIsNull(object? target, string protocol, int requirementId, string description);
        void CaptureRequirementIfIsNull(object? target, string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfIsFalse(bool condition, int requirementId, string description);
        void CaptureRequirementIfIsFalse(bool condition, int requirementId, string description, params object[] args);
        void CaptureRequirementIfIsFalse(bool condition, string protocol, int requirementId, string description);
        void CaptureRequirementIfIsFalse(bool condition, string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfAreEqual<T>(T expected, T actual, int requirementId, string description);
        void CaptureRequirementIfAreEqual<T>(T expected, T actual, int requirementId, string description, params object[] args);
        void CaptureRequirementIfAreEqual<T>(T expected, T actual, string protocol, int requirementId, string description);
        void CaptureRequirementIfAreEqual<T>(T expected, T actual, string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, int requirementId, string description);
        void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, int requirementId, string description, params object[] args);
        void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, string protocol, int requirementId, string description);
        void CaptureRequirementIfAreNotEqual<T>(T expected, T actual, string protocol, int requirementId, string description, params object[] args);

        void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, int requirementId, string description);
        void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, int requirementId, string description, params object[] args);
        void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, string protocol, int requirementId, string description);
        void CaptureRequirementIfIsInstanceOfType(object? value, Type expectedType, string protocol, int requirementId, string description, params object[] args);
    }

    public interface IAdapter
    {
        void Initialize(ITestSite site);
    }

    /// <summary>
    /// Optional helper attribute used by legacy PTF SUT control adapters to describe methods.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public sealed class MethodHelpAttribute : Attribute
    {
        public MethodHelpAttribute(string helpText)
        {
            this.HelpText = helpText;
        }

        public string HelpText { get; }
    }
}
