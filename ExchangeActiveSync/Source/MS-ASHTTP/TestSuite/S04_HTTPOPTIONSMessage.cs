namespace Microsoft.Protocols.TestSuites.MS_ASHTTP;

using System.Net;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// This scenario is designed to test HTTP OPTIONS.
/// </summary>
[TestClass]
public class S04_HTTPOPTIONSMessage : TestSuiteBase
{
    #region Class initialize and clean up
    /// <summary>
    /// Initialize the class.
    /// </summary>
    /// <param name="testContext">VSTS test context.</param>
    [ClassInitialize]
    public static void ClassInitialize(TestContext testContext)
    {
        Initialize(testContext);
    }

    /// <summary>
    /// Clear the class.
    /// </summary>
    [ClassCleanup]
    public static void ClassCleanup()
    {
        Cleanup();
    }
    #endregion

    #region Test cases
    /// <summary>
    /// This test case is intended to validate the HTTP OPTIONS command.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S04_TC01_HTTPOPTIONS()
    {
        #region Call HTTP OPTIONS command.
        var optionsResponse = HTTPAdapter.HTTPOPTIONS();
        Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, optionsResponse.StatusCode, "The StatusCode of HTTP OPTIONS command response should be OK, actual {0}.", optionsResponse.StatusCode);
        #endregion
    }
    #endregion
}