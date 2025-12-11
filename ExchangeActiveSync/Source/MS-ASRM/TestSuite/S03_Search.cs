namespace Microsoft.Protocols.TestSuites.MS_ASRM;

using System.Globalization;
using Common.DataStructures;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Request = Common.Request;

/// <summary>
/// This scenario is designed to test the Search command.
/// </summary>
[TestClass]
public class S03_Search : TestSuiteBase
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
    public static void ClassCleanUp()
    {
        Cleanup();
    }
    #endregion

    #region MSASRM_S03_TC01_Search_RightsManagedEmailMessages
    /// <summary>
    /// This test case is designed to call Search command to find a rights-managed e-mail message.
    /// </summary>
    [TestCategory("MSASRM"), TestMethod]
    public void MSASRM_S03_TC01_Search_RightsManagedEmailMessages()
    {
        CheckPreconditions();

        #region The client logs on User1's account, calls Settings command to get a templateID with all rights allowed.
        var templateID = GetTemplateID("MSASRM_AllRights_AllowedTemplate");
        #endregion

        #region The client logs on User1's account, calls SendMail command with the templateID to send a rights-managed e-mail message to User2, switches to User2, and calls FolderSync command.
        var subject = SendMailAndFolderSync(templateID, false, null);
        #endregion

        #region The client logs on User2's account, calls Search command to search the rights-managed e-mail message from server.
        var searchRequest = Common.CreateSearchRequest(subject, UserTwoInformation.InboxCollectionId);
        searchRequest.RequestData.Items[0].Options.Items = [string.Empty, string.Empty, true];
        searchRequest.RequestData.Items[0].Options.ItemsElementName =
        [
            Request.ItemsChoiceType6.RebuildResults,
            Request.ItemsChoiceType6.DeepTraversal,
            Request.ItemsChoiceType6.RightsManagementSupport
        ];

        var result = ASRMAdapter.Search(searchRequest);

        Site.Assert.AreEqual<int>(1, result.Results.Count, "There should be only 1 item fetched in ItemOperations command response.");
        var search = result.Results[0];
        Site.Assert.IsNotNull(search, "The returned item should not be null.");
        Site.Assert.IsNotNull(search.Email, "The expected rights-managed e-mail message should not be null.");
        Site.Assert.IsNull(search.Email.Attachments, "The Attachments element in expected rights-managed e-mail message should be null.");
        Site.Assert.IsNotNull(search.Email.RightsManagementLicense, "The RightsManagementLicense element in expected rights-managed e-mail message should not be null.");
        #endregion
    }
    #endregion
}