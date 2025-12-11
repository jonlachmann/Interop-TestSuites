namespace Microsoft.Protocols.TestSuites.MS_ASHTTP;

using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

/// <summary>
/// This scenario is designed to test optional headers of HTTP POST command.
/// </summary>
[TestClass]
public class S03_HTTPPOSTOptionalHeader : TestSuiteBase
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
    /// This test case is intended to validate the MS-ASAcceptMultiPart optional header in HTTP POST request.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S03_TC01_SetASAcceptMultiPartRequestHeader()
    {
        #region Call SendMail command to send email to User2.
        // Call ConfigureRequestPrefixFields to set the QueryValueType to PlainText.
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString());
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        // Call FolderSync command to synchronize the collection hierarchy.
        CallFolderSyncCommand();

        var sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
        var userOneMailboxAddress = Common.GetMailAddress(UserOneInformation.UserName, UserOneInformation.UserDomain);
        var userTwoMailboxAddress = Common.GetMailAddress(UserTwoInformation.UserName, UserTwoInformation.UserDomain);

        // Call SendMail command.
        CallSendMailCommand(userOneMailboxAddress, userTwoMailboxAddress, sendMailSubject, null);
        #endregion

        #region Get the received email.
        // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
        SwitchUser(UserTwoInformation, true);
        AddCreatedItemToCollection("User2", UserTwoInformation.InboxCollectionId, sendMailSubject);

        // Call Sync command to get the received email.
        var itemServerId = LoopToSyncItem(UserTwoInformation.InboxCollectionId, sendMailSubject, true);
        #endregion

        #region Call ItemOperation command with setting MS-ASAcceptMultiPart header to "T".
        // Call ConfigureRequestPrefixFields to set MS-ASAcceptMultiPart header to "T".
        requestPrefix.Add(HTTPPOSTRequestPrefixField.AcceptMultiPart, "T");
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        // Call ItemOperation command to fetch the received email.
        var itemOperationResponse = CallItemOperationsCommand(UserTwoInformation.InboxCollectionId, itemServerId, false);
        Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R154");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R154
        // The content is in multipart, so this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "application/vnd.ms-sync.multipart",
            itemOperationResponse.Headers["Content-Type"],
            154,
            @"[In MS-ASAcceptMultiPart] If this [MS-ASAcceptMultiPart] header is present and the value is 'T', the client is requesting that the server return content in multipart format.");
        #endregion

        #region Call ItemOperation command with setting MS-ASAcceptMultiPart header to "F".
        // Call ConfigureRequestPrefixFields to change the MS-ASAcceptMultiPart header to "F".
        requestPrefix[HTTPPOSTRequestPrefixField.AcceptMultiPart] = "F";
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        // Call ItemOperation command to fetch the received email.
        itemOperationResponse = CallItemOperationsCommand(UserTwoInformation.InboxCollectionId, itemServerId, false);
        Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R440");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R440
        // The content is not in multipart, so this requirement can be captured.
        Site.CaptureRequirementIfAreNotEqual<string>(
            "application/vnd.ms-sync.multipart",
            itemOperationResponse.Headers["Content-Type"],
            440,
            @"[In MS-ASAcceptMultiPart] If the [MS-ASAcceptMultiPart] header [is not present, or] is present and set to 'F', the client is requesting that the server return content in inline format.");
        #endregion

        #region Call ItemOperation command with setting MS-ASAcceptMultiPart header to null.
        // Call ConfigureRequestPrefixFields to change the MS-ASAcceptMultiPart header to null.
        requestPrefix[HTTPPOSTRequestPrefixField.AcceptMultiPart] = null;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        // Call ItemOperation command to fetch the received email.
        itemOperationResponse = CallItemOperationsCommand(UserTwoInformation.InboxCollectionId, itemServerId, false);
        Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R155");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R155
        // The content is not in multipart, so this requirement can be captured.
        Site.CaptureRequirementIfAreNotEqual<string>(
            "application/vnd.ms-sync.multipart",
            itemOperationResponse.Headers["Content-Type"],
            155,
            @"[In MS-ASAcceptMultiPart] If the [MS-ASAcceptMultiPart] header is not present [, or is present and set to 'F'], the client is requesting that the server return content in inline format.");
        #endregion

        #region Reset the query value type and credential.
        requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", Site);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        SwitchUser(UserOneInformation, false);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the User-Agent optional header in HTTP POST request.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S03_TC02_SetUserAgentRequestHeader()
    {
        #region Call ConfigureRequestPrefixFields to add the User-Agent header.
        var folderSyncRequestBody = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();
        var requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>
        {
            {
                HTTPPOSTRequestPrefixField.UserAgent, "ASOM"
            }
        };

        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
        #endregion

        #region Call FolderSync command.
        var folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);

        // Check the command is executed successfully.
        CheckResponseStatus(folderSyncResponse.ResponseDataXML);
        #endregion

        #region Reset the User-Agent header.
        requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = null;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the X-MS-PolicyKey optional header and Policy key optional field in HTTP POST request.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S03_TC03_SetPolicyKeyRequestHeader()
    {
        #region Change the query value type to PlainText.
        // Call ConfigureRequestPrefixFields to set the QueryValueType to PlainText.
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString());
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Call Provision command without setting X-MS-PolicyKey header.
        var provisionResponse = CallProvisionCommand(string.Empty);

        // Get the policy key from the response of Provision command.
        var policyKey = TestSuiteHelper.GetPolicyKeyFromSendString(provisionResponse);
        #endregion

        #region Call Provision command with setting X-MS-PolicyKey header of the PlainText encoded query value type.
        // Set the X-MS-PolicyKey header.
        requestPrefix.Add(HTTPPOSTRequestPrefixField.PolicyKey, policyKey);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        CallProvisionCommand(policyKey);

        // Reset the X-MS-PolicyKey header.
        requestPrefix[HTTPPOSTRequestPrefixField.PolicyKey] = string.Empty;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Change the query value type to Base64.
        // Call ConfigureRequestPrefixFields to set the QueryValueType to Base64.
        requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = QueryValueType.Base64.ToString();
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Call Provision command without setting Policy key field.
        provisionResponse = CallProvisionCommand(string.Empty);

        // Get the policy key from the response of Provision command.
        policyKey = TestSuiteHelper.GetPolicyKeyFromSendString(provisionResponse);
        #endregion

        #region Call Provision command with setting Policy key field of the base64 encoded query value type.
        // Set the Policy key field.
        requestPrefix[HTTPPOSTRequestPrefixField.PolicyKey] = policyKey;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        CallProvisionCommand(policyKey);

        // Reset the Policy key field.
        requestPrefix[HTTPPOSTRequestPrefixField.PolicyKey] = string.Empty;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Reset the query value type.
        requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", Site);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate server can use different values for the number of User-Agent header changes or the time
    /// period, or the time period that server blocks client from changing its User-Agent header value.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S03_TC04_LimitChangesToUserAgentHeader()
    {
        Site.Assume.IsTrue(Common.IsRequirementEnabled(456, Site), "Exchange server 2013 and above support using different values for the number of User-Agent changes or the time period.");
        Site.Assume.IsTrue(Common.IsRequirementEnabled(457, Site), "Exchange server 2013 and above support blocking clients for a different amount of time.");

        #region Call FolderSync command for the first time with User-Agent header.
        // Wait for 1 minute
        Thread.Sleep(new TimeSpan(0, 1, 0));

        var startTime = DateTime.Now;
        var folderSyncRequestBody = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();
        var requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>
        {
            {
                HTTPPOSTRequestPrefixField.UserAgent, Common.GenerateResourceName(Site, "ASOM", 1)
            }
        };

        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
        var folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);

        // Check the command is executed successfully.
        CheckResponseStatus(folderSyncResponse.ResponseDataXML);

        #endregion

        #region Call FolderSync command for the second time with updated User-Agent header.

        requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = Common.GenerateResourceName(Site, "ASOM", 2);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
        folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);

        // Check the command is executed successfully.
        CheckResponseStatus(folderSyncResponse.ResponseDataXML);

        #endregion

        #region Call FolderSync command for third time with updated User-Agent header.

        requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = Common.GenerateResourceName(Site, "ASOM", 3);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

        try
        {
            folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);
            Site.Assert.Fail("HTTP error 503 should be returned if server blocks a client from changing its User-Agent header value.");
        }
        catch (System.Net.WebException exception)
        {
            var statusCode = ((System.Net.HttpWebResponse)exception.Response).StatusCode.GetHashCode();
                
            if (Common.IsRequirementEnabled(456, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R456");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R456
                // Server configures the number of changes and the time period, and expected HTTP error is returned, this requirement can be captured.
                Site.CaptureRequirementIfAreEqual<int>(
                    503,
                    statusCode,
                    456,
                    @"[In Appendix A: Product Behavior] Implementation can be configured to use different values for the allowed number of changes and the time period. (<9> Section 3.2.5.1.1:  Exchange 2013 , Exchange 2016, and Exchange 2019 can be configured to use different values for the allowed number of changes and the time period.)");
            }
        }

        #endregion

        #region Call FolderSync command after server blocks client from changing its User-Agent header value.
        requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = Common.GenerateResourceName(Site, "ASOM", 4);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

        var isCorrectBlocked = false;
        try
        {
            folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);
        }
        catch (System.Net.WebException)
        {
            // HTTP error returns indicates server blocks client.
            isCorrectBlocked = true;
        }

        // Server sets blocking client for 1 minute, wait for 1 minute for un-blocking.
        Thread.Sleep(new TimeSpan(0, 1, 0));

        requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = Common.GenerateResourceName(Site, "ASOM", 5);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
        try
        {
            folderSyncResponse = HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequestBody);
            isCorrectBlocked = isCorrectBlocked && true;
        }
        catch (System.Net.WebException)
        {
            // HTTP error returns indicates server still blocks client.
            isCorrectBlocked = false;
        }

        if (Common.IsRequirementEnabled(457, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R457");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R457
            // FolderSync command runs successfully after the blocking time period, and it runs with exception during the time period,
            // this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isCorrectBlocked,
                457,
                @"[In Appendix A: Product Behavior] Implementation can be configured to block clients for an amount of time other than 14 hours. (<10> Section 3.2.5.1.1:  Exchange 2013, Exchange 2016, and Exchange 2019 can be configured to block clients for an amount of time other than 14 hours.)");
        }

        // Wait for 1 minute
        Thread.Sleep(new TimeSpan(0, 1, 0));

        #endregion

        #region Reset the User-Agent header.
        requestPrefixFields[HTTPPOSTRequestPrefixField.UserAgent] = null;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);
        #endregion
    }
    #endregion
}