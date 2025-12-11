namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using Common;
using Response = Common.Response;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Net;
using Request = Common.Request;

/// <summary>
/// This scenario is used to test the common status codes.
/// </summary>
[TestClass]
public class S21_CommonStatusCode : TestSuiteBase
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
    /// This test case is used to verify the server will return 166, when AccountId is invalid.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC01_CommonStatusCode_166()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 166 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 166 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call method SendMail to send e-mail messages with invalid AccountID value.
        var emailSubject = Common.GenerateResourceName(Site, "subject");

        // Send email with invalid AccountID value
        var sendMailResponse = SendPlainTextEmail("InvalidAccountID", emailSubject, User1Information.UserName, User2Information.UserName, null);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4956");
        Site.Log.Add(LogEntryKind.Debug, "When sending mail with invalid AccountID, server returns status {0}", sendMailResponse.ResponseData.Status);

        // Verify MS-ASCMD requirement: MS-ASCMD_R4956
        Site.CaptureRequirementIfAreEqual<string>(
            "166",
            sendMailResponse.ResponseData.Status,
            4956,
            @"[In Common Status Codes] [The meaning of the status value 166 is] The AccountId (section 2.2.3.3) value is not valid.<100>");

        #region Sync user2 mailbox changes
        // Switch to user2's mailbox
        SwitchUser(User2Information);
        GetMailItem(User2Information.InboxCollectionId, emailSubject);

        // Record user name, folder collectionId and item subject that is used in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, emailSubject);
        #endregion
    }

    /// <summary>
    /// This test case is used to verify server will return 173, when the picture does not exist.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC02_CommonStatusCode_173()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 173 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 173 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call method ResolveRecipients to resolve a list of supplied recipients, to retrieve their free/busy information, or retrieve their S/MIME certificates so that clients can send encrypted S/MIME e-mail messages.
        var displayName = User3Information.UserName;

        var resolveRecipientsRequest = new ResolveRecipientsRequest();
        var requestResolveRecipients = new Request.ResolveRecipients();

        var requestResolveRecipientsOption = new Request.ResolveRecipientsOptions
        {
            Picture = new Request.ResolveRecipientsOptionsPicture { MaxPictures = 3 }
        };

        requestResolveRecipients.Items = new object[] { requestResolveRecipientsOption, displayName };
        resolveRecipientsRequest.RequestData = requestResolveRecipients;

        var resolveRecipientsResponse = CMDAdapter.ResolveRecipients(resolveRecipientsRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4970");
        Site.Log.Add(LogEntryKind.Debug, "When the contact picture does not exit, server returns status {0}", resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status);

        // Verify MS-ASCMD requirement: MS-ASCMD_R4970
        Site.CaptureRequirementIfAreEqual<string>(
            "173",
            resolveRecipientsResponse.ResponseData.Response[0].Recipient[0].Picture[0].Status,
            4970,
            @"[In Common Status Codes] [The meaning of the status value 173 is] The user does not have a contact photo.<107>");
    }

    /// <summary>
    /// This test case is used to verify the server will return 165, when the required DeviceInformation element is missing in the Provision request.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC03_CommonStatusCode_165()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 165 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 165 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User calls Provision command without the DeviceInformation element

        var provisionRequest = GenerateDefaultProvisionRequest();
        provisionRequest.RequestData.DeviceInformation = null;

        var provisionResponse = CMDAdapter.Provision(provisionRequest);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4954");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4954
        Site.CaptureRequirementIfAreEqual<byte>(
            165,
            provisionResponse.ResponseData.Status,
            4954,
            @"[In Common Status Codes] [The meaning of the status value 165 is] The required DeviceInformation element (as specified in [MS-ASPROV] section 2.2.2.52) is missing in the Provision request.<99>");
    }

    /// <summary>
    /// This test case is used to verify the server will return 105, when the request contains a combination of parameters that is invalid.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC04_CommonStatusCode_105()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The DstFldId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User2 sends mail to User1 and do FolderSync in User1's mailbox.
        var subject = SendMailAndFolderSync();
        #endregion

        #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server, and get the ServerId of sent email item and the SyncKey
        var syncResponseInbox = GetMailItem(User1Information.InboxCollectionId, subject);
        var serverId = FindServerId(syncResponseInbox, "Subject", subject);
        #endregion

        #region Call method MoveItems with the email item's ServerId to move the email item from Inbox folder to recipient information cache.
        var moveItemsMove = new Request.MoveItemsMove
        {
            DstFldId = User1Information.RecipientInformationCacheCollectionId,
            SrcFldId = User1Information.InboxCollectionId,
            SrcMsgId = serverId
        };

        var moveItemsRequest = Common.CreateMoveItemsRequest(new Request.MoveItemsMove[] { moveItemsMove });
        var moveItemsResponse = CMDAdapter.MoveItems(moveItemsRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4821");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4821
        Site.CaptureRequirementIfAreEqual<int>(
            105,
            int.Parse(moveItemsResponse.ResponseData.Response[0].Status),
            4821,
            @"[In Common Status Codes] [The meaning of the status value 105 is] The request contains a combination of parameters that is invalid.");
    }

    /// <summary>
    /// This test case is used to verify the server returns 164, when the BodyPartPreference node has an unsupported Type element value.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC05_CommonStatusCode_164()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 164 is not returned when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Status value 164 is not returned when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User calls Sync command with option element

        // Set an unsupported Type element value in the BodyPartPreference node 
        var option = new Request.Options
        {
            Items = new object[]
            {
                new Request.BodyPartPreference()
                {
                    // As specified in [MS-ASAIRS] section 2.2.2.22.3, only a value of 2 (HTML) SHOULD be used in the Type element of a BodyPartPreference element.
                    // Then '3' is an unsupported Type element value.
                    Type = 3
                }
            },
            ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.BodyPartPreference }
        };

        var syncRequest = CreateEmptySyncRequest(User1Information.InboxCollectionId);
        Sync(syncRequest);

        syncRequest.RequestData.Collections[0].Options = new Request.Options[] { option };
        syncRequest.RequestData.Collections[0].SyncKey = LastSyncKey;
        var syncResponse = Sync(syncRequest);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5412");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5412
        Site.CaptureRequirementIfAreEqual<int>(
            164,
            int.Parse(syncResponse.ResponseData.Status),
            5412,
            @"[In Common Status Codes] [The meaning of the status value 164 is] The BodyPartPreference node (as specified in [MS-ASAIRS] section 2.2.2.7) has an unsupported Type element (as specified in [MS-ASAIRS] section 2.2.2.22.4) value.<98>");
    }

    /// <summary>
    /// This test case is used to verify the server returns 118, when the message was already sent in a previous request.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC06_CommonStatusCode_118()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User1 calls SendMail command to send email messages to user2.

        var emailSubject = Common.GenerateResourceName(Site, "subject");
        var from = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
        var to = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var content = Common.GenerateResourceName(Site, "Default Email");
        var mime = Common.CreatePlainTextMime(from, to, null, null, emailSubject, content);
        var sendMailRequest = Common.CreateSendMailRequest(ClientId, false, mime);
        var responseSendMail = CMDAdapter.SendMail(sendMailRequest);
        Site.Assert.AreEqual<string>(
            string.Empty,
            responseSendMail.ResponseDataXML,
            "The server should return an empty xml response data to indicate SendMail command success.");
            
        SwitchUser(User2Information);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, emailSubject);
        GetMailItem(User2Information.InboxCollectionId, emailSubject);
        SwitchUser(User1Information);
        #endregion

        #region User1 calls SendMail command with the same ClientId again.

        // Use the same ClientId to call SendMail command again
        responseSendMail = CMDAdapter.SendMail(sendMailRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4848");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4848
        Site.CaptureRequirementIfAreEqual<string>(
            "118",
            responseSendMail.ResponseData.Status,
            4848,
            @"[In Common Status Codes] [The meaning of the status value 118 is] The message was already sent in a previous request.");

        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns 102, when call FolderSync command containing a FileReference element which is not defined in FolderSync request.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC07_CommonStatusCode_102()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region call FolderSync command containing a FileReference element which is not defined in FolderSync request.
        var request = new FolderSyncRequest();
        var requestData = new Request.FolderSync { FileReference = "0", FileReferenceSpecified = true };
        request.RequestData = requestData;

        var folderSyncResponse = CMDAdapter.FolderSync(request);
        Site.CaptureRequirementIfAreEqual<int>(
            102,
            int.Parse(folderSyncResponse.ResponseData.Status),
            4815,
            @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 102 is] The request contains WBXML but it could not be decoded into XML.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns 103, when The XML provided in the request does not follow the protocol requirements.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC08_CommonStatusCode_103()
    {
        #region Call SendMail command without Mime and AccountID.
        var sendMail = new Request.SendMail
        {
            ClientId = ClientId
        };

        var request = new SendMailRequest();
        request.RequestData = sendMail;

        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("12.1"))
        {
            try
            {
                var response = CMDAdapter.SendMail(request);
            }
            catch (WebException ex)
            {
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.BadRequest,
                    ((HttpWebResponse)ex.Response).StatusCode,
                    7547,
                    @"[In Common Status Codes] When protocol version 2.5, 12.0, or 12.1 is used, an HTTP 400 response is returned instead of this status value [103].");
            }
        }
        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("14.0") ||
            Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("14.1") ||
            Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("16.0"))
        {
            var response = CMDAdapter.SendMail(request);

            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(response.ResponseData.Status),
                4817,
                @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 103 is] The XML provided in the request does not follow the protocol requirements.");
        }
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns 119, when message being sent contains no recipient.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC09_CommonStatusCode_119()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Send message which contains no recipient.
        var emailSubject = Common.GenerateResourceName(Site, "subject");
        var body = Common.GenerateResourceName(Site, "Default Email");
        string from = from = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);

        var mime = Common.CreatePlainTextMime(from, string.Empty, null, null, emailSubject, body);
        var sendMailRequest = Common.CreateSendMailRequest(ClientId, false, mime);
        var response = CMDAdapter.SendMail(sendMailRequest);
            
        Site.CaptureRequirementIfAreEqual<int>(
            119,
            int.Parse(response.ResponseData.Status),
            4853,
            @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 119 is] The message being sent contains no recipient.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns 109, when the device type is either missing or has an invalid format.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC10_CommonStatusCode_109()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call FolderSync with a invaild format deviceType
        CMDAdapter.ChangeHeaderEncodingType(QueryValueType.PlainText);
        CMDAdapter.ChangeDeviceType("123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ");
        var response = FolderSync();

        Site.CaptureRequirementIfAreEqual<int>(
            109,
            int.Parse(response.ResponseData.Status),
            4830,
            @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 109 is] The device type is either missing or has an invalid format.");
        CMDAdapter.ChangeDeviceType(Common.GetConfigurationPropertyValue("DeviceType", Site));
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns 108, when the device ID is either missing or has an invalid format.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC11_CommonStatusCode_108()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call FolderSync with a invaild format device ID
        CMDAdapter.ChangeHeaderEncodingType(QueryValueType.PlainText);
        CMDAdapter.ChangeDeviceID("123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ");
        var response = FolderSync();

        Site.CaptureRequirementIfAreEqual<int>(
            108,
            int.Parse(response.ResponseData.Status),
            4828,
            @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 108 is] The device ID is either missing or has an invalid format.");

        CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceType", Site));
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns 126, when the user object in the directory service indicates that this user is not allowed to use ActiveSync.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC12_CommonStatusCode_126()
    {
        var disableUser=new UserInformation()
        {
            UserName = Common.GetConfigurationPropertyValue("User4Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User4Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        CMDAdapter.SwitchUser(disableUser.UserName, disableUser.UserPassword, disableUser.UserDomain);

        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("12.1"))
        {
            try
            {
                var response = FolderSync();
            }
            catch (WebException ex)
            {
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Forbidden,
                    ((HttpWebResponse)ex.Response).StatusCode,
                    7555,
                    @"[In Common Status Codes] When protocol version 2.5, 12.0, or 12.1 is used, an HTTP 403 response is returned instead of this status value [126].");
            }
        }

        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("14.0") ||
            Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("14.1") ||
            Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("16.0"))
        {
            var response = FolderSync();

            Site.CaptureRequirementIfAreEqual<int>(
                126,
                int.Parse(response.ResponseData.Status),
                4869,
                @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 126 is] The user object in the directory service indicates that this user is not allowed to use ActiveSync.");
        }
    }

    /// <summary>
    /// This test case is used to verify the server returns 150, when the value of either the ItemId element or the InstanceId element specified in the SmartReply or the SmartForward command request could not be found in the mailbox.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC13_CommonStatusCode_150()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call SendMail command to send plain text email messages to user2.
        var emailSubject = Common.GenerateResourceName(Site, "subject");
        SendPlainTextEmail(null, emailSubject, User1Information.UserName, User2Information.UserName, null);
        #endregion

        #region Call Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncChangeResponse = GetMailItem(User2Information.InboxCollectionId, emailSubject);
        var originalServerID = FindServerId(syncChangeResponse, "Subject", emailSubject);
        var originalContent = GetDataFromResponseBodyElement(syncChangeResponse, originalServerID);
        #endregion

        #region Record user name, folder collectionId and item subject that are used in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, emailSubject);
        #endregion

        #region Call Sync command to delete the email in user2's mailbox.
        var syncRequest = CreateSyncDeleteRequest(LastSyncKey, User2Information.InboxCollectionId, originalServerID);
        syncRequest.RequestData.Collections[0].DeletesAsMoves = false;
        syncRequest.RequestData.Collections[0].DeletesAsMovesSpecified = true;
        Sync(syncRequest);
        #endregion

        #region Call SmartForward command to forward messages without retrieving the full, original message from the server.
        var forwardSubject = string.Format("FW:{0}", emailSubject);
        var forwardFromUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User3Information.UserName, User3Information.UserDomain);
        var forwardContent = Common.GenerateResourceName(Site, "forward:body");
            
        var smartForwardRequest = CreateSmartForwardRequest(User2Information.InboxCollectionId, originalServerID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        #endregion

        Site.CaptureRequirementIfAreEqual<int>(
            150,
            int.Parse(smartForwardResponse.ResponseData.Status),
            4930,
            @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 150 is] The value of either the ItemId element (section 2.2.3.88) or the InstanceId element (section 2.2.3.87.2) specified in the SmartReply (section 2.2.2.19) or the SmartForward (section 2.2.2.18) command request could not be found in the mailbox.");
            
    }

    /// <summary>
    /// This test case is used to verify the server returns 145, when the The device claimed to be externally managed, but the server doesn't allow externally managed devices to sync.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S21_TC14_CommonStatusCode_145()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "This test case is not supported when the MS-ASProtocolVersion header is set to 12.1.. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User calls Provision command to download policies from server
        // Calls Provision command to download policies
        var provisionRequest = GenerateDefaultProvisionRequest();
        var provisionResponse = CMDAdapter.Provision(provisionRequest);
            

        // Get policyKey, policyType and statusCode from server response
        string policyKey = null;
        if (null != provisionResponse.ResponseData.Policies)
        {
            var policyInResponse = provisionResponse.ResponseData.Policies.Policy;
            if (policyInResponse != null)
            {
                policyKey = policyInResponse.PolicyKey;
            }
        }

        var policyType = provisionResponse.ResponseData.Policies.Policy.PolicyType;
        var data = provisionResponse.ResponseData.Policies.Policy.Data;
        var statusCode = provisionResponse.ResponseData.Status;
        #endregion

        #region User calls Provision command to acknowledge policies.

        // Set acknowledgeStatus value to 1, means accept the policy.
        var provisionAcknowledgeRequest = GenerateDefaultProvisionRequest();
        provisionAcknowledgeRequest.RequestData.Policies.Policy.PolicyKey = policyKey;
        provisionAcknowledgeRequest.RequestData.Policies.Policy.Status = "4";

        // Calls Provision command
        var provisionAcknowledgeResponse = CMDAdapter.Provision(provisionAcknowledgeRequest);
        statusCode = provisionAcknowledgeResponse.ResponseData.Status;

        Site.CaptureRequirementIfAreEqual<byte>(
            145,
            statusCode,
            4917,
            @"[In Common Status Codes] When the protocol version is 14.0, 14.1 or 16.0,  [The meaning of the status value 145 is] The device claimed to be externally managed, but the server doesn't allow externally managed devices to sync.");
        #endregion
    }
    #endregion
}