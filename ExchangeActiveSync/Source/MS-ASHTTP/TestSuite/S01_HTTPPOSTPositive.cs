namespace Microsoft.Protocols.TestSuites.MS_ASHTTP;

using System;
using System.Collections.Generic;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Response = Common.Response;

/// <summary>
/// This scenario is designed to test HTTP POST commands with positive response.
/// </summary>
[TestClass]
public class S01_HTTPPOSTPositive : TestSuiteBase
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
    /// This test case is intended to validate the Content-Type response header.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC01_VerifyContentTypeResponseHeader()
    {
        #region Synchronize the collection hierarchy via FolderSync command.
        var folderSyncResponse = CallFolderSyncCommand();
        Site.Assert.IsNotNull(folderSyncResponse.Headers["Content-Type"], "The Content-Type header should not be null.");
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R219");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R219
        // If the content type is WBXML, this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "application/vnd.ms-sync.wbxml",
            folderSyncResponse.Headers["Content-Type"],
            219,
            @"[In Content-Type] If the response body is WBXML, the value of this [Content-Type] header MUST be ""application/vnd.ms-sync.wbxml"".");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R229");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R229
        // If the content type is WBXML, this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "application/vnd.ms-sync.wbxml",
            folderSyncResponse.Headers["Content-Type"],
            229,
            @"[In Response Body] The response body [except the Autodiscover command], if any, is in WBXML.");
    }

    /// <summary>
    /// This test case is intended to validate the Content-Encoding response header.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC02_VerifyContentEncodingResponseHeader()
    {
        #region Call FolderSync command without setting AcceptEncoding header.
        var folderSyncResponse = CallFolderSyncCommand();

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R412");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R412
        // If the Content-Encoding header doesn't exist, this requirement can be captured.
        Site.CaptureRequirementIfIsFalse(
            folderSyncResponse.Headers.ToString().Contains("Content-Encoding"),
            412,
            @"[In Response Headers] [[Header] Content-Encoding [is] required when the content is compressed ;] otherwise, this header [Content-Encoding] is not included.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R215");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R215
        // If the Content-Encoding header doesn't exist, this requirement can be captured.
        Site.CaptureRequirementIfIsFalse(
            folderSyncResponse.Headers.ToString().Contains("Content-Encoding"),
            215,
            @"[In Content-Encoding] Otherwise [if the response body is not compressed], it [Content-Encoding header] is omitted.");
        #endregion

        #region Call ConfigureRequestPrefixFields to set the AcceptEncoding header to "gzip".
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        requestPrefix.Add(HTTPPOSTRequestPrefixField.AcceptEncoding, "gzip");
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Call FolderSync command.
        folderSyncResponse = LoopCallFolderSyncCommand();

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R197");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R197
        // If the Content-Encoding header is gzip, this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "gzip",
            folderSyncResponse.Headers["Content-Encoding"],
            197,
            @"[In Response Headers] [Header] Content-Encoding [is] required when the content is compressed [; otherwise, this header [Content-Encoding] is not included].");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R214");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R214
        // If the Content-Encoding header is gzip, this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "gzip",
            folderSyncResponse.Headers["Content-Encoding"],
            214,
            @"[In Content-Encoding] This [Content-Encoding] header is required if the response body is compressed.");
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the AttachmentName command parameter with Base64 encoding query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC03_CommandParameter_AttachmentName_Base64()
    {
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        VerifyGetAttachmentsCommandParameter(QueryValueType.Base64);
    }

    /// <summary>
    /// This test case is intended to validate the AttachmentName command parameter with Plain Text query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC04_CommandParameter_AttachmentName_PlainText()
    {
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The GetAttachment command is not supported when the MS-ASProtocolVersion header is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        var getAttachmentResponse = VerifyGetAttachmentsCommandParameter(QueryValueType.PlainText);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R115");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R115
        // The GetAttachment command executes successfully when the AttachmentName command parameter is set, so this requirement can be captured.
        Site.CaptureRequirement(
            115,
            @"[In Command-Specific URI Parameters] [Parameter] AttachmentName [is used by] GetAttachment.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R487");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R487
        // The GetAttachment command executes successfully when the AttachmentName command parameter is set, so this requirement can be captured.
        Site.CaptureRequirement(
            487,
            @"[In Command-Specific URI Parameters] [Parameter] AttachmentName [is described as] A string that specifies the name of the attachment file to be retrieved. ");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R230");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R230
        // The GetAttachment command response has no xml body, so this requirement can be captured.
        Site.CaptureRequirementIfIsFalse(
            IsXml(getAttachmentResponse.ResponseDataXML),
            230,
            @"[In Response Body] Three commands have no XML body in certain contexts: GetAttachment, [Sync, and Ping].");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R490");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R490
        // The GetAttachment command was executed successfully, so this requirement can be captured.
        Site.CaptureRequirement(
            490,
            @"[In Command Codes] [Command] GetAttachment retrieves an e-mail attachment from the server.");
    }

    /// <summary>
    /// This test case is intended to validate the SaveInSent, CollectionId and ItemId command parameters with Base64 encoding query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC05_CommandParameter_SaveInSent_Base64()
    {
        VerifySaveInSentCommandParameter(QueryValueType.Base64, "1", "1", "1");
    }

    /// <summary>
    /// This test case is intended to validate the SaveInSent, CollectionId and ItemId command parameters with Plain Text query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC06_CommandParameter_SaveInSent_PlainText()
    {
        VerifySaveInSentCommandParameter(QueryValueType.PlainText, "T", "F", null);
    }

    /// <summary>
    /// This test case is intended to validate the LongId command parameter with Base64 encoding query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC07_CommandParameter_LongId_Base64()
    {
        VerifyLongIdCommandParameter(QueryValueType.Base64);
    }

    /// <summary>
    /// This test case is intended to validate the LongId command parameter with Plain Text query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC08_CommandParameter_LongId_PlainText()
    {
        VerifyLongIdCommandParameter(QueryValueType.PlainText);
    }

    /// <summary>
    /// This test case is intended to validate the Occurrence command parameter with Base64 encoding query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC09_CommandParameter_Occurrence_Base64()
    {
        #region User3 calls SendMail to send a meeting request to User2.
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        var sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
        var smartForwardSubject = Common.GenerateResourceName(Site, "SmartForward");

        // Call ConfigureRequestPrefixFields to change the QueryValueType to Base64.
        requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.Base64.ToString());
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        // Switch the current user to user3 and synchronize the collection hierarchy.
        SwitchUser(UserThreeInformation, true);

        // Call SendMail command to send the meeting request to User2.
        SendMeetingRequest(sendMailSubject);
        #endregion

        #region User2 calls MeetingResponse command to accept the received meeting request and forward it to User1.
        // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
        SwitchUser(UserTwoInformation, true);

        // Call Sync command to get the ServerId of the received meeting request.
        var itemServerId = LoopToSyncItem(UserTwoInformation.InboxCollectionId, sendMailSubject, true);

        // Add the received item to the item collection of User2.
        var inboxItemForUserTwo = new CreatedItems
        {
            CollectionId = UserTwoInformation.InboxCollectionId
        };
        inboxItemForUserTwo.ItemSubject.Add(sendMailSubject);
        UserTwoInformation.UserCreatedItems.Add(inboxItemForUserTwo);

        // Check the calendar item if is exist.
        var calendarItemServerId = LoopToSyncItem(UserTwoInformation.CalendarCollectionId, sendMailSubject, true);

        var calendarItemForUserTwo = new CreatedItems
        {
            CollectionId = UserTwoInformation.CalendarCollectionId
        };
        calendarItemForUserTwo.ItemSubject.Add(sendMailSubject);
        UserTwoInformation.UserCreatedItems.Add(calendarItemForUserTwo);

        // Call MeetingResponse command to accept the received meeting request.
        CallMeetingResponseCommand(UserTwoInformation.InboxCollectionId, itemServerId);

        // The accepted meeting request will be moved to Delete Items folder.
        itemServerId = LoopToSyncItem(UserTwoInformation.DeletedItemsCollectionId, sendMailSubject, true);
 
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R432");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R432
        // MeetingResponse command is executed successfully, so this requirement can be captured.
        Site.CaptureRequirementIfIsNotNull(
            itemServerId,
            432,
            @"[In Command Codes] [Command] MeetingResponse [is] used to accept [, tentatively accept , or decline] a meeting request in the user's Inbox folder.");

        // Remove the inboxItemForUserTwo object from the clean up list since it has been moved to Delete Items folder.
        UserTwoInformation.UserCreatedItems.Remove(inboxItemForUserTwo);
        AddCreatedItemToCollection("User2", UserTwoInformation.DeletedItemsCollectionId, sendMailSubject);

        // Call SmartForward command to forward the meeting to User1
        var startTime = (string)GetElementValueFromSyncResponse(UserTwoInformation.CalendarCollectionId, calendarItemServerId, Response.ItemsChoiceType8.StartTime);
        var occurrence = TestSuiteHelper.ConvertInstanceIdFormat(startTime);
        var userOneMailboxAddress = Common.GetMailAddress(UserOneInformation.UserName, UserOneInformation.UserDomain);
        var userTwoMailboxAddress = Common.GetMailAddress(UserTwoInformation.UserName, UserTwoInformation.UserDomain);
        CallSmartForwardCommand(userTwoMailboxAddress, userOneMailboxAddress, itemServerId, smartForwardSubject, null, null, occurrence);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R513");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R513
        // SmartForward command executed successfully with setting Occurrence command parameter, so this requirement can be captured.
        Site.CaptureRequirement(
            513,
            @"[In Command-Specific URI Parameters] [Parameter] Occurrence [is described as] A string that specifies the ID of a particular occurrence in a recurring meeting.");

        AddCreatedItemToCollection("User3", UserThreeInformation.DeletedItemsCollectionId, "Meeting Forward Notification: " + smartForwardSubject);
        #endregion

        #region User1 gets the forwarded meeting request 
        // Call ConfigureRequestPrefixFields to switch the credential to User1 and synchronize the collection hierarchy.
        SwitchUser(UserOneInformation, true);

        AddCreatedItemToCollection("User1", UserOneInformation.InboxCollectionId, smartForwardSubject);
        AddCreatedItemToCollection("User1", UserOneInformation.CalendarCollectionId, smartForwardSubject);

        // Call Sync command to get the ServerId of the received meeting request.
        LoopToSyncItem(UserOneInformation.InboxCollectionId, smartForwardSubject, true);

        // Call Sync command to get the ServerId of the calendar item.
        LoopToSyncItem(UserOneInformation.CalendarCollectionId, smartForwardSubject, true);
        #endregion

        #region User3 gets the Meeting Forward Notification email in the Deleted Items folder.
        // Call ConfigureRequestPrefixFields to switch the credential to User3 and synchronize the collection hierarchy.
        SwitchUser(UserThreeInformation, false);

        // Call Sync command to get the ServerId of the received meeting request and the notification email.
        LoopToSyncItem(UserThreeInformation.DeletedItemsCollectionId, "Meeting Forward Notification: " + smartForwardSubject, true);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R119");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R119
        // SmartForward command executed successfully with setting Occurrence command parameter, so this requirement can be captured.
        Site.CaptureRequirement(
            119,
            @"[In Command-Specific URI Parameters] [Parameter] Occurrence [is used by] SmartForward.");
        #endregion

        #region Reset the query value type.
        requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", Site);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the Occurrence command parameter with Plain Text query value type.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC10_CommandParameter_Occurrence_PlainText()
    {
        #region User3 calls SendMail to send a meeting request to User2
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        var sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
        var smartReplySubject = Common.GenerateResourceName(Site, "SmartReply");

        // Call ConfigureRequestPrefixFields to change the QueryValueType.
        requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString());
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

        // Switch the current user to user3 and synchronize the collection hierarchy.
        SwitchUser(UserThreeInformation, true);

        // Call SendMail command to send the meeting request to User2.
        SendMeetingRequest(sendMailSubject);
        #endregion

        #region User2 calls SmartReply command to reply the request to User3 with Occurrence command parameter
        // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
        SwitchUser(UserTwoInformation, true);

        // Call Sync command to get the ServerId of the received meeting request.
        var itemServerId = LoopToSyncItem(UserTwoInformation.InboxCollectionId, sendMailSubject, true);

        // Add the received item to the item collection of User2.
        var inboxItemForUserTwo = new CreatedItems
        {
            CollectionId = UserTwoInformation.InboxCollectionId
        };
        inboxItemForUserTwo.ItemSubject.Add(sendMailSubject);
        UserTwoInformation.UserCreatedItems.Add(inboxItemForUserTwo);

        // Call Sync command to get the ServerId of the calendar item.
        var calendarItemServerId = LoopToSyncItem(UserTwoInformation.CalendarCollectionId, sendMailSubject, true);

        var calendarItemForUserTwo = new CreatedItems
        {
            CollectionId = UserTwoInformation.CalendarCollectionId
        };
        calendarItemForUserTwo.ItemSubject.Add(sendMailSubject);
        UserTwoInformation.UserCreatedItems.Add(calendarItemForUserTwo);

        // Call SmartReply command with the Occurrence command parameter.
        var startTime = (string)GetElementValueFromSyncResponse(UserTwoInformation.CalendarCollectionId, calendarItemServerId, Response.ItemsChoiceType8.StartTime);
        var occurrence = TestSuiteHelper.ConvertInstanceIdFormat(startTime);
        var userTwoMailboxAddress = Common.GetMailAddress(UserTwoInformation.UserName, UserTwoInformation.UserDomain);
        var userThreeMailboxAddress = Common.GetMailAddress(UserThreeInformation.UserName, UserThreeInformation.UserDomain);
        CallSmartReplyCommand(userTwoMailboxAddress, userThreeMailboxAddress, itemServerId, smartReplySubject, null, null, occurrence);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R513");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R513
        // SmartReply command executed successfully with setting Occurrence command parameter, so this requirement can be captured.
        Site.CaptureRequirement(
            513,
            @"[In Command-Specific URI Parameters] [Parameter] Occurrence [is described as] A string that specifies the ID of a particular occurrence in a recurring meeting.");

        #endregion

        #region User3 gets the reply mail
        // Call ConfigureRequestPrefixFields to switch the credential to User3 and synchronize the collection hierarchy.
        SwitchUser(UserThreeInformation, false);

        // Call Sync command to get the ServerId of the received the reply.
        LoopToSyncItem(UserThreeInformation.InboxCollectionId, smartReplySubject, true);
        AddCreatedItemToCollection("User3", UserThreeInformation.InboxCollectionId, smartReplySubject);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R529");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R529
        // SmartReply command executed successfully with setting Occurrence command parameter, so this requirement can be captured.
        Site.CaptureRequirement(
            529,
            @"[In Command-Specific URI Parameters] [Parameter] Occurrence [is used by] SmartReply.");
        #endregion

        #region Reset the query value type and user credential.
        requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", Site);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        SwitchUser(UserOneInformation, false);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the FolderSync, FolderCreate, FolderUpdate and FolderDelete command codes.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC11_CommandCode_FolderRelatedCommands()
    {
        var folderNameToCreate = Common.GenerateResourceName(Site, "CreatedFolder");
        var folderNameToUpdate = Common.GenerateResourceName(Site, "UpdatedFolder");

        #region Call FolderSync command to synchronize the folder hierarchy.
        var folderSyncResponse = CallFolderSyncCommand();
        #endregion

        #region Call FolderCreate command to create a sub folder under Inbox folder.
        var folderCreateResponse = CallFolderCreateCommand(folderSyncResponse.ResponseData.SyncKey, folderNameToCreate, Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, Site));
        #endregion

        #region Call FolderSync command to synchronize the folder hierarchy.
        folderSyncResponse = CallFolderSyncCommand();

        // Get the created folder name using the ServerId returned in FolderSync response.
        var createdFolderName = GetFolderFromFolderSyncResponse(folderSyncResponse, folderCreateResponse.ResponseData.ServerId, "DisplayName");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R493");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R493
        // The created folder could be got in FolderSync response, so this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            folderNameToCreate,
            createdFolderName,
            493,
            @"[In Command Codes] [Command] FolderCreate creates an e-mail, [calendar, or contacts folder] on the server.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R491");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R491
        // R493 is captured, so this requirement can be captured directly.
        Site.CaptureRequirement(
            491,
            @"[In Command Codes] [Command] FolderSync synchronizes the folder hierarchy.");

        // Call the Sync command with latest SyncKey without change in folder.
        var syncResponse = CallSyncCommand(folderCreateResponse.ResponseData.ServerId);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R482");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R482
        // If response is not in xml, this requirement can be captured.
        Site.CaptureRequirementIfIsNull(
            syncResponse.SyncKey,
            482,
            @"[In Response Body] Three commands have no XML body in certain contexts: [GetAttachment,] Sync [, and Ping].");
        #endregion

        #region Call FolderUpdate command to update the name of the created folder to a new folder name and move the created folder to SentItems folder.
        CallFolderUpdateCommand(folderSyncResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId, folderNameToUpdate, Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, Site));
        #endregion

        #region Call FolderSync command to synchronize the folder hierarchy.
        folderSyncResponse = CallFolderSyncCommand();

        // Get the updated folder name using the ServerId returned in FolderSync response.
        var updatedFolderName = GetFolderFromFolderSyncResponse(folderSyncResponse, folderCreateResponse.ResponseData.ServerId, "DisplayName");
        var updatedParentId = GetFolderFromFolderSyncResponse(folderSyncResponse, folderCreateResponse.ResponseData.ServerId, "ParentId");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R431");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R431
        // The folder name is updated to the specified name, so this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            folderNameToUpdate,
            updatedFolderName,
            431,
            @"[In Command Codes] [Command] FolderUpdate is used to rename folders.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R68");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R68
        // The folder has been moved to the new created folder, so this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, Site),
            updatedParentId,
            68,
            @"[In Command Codes] [Command] FolderUpdate moves a folder from one location to another on the server.");
        #endregion

        #region Call FolderDelete to delete the folder from the server.
        CallFolderDeleteCommand(folderSyncResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Call FolderSync command to synchronize the folder hierarchy.
        folderSyncResponse = CallFolderSyncCommand();

        // Get the created folder name using the ServerId returned in FolderSync response.
        updatedFolderName = GetFolderFromFolderSyncResponse(folderSyncResponse, folderCreateResponse.ResponseData.ServerId, "DisplayName");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R496");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R496
        // The folder with the specified ServerId could not be got, so this requirement can be captured.
        Site.CaptureRequirementIfIsNull(
            updatedFolderName,
            496,
            @"[In Command Codes] [Command] FolderDelete deletes a folder from the server.");
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the Ping, MoveItems, GetItemEstimate and ItemOperations command codes.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC12_CommandCode_ItemRelatedCommands()
    {
        #region Call ConfigureRequestPrefixFields to change the query value type to Base64.
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.Base64.ToString());
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Call SendMail command to send email to User2.
        var sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
        var folderNameToCreate = Common.GenerateResourceName(Site, "CreatedFolder");
        var userOneMailboxAddress = Common.GetMailAddress(UserOneInformation.UserName, UserOneInformation.UserDomain);
        var userTwoMailboxAddress = Common.GetMailAddress(UserTwoInformation.UserName, UserTwoInformation.UserDomain);

        // Call SendMail command to send email to User2.
        CallSendMailCommand(userOneMailboxAddress, userTwoMailboxAddress, sendMailSubject, null);
        #endregion

        #region Call Ping command for changes that would require the client to resynchronize.
        // Switch the user to User2 and synchronize the collection hierarchy.
        SwitchUser(UserTwoInformation, true);

        // Call FolderSync command to synchronize the collection hierarchy.
        var folderSyncResponse = CallFolderSyncCommand();

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R428");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R428
        // The received email could not be got by FolderSync command, so this requirement can be captured.
        Site.CaptureRequirementIfIsFalse(
            folderSyncResponse.ResponseDataXML.Contains(sendMailSubject),
            428,
            @"[In Command Codes] But [command] FolderSync does not synchronize the items in the folders.");

        // Call Ping command for changes of Inbox folder.
        var pingResponse = CallPingCommand(UserTwoInformation.InboxCollectionId);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R504");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R504
        // The Status of the Ping command is 2 which means this folder needs to be synced, so this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "2",
            pingResponse.ResponseData.Status.ToString(),
            504,
            @"[In Command Codes] [Command] Ping requests that the server monitor specified folders for changes that would require the client to resynchronize.");
        #endregion

        #region Get the ServerId of the received email.
        // Call Sync command to get the ServerId of the received email.
        var receivedItemServerId = LoopToSyncItem(UserTwoInformation.InboxCollectionId, sendMailSubject, true);
        #endregion

        #region Call FolderCreate command to create a sub folder under Inbox folder.
        var folderCreateResponse = CallFolderCreateCommand(folderSyncResponse.ResponseData.SyncKey, folderNameToCreate, UserTwoInformation.InboxCollectionId);

        // Get the ServerId of the created folder.
        var createdFolder = folderCreateResponse.ResponseData.ServerId;
        #endregion

        #region Move the received email from Inbox folder to the created folder.
        CallMoveItemsCommand(receivedItemServerId, UserTwoInformation.InboxCollectionId, createdFolder);
        #endregion

        #region Get the moved email in the created folder.
        // Call Sync command to get the received email.
        receivedItemServerId = LoopToSyncItem(createdFolder, sendMailSubject, true);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R499");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R499
        // The moved email could be got in the new created folder, so this requirement can be captured.
        Site.CaptureRequirementIfIsNotNull(
            receivedItemServerId,
            499,
            @"[In Command Codes] [Command] MoveItems moves items from one folder to another.");
        #endregion

        #region Call ItemOperation command to fetch the email in Sent Items folder with AcceptMultiPart command parameter.
        var itemOperationResponse = CallItemOperationsCommand(createdFolder, receivedItemServerId, true);
        Site.Assert.IsNotNull(itemOperationResponse.Headers["Content-Type"], "The Content-Type header should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R94");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R94
        // The content is in multipart, so this requirement can be captured.
        Site.CaptureRequirementIfAreEqual<string>(
            "application/vnd.ms-sync.multipart",
            itemOperationResponse.Headers["Content-Type"],
            94,
            @"[In Command Parameters] [When flag] AcceptMultiPart [value is] 0x02, [the meaning is] setting this flag [AcceptMultiPart] to instruct the server to return the requested item in multipart format.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R95");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R95
        // R94 can be captured, so this requirement can be captured directly.
        Site.CaptureRequirement(
            95,
            @"[In Command Parameters] [When flag] AcceptMultiPart [value is] 0x02, [it is] valid for ItemOperations.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R534");

        // Verify MS-ASHTTP requirement: MS-ASHTTP_R534
        // R94 can be captured, so this requirement can be captured directly.
        Site.CaptureRequirement(
            534,
            @"[In Command Parameters] [Parameter] Options [ is used by] ItemOperations.");
        #endregion

        #region Call FolderDelete to delete a folder from the server.
        CallFolderDeleteCommand(folderCreateResponse.ResponseData.SyncKey, createdFolder);
        #endregion

        #region Reset the query value type and user credential.
        requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", Site);
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        SwitchUser(UserOneInformation, false);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the ResolveRecipients and Settings command codes.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC13_CommandCode_UserRelatedCommands()
    {
        #region Call ResolveRecipients command.
        var items = new object[] { Common.GetConfigurationPropertyValue("User1Name", Site) };
        var resolveRecipientsRequest = Common.CreateResolveRecipientsRequest(items);
        var resolveRecipientsResponse = HTTPAdapter.HTTPPOST(CommandName.ResolveRecipients, null, resolveRecipientsRequest.GetRequestDataSerializedXML());

        // Check the command is executed successfully.
        CheckResponseStatus(resolveRecipientsResponse.ResponseDataXML);
        #endregion

        #region Call Settings command.
        var settingsRequest = Common.CreateSettingsRequest();
        var settingsResponse = HTTPAdapter.HTTPPOST(CommandName.Settings, null, settingsRequest.GetRequestDataSerializedXML());

        // Check the command is executed successfully.
        CheckResponseStatus(settingsResponse.ResponseDataXML);
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate the ValidateCert command code.
    /// </summary>
    [TestCategory("MSASHTTP"), TestMethod]
    public void MSASHTTP_S01_TC14_CommandCode_ValidateCert()
    {
        #region Call ValidateCert command.
        var validateCertRequest = Common.CreateValidateCertRequest();
        var validateCertResponse = HTTPAdapter.HTTPPOST(CommandName.ValidateCert, null, validateCertRequest.GetRequestDataSerializedXML());

        // Check the command is executed successfully.
        CheckResponseStatus(validateCertResponse.ResponseDataXML);
        #endregion
    }
    #endregion

    #region Private methods
    /// <summary>
    /// Loop to call FolderSync command and get the response till FolderSync Response's Content-Encoding header is gzip.
    /// </summary>
    /// <returns>The response of FolderSync command.</returns>
    private FolderSyncResponse LoopCallFolderSyncCommand()
    {
        #region Loop to call FolderSync
        var counter = 0;
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        var upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var folderSyncResponse = CallFolderSyncCommand();
        while (counter < upperBound && !folderSyncResponse.Headers.ToString().Contains("Content-Encoding"))
        {
            Thread.Sleep(waitTime);
            folderSyncResponse = CallFolderSyncCommand();
            counter++;
        } 
        #endregion

        #region Call ConfigureRequestPrefixFields to reset the AcceptEncoding header.
        IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
        requestPrefix[HTTPPOSTRequestPrefixField.AcceptEncoding] = null;
        HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        #endregion

        #region Return FolderSync response
        Site.Assert.IsNotNull(folderSyncResponse.Headers["Content-Encoding"], "The Content-Encoding header should exist in the response headers after retry {0} times", counter);

        return folderSyncResponse; 
        #endregion
    }
    #endregion
}