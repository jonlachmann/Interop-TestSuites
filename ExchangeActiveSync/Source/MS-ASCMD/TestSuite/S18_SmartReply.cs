namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using System;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// This scenario is used to test the SmartReply command.
/// </summary>
[TestClass]
public class S18_SmartReply : TestSuiteBase
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
    /// This test case is used to verify the replied mail contains original message.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S18_TC01_SmartReply_ContainOriginalMessage()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Recurrences cannot be added in protocol version 16.0");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Recurrences cannot be added in protocol version 16.1");
        #region User1 calls SendMail command to send one recurring meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
        #endregion

        #region Call Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncChangeResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);

        // Get User2 original mail content
        var originalServerId = FindServerId(syncChangeResponse, "Subject", meetingRequestSubject);
        var originalContent = GetDataFromResponseBodyElement(syncChangeResponse, originalServerId);
        #endregion

        #region Record user name, folder collectionId and item subject that are useed in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region Call SmartReply command to reply to messages without instanceID in request
        var smartReplySubject = string.Format("REPLY: {0}", meetingRequestSubject);
        var smartReplyRequest = CreateDefaultReplyRequest(smartReplySubject, calendarItemID);

        // Add elements to smartReplyRequest
        smartReplyRequest.RequestData.Source.InstanceId = null;
        smartReplyRequest.RequestData.Source.FolderId = User2Information.CalendarCollectionId;
        smartReplyRequest.RequestData.Source.ItemId = calendarItemID;
        CMDAdapter.SmartReply(smartReplyRequest);
        #endregion

        #region Call Sync command to sync user1 mailbox changes
        SwitchUser(User1Information);
        var syncResponseOnUserOne = GetMailItem(User1Information.InboxCollectionId, smartReplySubject);

        // Get replied mail content
        var replyMailServerId = FindServerId(syncResponseOnUserOne, "Subject", smartReplySubject);
        var replyMailContent = GetDataFromResponseBodyElement(syncResponseOnUserOne, replyMailServerId);
        #endregion

        #region Record user name, folder collectionId and item subject that are useed in this case
        RecordCaseRelativeItems(User1Information, User1Information.InboxCollectionId, smartReplySubject);
        #endregion

        #region Verify Requirements MS-ASCMD_R5420, MS-ASCMD_R580, MS-ASCMD_R569

        if (Common.IsRequirementEnabled(5420, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5420");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5420
            // If SmartReply command executes successful, the reply mail content contains the original mail content
            Site.CaptureRequirementIfIsTrue(
                replyMailContent.Contains(originalContent),
                5420,
                @"[In Appendix A: Product Behavior] If SmartReply is applied to a recurring meeting and the InstanceId element is absent, the implementation does reply for the entire recurring meeting. (Exchange 2007 and above follow this behavior.)");
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R580");

        // Verify MS-ASCMD requirement: MS-ASCMD_R580
        Site.CaptureRequirementIfIsTrue(
            replyMailContent.Contains(originalContent),
            580,
            @"[In SmartReply] The full text of the original message is retrieved and sent by the server.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R569");

        // Verify MS-ASCMD requirement: MS-ASCMD_R569
        Site.CaptureRequirementIfIsTrue(
            replyMailContent.Contains(originalContent),
            569,
            @"[In SmartReply] The SmartReply command is used by clients to reply to messages without retrieving the full, original message from the server.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the server returns an empty response, when mail is replied successfully.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S18_TC02_SmartReply_Success()
    {
        #region Call SendMail command to send one plain text email to user2.
        var emailSubject = Common.GenerateResourceName(Site, "subject");
        SendPlainTextEmail(null, emailSubject, User1Information.UserName, User2Information.UserName, null);
        #endregion

        #region Call Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncChangeResponse = GetMailItem(User2Information.InboxCollectionId, emailSubject);
        var originalServerId = FindServerId(syncChangeResponse, "Subject", emailSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are useed in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, emailSubject);
        #endregion

        #region Call SmartReply command to reply to messages without retrieving the full, original message from the server
        var smartReplySubject = string.Format("REPLY: {0}", emailSubject);
        var smartReplyRequest = CreateDefaultReplyRequest(smartReplySubject, originalServerId);

        var smartReplyResponse = CMDAdapter.SmartReply(smartReplyRequest);

        #endregion

        #region Call Sync command to sync user1 mailbox changes
        SwitchUser(User1Information);
        GetMailItem(User1Information.InboxCollectionId, smartReplySubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are useed in this case
        RecordCaseRelativeItems(User1Information, User1Information.InboxCollectionId, smartReplySubject);
        #endregion

        #region Verify Requirements MS-ASCMD_R605, MS-ASCMD_R5776
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R605");

        // Verify MS-ASCMD requirement: MS-ASCMD_R605
        // If the message was sent successfully, the server won't return any XML data, then MS-ASCMD_R605 is verified.
        Site.CaptureRequirementIfAreEqual<string>(
            string.Empty,
            smartReplyResponse.ResponseDataXML,
            605,
            @"[In SmartReply] If the message was sent successfully, the server returns an empty response.");
            
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5776");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5776
        // If the message was sent successfully, the server won't return any XML data, then MS-ASCMD_R776 is verified.
        Site.CaptureRequirementIfAreEqual<string>(
            string.Empty,
            smartReplyResponse.ResponseDataXML,
            5776,
            @"[In Status(SmartForward and SmartReply)] If the [SmartForward command request or] SmartReply command request succeeds, no XML body is returned in the response.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify if the value of the InstanceId element is invalid, the server responds with Status value 104.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S18_TC03_SmartReply_Status104()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User1 calls SendMail command to send one recurring meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
        #endregion

        #region User2 calls Sync command to sync user2 mailbox changes.
        SwitchUser(User2Information);
        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 creates SmartReply request with invalid InstanceId value, then calls SmartReply command.
        // Set instanceID with format that is not the same as required "2010-03-20T22:40:00.000Z".
        var instanceID = DateTime.Now.ToString();
        var smartReplySubject = string.Format("REPLY: {0}", meetingRequestSubject);
        var smartReplyRequest = CreateDefaultReplyRequest(smartReplySubject, calendarItemID);

        // Add instanceId element to smartReplyRequest.
        smartReplyRequest.RequestData.Source.InstanceId = instanceID;
        smartReplyRequest.RequestData.Source.FolderId = User2Information.CalendarCollectionId;
        var smardReplyResponse = CMDAdapter.SmartReply(smartReplyRequest);
        #endregion

        #region Verify Requirements MS-ASCMD_R5777, MS-ASCMD_R578
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5777");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5777
        // If server returns status code which has value, that indicates the type of failure then MS-ASCMD_R5777 is verified.
        Site.CaptureRequirementIfIsNotNull(
            smardReplyResponse.ResponseData.Status,
            5777,
            @"[In Status(SmartForward and SmartReply)] If the [SmartForward command request or ] SmartReply command request fails, the Status element contains a code that indicates the type of failure.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R578");

        // Verify MS-ASCMD requirement: MS-ASCMD_R578
        Site.CaptureRequirementIfAreEqual<string>(
            "104",
            smardReplyResponse.ResponseData.Status,
            578,
            @"[In SmartReply] If the value of the InstanceId element is invalid, the server responds with Status element (section 2.2.3.162.15) value 104, as specified in section 2.2.4.");
        #endregion
    }
    #endregion

    #region Private method
    /// <summary>
    /// Create default SmartReply request.
    /// </summary>
    /// <param name="replySubject">Reply email subject.</param>
    /// <param name="originalServerId">The  server ID of the original email.</param>
    /// <returns>Smart Reply request.</returns>
    private SmartReplyRequest CreateDefaultReplyRequest(string replySubject, string originalServerId)
    {
        var smartReplyFromUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var smartReplyToUser = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
        var smartReplyContent = Common.GenerateResourceName(Site, "reply:body");
        var smartReplyRequest = CreateSmartReplyRequest(User2Information.InboxCollectionId, originalServerId, smartReplyFromUser, smartReplyToUser, string.Empty, string.Empty, replySubject, smartReplyContent);
        return smartReplyRequest;
    }
    #endregion
}