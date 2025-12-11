namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using System;
using System.Collections.Generic;
using System.Text;
using Common;
using Common.DataStructures;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Request = Common.Request;
using Response = Common.Response;

/// <summary>
/// This scenario is used to test the SmartForward command.
/// </summary>
[TestClass]
public class S17_SmartForward : TestSuiteBase
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
    /// This test case is used to verify the server returns an empty response, when mail forward successfully.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC01_SmartForward_Success()
    {
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

        #region Call SmartForward command to forward messages without retrieving the full, original message from the server.
        var forwardSubject = $"FW:{emailSubject}";
        var smartForwardRequest = CreateDefaultForwardRequest(originalServerID, forwardSubject, User2Information.InboxCollectionId);
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        #endregion

        #region Verify Requirements MS-ASCMD_R568, MS-ASCMD_R4407
        // If the message was forwarded successfully, server returns an empty response without XML body, then MS-ASCMD_R568, MS-ASCMD_R4407 are verified.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R568");

        // Verify MS-ASCMD requirement: MS-ASCMD_R568
        Site.CaptureRequirementIfAreEqual<string>(
            string.Empty,
            smartForwardResponse.ResponseDataXML,
            568,
            @"[In SmartForward] If the message was forwarded successfully, the server returns an empty response.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4407");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4407
        Site.CaptureRequirementIfAreEqual<string>(
            string.Empty,
            smartForwardResponse.ResponseDataXML,
            4407,
            @"[In Status(SmartForward and SmartReply)] If the SmartForward command request [or SmartReply command request] succeeds, no XML body is returned in the response.");
        #endregion

        #region After user2 forwarded email to user3, sync user3 mailbox changes
        SwitchUser(User3Information);
        var syncForwardResult = GetMailItem(User3Information.InboxCollectionId, forwardSubject);
        var forwardItemServerID = FindServerId(syncForwardResult, "Subject", forwardSubject);
        var forwardItemContent = GetDataFromResponseBodyElement(syncForwardResult, forwardItemServerID);
        #endregion

        #region Record user name, folder collectionId and item subject that are used in this case
        RecordCaseRelativeItems(User3Information, User3Information.InboxCollectionId, forwardSubject);

        #endregion

        // Compare original content with forward content
        var isContained = forwardItemContent.Contains(originalContent);

        #region Verify Requirements  MS-ASCMD_R532            

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R532");

        // Verify MS-ASCMD requirement: MS-ASCMD_R532
        Site.CaptureRequirementIfIsNotNull(
            forwardItemServerID,
            532,
            @"[In SmartForward] The SmartForward command is used by clients to forward messages without retrieving the full, original message from the server.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify server returns status code, when SmartForward is failed.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC02_SmartForward_Fail()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The AccountID element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The AccountID element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        #region Call SendMail command to send one plain text email messages to user2.
        var emailSubject = Common.GenerateResourceName(Site, "subject");
        SendPlainTextEmail(null, emailSubject, User1Information.UserName, User2Information.UserName, null);
        #endregion

        #region Call Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncChangeResponse = GetMailItem(User2Information.InboxCollectionId, emailSubject);
        var originalServerId = FindServerId(syncChangeResponse, "Subject", emailSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are used in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, emailSubject);
        #endregion

        #region Call SmartForward command to forward messages without retrieving the full, original message from the server
        // Create invalid SmartForward request
        var smartForwardRequest = new SmartForwardRequest
        {
            RequestData = new Request.SmartForward
            {
                ClientId = Guid.NewGuid().ToString(),
                Source = new Request.Source
                {
                    FolderId = User2Information.InboxCollectionId,
                    ItemId = originalServerId
                },
                Mime = string.Empty,
                AccountId = "InvalidValueAccountID"
            }
        };

        smartForwardRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>
        {
            {
                CmdParameterName.CollectionId, User2Information.InboxCollectionId
            },
            {
                CmdParameterName.ItemId, "5:" + originalServerId
            }
        });

        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4408");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4408
        // If SmartForward operation failed, server will return Status element in response, then MS-ASCMD_4408 is verified.
        Site.CaptureRequirementIfIsNotNull(
            smartForwardResponse.ResponseData.Status,
            4408,
            @"[In Status(SmartForward and SmartReply)] If the SmartForward command request [or SmartReply command request] fails, the Status element contains a code that indicates the type of failure.");
    }

    /// <summary>
    /// This test case is used to verify when the SmartForward command is used for an appointment, the original message is included as an attachment in the outgoing message.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC03_SmartForwardAppointment()
    {
        #region User1 calls Sync command uploading one calendar item to create one appointment
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var calendar = new Calendar
        {
            OrganizerEmail = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain),
            OrganizerName = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain),
            UID = Guid.NewGuid().ToString(),
            Subject = meetingRequestSubject
        };

        SyncAddCalendar(calendar);

        // Calls Sync command to sync user1's calendar folder
        var syncUser1CalendarFolder = GetMailItem(User1Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemId = FindServerId(syncUser1CalendarFolder, "Subject", meetingRequestSubject);

        // Record items need to be cleaned up.
        RecordCaseRelativeItems(User1Information, User1Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User1 calls smartForward command to forward mail to user2
        var forwardFromUser = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardSubject = $"FW:{meetingRequestSubject}";
        var forwardContent = Common.GenerateResourceName(Site, "forward:Appointment body");
        var smartForwardRequest = CreateSmartForwardRequest(User1Information.CalendarCollectionId, calendarItemId, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "If SmartForward command executes successfully, server should return empty xml data");
        #endregion

        #region User2 calls Sync command to get the forward mail sent by user1
        SwitchUser(User2Information);
        var user2MailboxChange = GetMailItem(User2Information.InboxCollectionId, forwardSubject);

        // Record items need to be cleaned up.
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, forwardSubject);
        var mailItemServerId = FindServerId(user2MailboxChange, "Subject", forwardSubject);

        var attachments = (Response.Attachments)GetElementValueFromSyncResponse(user2MailboxChange, mailItemServerId, Response.ItemsChoiceType8.Attachments);
        Site.Assert.AreEqual<int>(1, attachments.Items.Length, "Server should return one attachment, if SmartForward one appointment executes successfully.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify if the value of the InstanceId element is invalid, the server returns status value 104.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC04_SmartForwardWithInvalidInstanceId()
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

        #region User2 calls Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls SmartForward command to forward the calendar item to user3 with invalid InstanceId value in SmartForward request
        var forwardFromUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User3Information.UserName, User3Information.UserDomain);
        var forwardSubject = $"FW:{meetingRequestSubject}";
        var forwardContent = Common.GenerateResourceName(Site, "forward:Meeting Instance body");
        var smartForwardRequest = CreateSmartForwardRequest(User2Information.CalendarCollectionId, calendarItemID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);

        // Set instanceID with format not the same as required format "2010-03-20T22:40:00.000Z".
        var instanceID = DateTime.Now.ToString();
        smartForwardRequest.RequestData.Source.InstanceId = instanceID;
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R541");

        // Verify MS-ASCMD requirement: MS-ASCMD_R541
        Site.CaptureRequirementIfAreEqual<string>(
            "104",
            smartForwardResponse.ResponseData.Status,
            541,
            @"[In SmartForward] If the value of the InstanceId element is invalid, the server responds with Status element (section 2.2.3.162.15) value 104, as specified in section 2.2.4.");
    }

    /// <summary>
    /// This test case is used to verify when SmartForward is applied to a recurring meeting, the InstanceId element specifies the ID of a particular occurrence in the recurring meeting.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC05_SmartForwardWithInstanceIdSuccess()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Recurrences cannot be added in protocol version 16.0");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Recurrences cannot be added in protocol version 16.1");

        #region User1 calls SendMail command to send one recurring meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
        #endregion

        #region User2 calls Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncMeetingMailResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);

        // Record relative items for clean up.
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to accept the meeting
        var serverIDForMeetingRequest = FindServerId(syncMeetingMailResponse, "Subject", meetingRequestSubject);
        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

        // If the user accepts the meeting request, the meeting request mail will be deleted and calendar item will be created.
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        Site.Assert.IsNotNull(meetingResponseResponse.ResponseData.Result[0].CalendarId, "If the meeting was accepted, server should return calendarId in response");
        RemoveRecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        GetMailItem(User2Information.DeletedItemsCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.DeletedItemsCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls Sync command to sync user calendar changes
        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        var startTime = (string)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);
        var recurrence = (Response.Recurrence)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.Recurrence);
        Site.Assert.IsNotNull(recurrence, "If user2 received recurring meeting request, the calendar item should contain recurrence element");

        // Record relative items for clean up.
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls SmartForward command to forward the calendar item to user3 with correct InstanceId value in SmartForward request
        var forwardFromUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User3Information.UserName, User3Information.UserDomain);
        var forwardSubject = $"FW:{meetingRequestSubject}";
        var forwardContent = Common.GenerateResourceName(Site, "forward:Meeting Instance body");
        var smartForwardRequest = CreateSmartForwardRequest(User2Information.CalendarCollectionId, calendarItemID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);

        // Set instanceID with format the same as required format "2010-03-20T22:40:00.000Z".
        var instanceID = ConvertInstanceIdFormat(startTime);
        smartForwardRequest.RequestData.Source.InstanceId = instanceID;
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "If SmartForward command executes successfully, server should return empty xml data");
        #endregion

        #region After user2 forwards email to user3, sync user3 mailbox changes
        SwitchUser(User3Information);
        var syncForwardResult = GetMailItem(User3Information.InboxCollectionId, forwardSubject);
        var forwardItemServerID = FindServerId(syncForwardResult, "Subject", forwardSubject);

        // Sync user3 Calendar folder 
        var syncUser3CalendarFolder = GetMailItem(User3Information.CalendarCollectionId, forwardSubject);
        var user3CalendarItemID = FindServerId(syncUser3CalendarFolder, "Subject", forwardSubject);

        // Record email items for clean up
        RecordCaseRelativeItems(User3Information, User3Information.InboxCollectionId, forwardSubject);
        RecordCaseRelativeItems(User3Information, User3Information.CalendarCollectionId, forwardSubject);
        #endregion

        #region Record the meeting forward notification mail which sent from server to User1.
        SwitchUser(User1Information);
        var notificationSubject = "Meeting Forward Notification: " + forwardSubject;
        RecordCaseRelativeItems(User1Information, User1Information.DeletedItemsCollectionId, notificationSubject);
        GetMailItem(User1Information.DeletedItemsCollectionId, notificationSubject);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R538");

        // Verify MS-ASCMD requirement: MS-ASCMD_R538
        // If the calendar item with specified subject exists in user3 Calendar folder and email item exists in user3 Inbox folder which means user3 gets the forwarded mail.
        Site.CaptureRequirementIfIsTrue(
            user3CalendarItemID != null && forwardItemServerID != null,
            538,
            @"[In SmartForward] When SmartForward is applied to a recurring meeting, the InstanceId element (section 2.2.3.83.2) specifies the ID of a particular occurrence in the recurring meeting.");
    }

    /// <summary>
    /// This test case is used to verify when SmartForward request without the InstanceId element, the implementation forward the entire recurring meeting.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC06_SmartForwardRecurringMeetingWithoutInstanceId()
    {
        Site.Assume.IsTrue(Common.IsRequirementEnabled(5834, Site), "[In Appendix A: Product Behavior] If SmartForward is applied to a recurring meeting and the InstanceId element is absent, the implementation does forward the entire recurring meeting. (Exchange 2007 and above follow this behavior.)");
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Recurrences cannot be added in protocol version 16.0");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "Recurrences cannot be added in protocol version 16.1");

        #region User1 calls SendMail command to send one recurring meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);

        // Record relative items for clean up
        RecordCaseRelativeItems(User1Information, User1Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User1Information, User1Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);

        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);

        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        var recurrence = (Response.Recurrence)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.Recurrence);
        Site.Assert.IsNotNull(recurrence, "If user2 received recurring meeting request, the calendar item should contain recurrence element");

        // Record relative items for clean up
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls SmartForward command to forward the calendar item to user3 without InstanceId element in SmartForward request
        var forwardFromUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User3Information.UserName, User3Information.UserDomain);
        var forwardSubject = $"FW:{meetingRequestSubject}";
        var forwardContent = Common.GenerateResourceName(Site, "forward:Meeting Instance body");
        var smartForwardRequest = CreateSmartForwardRequest(User2Information.CalendarCollectionId, calendarItemID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);

        smartForwardRequest.RequestData.Source.InstanceId = null;
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "If SmartForward command executes successfully, server should return empty xml data");
        #endregion

        #region After user2 forwards email to user3, sync user3 mailbox changes
        SwitchUser(User3Information);
        var syncForwardResult = GetMailItem(User3Information.InboxCollectionId, forwardSubject);
        var forwardItemServerID = FindServerId(syncForwardResult, "Subject", forwardSubject);
        var user3Meeting = (Response.MeetingRequest)GetElementValueFromSyncResponse(syncForwardResult, forwardItemServerID, Response.ItemsChoiceType8.MeetingRequest);

        // Record email items for clean up
        RecordCaseRelativeItems(User3Information, User2Information.InboxCollectionId, forwardSubject);
        RecordCaseRelativeItems(User3Information, User2Information.CalendarCollectionId, forwardSubject);
        #endregion

        #region Check the meeting forward notification mail which is sent from server to User1.
        SwitchUser(User1Information);
        var notificationSubject = "Meeting Forward Notification: " + forwardSubject;
        CheckMeetingForwardNotification(User1Information, notificationSubject);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5834");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5834
        // If the calendar item with specified subject contains Recurrence element, which indicates user3 received the entire meeting request.
        Site.CaptureRequirementIfIsTrue(
            user3Meeting.Recurrences != null && forwardItemServerID != null,
            5834,
            @"[In Appendix A: Product Behavior] If SmartForward is applied to a recurring meeting and the InstanceId element is absent, the implementation does forward the entire recurring meeting. (Exchange 2007 and above follow this behavior.)");
    }

    /// <summary>
    /// This test case is used to verify when ReplaceMime is present in the request, the body or attachment is not included.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S17_TC07_SmartForward_ReplaceMime()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "ReplaceMime is not support when MS-ASProtocolVersion header is set to 12.1.MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call SendMail command to send plain text email messages to user2.
        var emailSubject = Common.GenerateResourceName(Site, "subject");
        var emailBody = Common.GenerateResourceName(Site, "NormalAttachment_Body");
        SendEmailWithAttachment(emailSubject, emailBody);
        #endregion

        #region Call Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncChangeResponse = GetMailItem(User2Information.InboxCollectionId, emailSubject);
        var originalServerID = FindServerId(syncChangeResponse, "Subject", emailSubject);
        var originalContent = GetDataFromResponseBodyElement(syncChangeResponse, originalServerID);
        var originalAttachments = GetEmailAttachments(syncChangeResponse, emailSubject);
        Site.Assert.IsTrue(originalAttachments != null && originalAttachments.Length == 1, "The email should contain a single attachment.");

        #endregion

        #region Record user name, folder collectionId and item subject that are used in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, emailSubject);
        #endregion

        #region Call SmartForward command to forward messages with ReplaceMime.
        var forwardSubject = $"FW:{emailSubject}";
        var smartForwardRequest = CreateDefaultForwardRequest(originalServerID, forwardSubject, User2Information.InboxCollectionId);
        smartForwardRequest.RequestData.ReplaceMime = string.Empty;
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        #endregion

        #region After user2 forwarded email to user3, sync user3 mailbox changes
        SwitchUser(User3Information);
        var syncForwardResult = GetMailItem(User3Information.InboxCollectionId, forwardSubject);
        var forwardItemServerID = FindServerId(syncForwardResult, "Subject", forwardSubject);
        var forwardItemContent = GetDataFromResponseBodyElement(syncForwardResult, forwardItemServerID);
        var forwardAttachments = GetEmailAttachments(syncForwardResult, forwardSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are used in this case
        RecordCaseRelativeItems(User3Information, User3Information.InboxCollectionId, forwardSubject);
        #endregion

        // Compare original content with forward content
        var isContained = forwardItemContent.Contains(originalContent);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3775");

        Site.Assert.IsNull(
            forwardAttachments,
            @"The attachment should not be returned");

        Site.CaptureRequirementIfIsFalse(
            isContained,
            3775,
            @"[In ReplaceMime] When the ReplaceMime element is present, the server MUST not include the body or attachments of the original message being forwarded.");
    }
    #endregion

    #region Private methods
    /// <summary>
    /// Try to parse the no separator time string to DateTime
    /// </summary>
    /// <param name="time">The specified DateTime string</param>
    /// <returns>Return the DateTime with instanceId specified format</returns>
    private static string ConvertInstanceIdFormat(string time)
    {
        var stringBuilder = new StringBuilder();
        stringBuilder.Append(time.Substring(0, 4));
        stringBuilder.Append("-");
        stringBuilder.Append(time.Substring(4, 2));
        stringBuilder.Append("-");
        stringBuilder.Append(time.Substring(6, 5));
        stringBuilder.Append(":");
        stringBuilder.Append(time.Substring(11, 2));
        stringBuilder.Append(":");
        stringBuilder.Append(time.Substring(13, 2));
        stringBuilder.Append(".000");
        stringBuilder.Append(time.Substring(15));
        return stringBuilder.ToString();
    }

    /// <summary>
    /// Set sync request application data with calendar value
    /// </summary>
    /// <param name="calendar">The calendar instance</param>
    /// <returns>The application data for sync request</returns>
    private static Request.SyncCollectionAddApplicationData SetApplicationDataFromCalendar(Calendar calendar)
    {
        var applicationData = new Request.SyncCollectionAddApplicationData();
        var elementName = new List<Request.ItemsChoiceType8>();
        var elementValue = new List<object>();

        // Set application data
        elementName.Add(Request.ItemsChoiceType8.Timezone);
        elementValue.Add(calendar.Timezone);

        elementName.Add(Request.ItemsChoiceType8.Subject);
        elementValue.Add(calendar.Subject);

        elementName.Add(Request.ItemsChoiceType8.Sensitivity);
        elementValue.Add(calendar.Sensitivity);

        elementName.Add(Request.ItemsChoiceType8.BusyStatus);
        elementValue.Add(calendar.BusyStatus);

        elementName.Add(Request.ItemsChoiceType8.AllDayEvent);
        elementValue.Add(calendar.AllDayEvent);

        applicationData.ItemsElementName = elementName.ToArray();
        applicationData.Items = elementValue.ToArray();
        return applicationData;
    }

    /// <summary>
    /// Create default SmartForward request to forward an item from user 2 to user 3.
    /// </summary>
    /// <param name="originalServerID">The item serverID</param>
    /// <param name="forwardSubject">The forward mail subject</param>
    /// <param name="senderCollectionId">The sender inbox collectionId</param>
    /// <returns>The SmartForward request</returns>
    private SmartForwardRequest CreateDefaultForwardRequest(string originalServerID, string forwardSubject, string senderCollectionId)
    {
        var forwardFromUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User3Information.UserName, User3Information.UserDomain);
        var forwardContent = Common.GenerateResourceName(Site, "forward:body");
        var smartForwardRequest = CreateSmartForwardRequest(senderCollectionId, originalServerID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);
        return smartForwardRequest;
    }

    /// <summary>
    /// Add a meeting or appointment to server
    /// </summary>
    /// <param name="calendar">the calendar item</param>
    private void SyncAddCalendar(Calendar calendar)
    {
        var applicationData = SetApplicationDataFromCalendar(calendar);

        GetInitialSyncResponse(User1Information.CalendarCollectionId);
        var addCalendar = new Request.SyncCollectionAdd
        {
            ClientId = ClientId,
            ApplicationData = applicationData
        };

        var syncAddCalendarRequest = CreateSyncAddRequest(LastSyncKey, User1Information.CalendarCollectionId, addCalendar);
        var syncAddCalendarResponse = CMDAdapter.Sync(syncAddCalendarRequest);

        // Get data from response
        var syncCollections = (Response.SyncCollections)syncAddCalendarResponse.ResponseData.Item;
        Response.SyncCollectionsCollectionResponses syncResponses = null;
        for (var index = 0; index < syncCollections.Collection[0].ItemsElementName.Length; index++)
        {
            if (syncCollections.Collection[0].ItemsElementName[index] == Response.ItemsChoiceType10.Responses)
            {
                syncResponses = (Response.SyncCollectionsCollectionResponses)syncCollections.Collection[0].Items[index];
                break;
            }
        }

        Site.Assert.AreEqual(1, syncResponses.Add.Length, "User only upload one calendar item");
        var statusCode = int.Parse(syncResponses.Add[0].Status);
        Site.Assert.AreEqual(1, statusCode, "If upload calendar item successful, server should return status 1");
    }
    #endregion
}