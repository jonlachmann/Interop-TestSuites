namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using System;
using System.Collections.Generic;
using System.Threading;
using Common;
using Common.DataStructures;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Request = Common.Request;
using Response = Common.Response;

/// <summary>
/// This scenario is used to test the MeetingResponse command.
/// </summary>
[TestClass]
public class S09_MeetingResponse : TestSuiteBase
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

    #region Test Cases
    /// <summary>
    /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC01_MeetingResponse_AcceptMeeting()
    {
        #region User1 calls SendMail command to send one meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region Get new added meeting request email
        // Switch to user2 mailbox
        SwitchUser(User2Information);

        // Sync Inbox folder
        var syncResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var serverIDForMeetingRequest = FindServerId(syncResponse, "Subject", meetingRequestSubject);
        var meetingRequest = (Response.MeetingRequest)GetElementValueFromSyncResponse(syncResponse, serverIDForMeetingRequest, Response.ItemsChoiceType8.MeetingRequest);

        // Sync Calendar folder
        var syncCalendarBeforeMeetingResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemServerID = FindServerId(syncCalendarBeforeMeetingResponse, "Subject", meetingRequestSubject);
        var messageClass = (string)GetElementValueFromSyncResponse(syncResponse, serverIDForMeetingRequest, Response.ItemsChoiceType8.MessageClass);
        #endregion

        #region Verify Requirements MS-ASCMD_R5068, MS-ASCMD_R5085, MS-ASCMD_R5828
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5068");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5068
        Site.CaptureRequirementIfAreEqual<string>(
            "IPM.Schedule.Meeting.Request",
            messageClass,
            5068,
            @"[In Receiving and Accepting Meeting Requests] The message contains an email:MessageClass element (as specified in [MS-ASEMAIL] section 2.2.2.41) that has a value of ""IPM.Schedule.Meeting.Request"".");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5085");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5085
        // User calls Sync command to get Inbox folder changes, if Sync response contains MeetingRequest element then MS-ASCMD_R5085 is verified.
        Site.CaptureRequirementIfIsNotNull(
            meetingRequest,
            5085,
            @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 2*:] The server responds with airsync:Add elements (section 2.2.3.7.2) for items in the Inbox collection, including a meeting request item.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5828");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5828
        // User calls Sync command to get Inbox folder changes, if Sync response contains MeetingRequest element then MS-ASCMD_R5828 is verified.
        Site.CaptureRequirementIfIsNotNull(
            meetingRequest,
            5828,
            @"[In Receiving and Accepting Meeting Requests] Its [The message's] airsync:ApplicationData element (section 2.2.3.11) contains an email:MeetingRequest element (as specified in [MS-ASEMAIL] section 2.2.2.40).");
        #endregion

        #region Call method MeetingResponse to accept the meeting request in the user2's Inbox folder.
        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

        // If the user accepts the meeting request, the meeting request mail will be deleted and calendar item will be created.
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        #region Verify Requirements MS-ASCMD_R5082, MS-ASCMD_R4180, MS-ASCMD_R830, MS-ASCMD_R831, MS-ASCMD_R3843, MS-ASCMD_R5071, MS-ASCMD_R5089, MS-ASCMD_R5723

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5082");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5082
        // User calls Sync command with the SyncKey element value of 0 for Inbox folder, if this.LastSynKey is not null, means Sync operation success and server returned SyncKey value, then MS-ASCMD_R830 is verified.
        Site.CaptureRequirementIfIsNotNull(
            LastSyncKey,
            5082,
            @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 1:] The server responds with the airsync:SyncKey for the collection, to be used in successive synchronizations.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4180");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4180
        Site.CaptureRequirementIfAreEqual<string>(
            "1",
            meetingResponseResponse.ResponseData.Result[0].Status,
            4180,
            @"[In Status(MeetingResponse)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R830");

        // Verify MS-ASCMD requirement: MS-ASCMD_R830
        // If user accept the meeting request, server will return CalendarId element in response.
        Site.CaptureRequirementIfIsNotNull(
            meetingResponseResponse.ResponseData.Result[0].CalendarId,
            830,
            @"[In CalendarId] The CalendarId element is included in the MeetingResponse command response that is sent to the client if the meeting request was not declined.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R831");

        // Verify MS-ASCMD requirement: MS-ASCMD_R831
        // If MeetingResponse command executes successfully, the server will return calendarId element in the response, then MS-ASCMD_R831 is verified.
        Site.CaptureRequirementIfIsNotNull(
            meetingResponseResponse.ResponseData.Result[0].CalendarId,
            831,
            @"[In CalendarId] If the meeting is accepted [or tentatively accepted], the server adds a new item to the calendar and returns its server ID in the CalendarId element in the response.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3843");

        // Verify MS-ASCMD requirement: MS-ASCMD_R3843
        // If the meeting request is accepted, the server will respond calendarId element in response, which is the calendar item server ID, then MS-ASCMD_R4383 is verified.
        Site.CaptureRequirementIfIsNotNull(
            meetingResponseResponse.ResponseData.Result[0].CalendarId,
            3843,
            @"[In Result(MeetingResponse)] If the meeting request is accepted, the server ID of the calendar item is also returned.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5071");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5071
        // If the meeting request is accepted, the server will response calendarId element in response, which is the calendar item server ID, then MS-ASCMD_R5071 is verified.
        Site.CaptureRequirementIfIsNotNull(
            meetingResponseResponse.ResponseData.Result[0].CalendarId,
            5071,
            @"[In Receiving and Accepting Meeting Requests] If the response to the meeting is accepted [or is tentatively accepted], the server will add or update the corresponding calendar item and return its server ID in the meetingresponse:CalendarId element (section 2.2.3.18) of the response.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5089");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5089
        Site.CaptureRequirementIfIsTrue(
            meetingResponseResponse.ResponseData.Result[0].CalendarId != null && int.Parse(meetingResponseResponse.ResponseData.Result[0].Status) != 0,
            5089,
            @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 4:] The server sends a response that contains the MeetingResponse command request status along with the ID of the calendar item that corresponds to this meeting request if the meeting was not declined.");

        if (Common.IsRequirementEnabled(5723, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5723");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5723
            Site.CaptureRequirementIfIsTrue(
                meetingResponseResponse.ResponseData.Result[0].CalendarId != null && int.Parse(meetingResponseResponse.ResponseData.Result[0].Status) != 0,
                5723,
                @"[In Appendix A: Product Behavior] Implementation does use the MeetingResponse command to accept a meeting request in the user's Inbox folder. (Exchange 2007 and above follow this behavior.)");
        }
        #endregion

        #region Sync Calendar folder change and get accepted calendar item's serverID
        var syncCalendarResponseAfterAcceptMeeting = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponseAfterAcceptMeeting, "Subject", meetingRequestSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.DeletedItemsCollectionId, meetingRequestSubject);
        #endregion

        // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.0")
        {
            // Get calendar item responseType value
            var responseTypeBeforeMeetingResponse = (uint)GetElementValueFromSyncResponse(syncCalendarBeforeMeetingResponse, calendarItemServerID, Response.ItemsChoiceType8.ResponseType);
            var responseTypeAfterMeetingResponse = (uint)GetElementValueFromSyncResponse(syncCalendarResponseAfterAcceptMeeting, calendarItemID, Response.ItemsChoiceType8.ResponseType);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5094");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5094
            Site.CaptureRequirementIfIsTrue(
                responseTypeBeforeMeetingResponse != responseTypeAfterMeetingResponse && meetingResponseResponse.ResponseData.Result[0].CalendarId != null,
                5094,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 6:] The server responds with any changes to the Calendar folder caused by the last synchronization and the new calendar item for the accepted meeting.");
        }

        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "12.1")
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5094");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5094
            // If the meeting was accepted, server should return calendarId in response
            Site.CaptureRequirementIfIsTrue(
                meetingResponseResponse.ResponseData.Result[0].CalendarId != null,
                5094,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 6:] The server responds with any changes to the Calendar folder caused by the last synchronization and the new calendar item for the accepted meeting.");
        }
    }

    /// <summary>
    /// This test case is used to verify the response that has no CalendarId, when meeting response is declined.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC02_MeetingResponse_DeclineMeeting()
    {
        #region User1 calls SendMail command to send one meeting request to user2

        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region Get new added meeting request email
        // Switch to user2 mailbox
        SwitchUser(User2Information);
        var serverIDForMeetingRequest = GetItemServerIdFromSpecialFolder(User2Information.InboxCollectionId, meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse to decline the meeting request in the Inbox folder.
        var meetingResponseRequest = CreateMeetingResponseRequest(3, User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

        // If the user declines the meeting request, the meeting request mail will be deleted and no calendar item will be created.
        var responseMeetingResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        var itemServerIDInDeletefolder = GetItemServerIdFromSpecialFolder(User2Information.DeletedItemsCollectionId, meetingRequestSubject);
        Site.Assert.IsNotNull(itemServerIDInDeletefolder, "If user decline the meeting request, the meeting request mail should be deleted");
        DeleteAll(User2Information.DeletedItemsCollectionId);

        var syncInboxFolder = SyncChanges(User2Information.InboxCollectionId);
        var itemServerIDInInboxFolder = FindServerId(syncInboxFolder, "Subject", meetingRequestSubject);

        if (Common.IsRequirementEnabled(5725, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5725");

            // If call the MeetingResponse command to decline a meeting request, the original meeting request item will move from Inbox folder to DeleteItems folder.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5725
            Site.CaptureRequirementIfIsTrue(
                itemServerIDInDeletefolder != null && itemServerIDInInboxFolder == null,
                5725,
                @"[In Appendix A: Product Behavior] Implementation does use the MeetingResponse command to decline a meeting request in the user's Inbox folder. (Exchange 2007 and above follow this behavior.)");
        }
        #endregion

        #region Verify Requirements MS-ASCMD_R837, MS-ASCMD_R5072
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R837");

        // Verify MS-ASCMD requirement: MS-ASCMD_R837
        // If user declined the meeting the server response will not return calendarId element
        Site.CaptureRequirementIfIsNull(
            responseMeetingResponse.ResponseData.Result[0].CalendarId,
            837,
            @"[In CalendarId] If the meeting is declined, the response does not contain a CalendarId element.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5072");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5072
        // If user declined the meeting the server response will not return calendarId element
        Site.CaptureRequirementIfIsNull(
            responseMeetingResponse.ResponseData.Result[0].CalendarId,
            5072,
            @"[In Receiving and Accepting Meeting Requests] If the response to the meeting is declined, the response will not contain a meetingresponse:CalendarId element because the server will delete the corresponding calendar item.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify returned status value is 2, when userResponse is invalid.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC03_MeetingResponse_InvalidMeeting()
    {
        #region User1 calls SendMail command to send one meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region Get new added meeting request email
        SwitchUser(User2Information);
        var serverIDForMeetingRequest = GetItemServerIdFromSpecialFolder(User2Information.InboxCollectionId, meetingRequestSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse with invalid UserResponse element.
        // Set invalid UserResponse value "5"
        var meetingResponseRequest = CreateMeetingResponseRequest(5, User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);
        var responseMeetingResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4182");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4182
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(responseMeetingResponse.ResponseData.Result[0].Status),
            4182,
            @"[In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The client has sent a malformed or invalid item.");
    }

    /// <summary>
    /// This test case is used to verify RequestId is present in MeetingResponse command response if it was present in the corresponding MeetingResponse command request.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC04_MeetingResponse_RequestID()
    {
        #region User1 calls SendMail command to send one meeting request to user2

        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region Get new added meeting request email.
        SwitchUser(User2Information);
        var requestID = GetItemServerIdFromSpecialFolder(User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse to tentatively accept the meeting request in Inbox folder.
        // Set UserResponse value 2 to tentatively accepted
        var meetingResponseRequest = CreateMeetingResponseRequest(2, User2Information.InboxCollectionId, requestID, string.Empty);
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        Site.Assert.AreEqual<int>(1, int.Parse(meetingResponseResponse.ResponseData.Result[0].Status), "If MeetingResponse command executes successfully, server should return status 1");
        var itemServerID = GetItemServerIdFromSpecialFolder(User2Information.CalendarCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.DeletedItemsCollectionId, meetingRequestSubject);

        // If user tentatively accepted the meeting, the calendar item will be found in Calendar folder.
        if (Common.IsRequirementEnabled(5724, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5724");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5724
            Site.CaptureRequirementIfIsTrue(
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status) == 1 && itemServerID != null,
                5724,
                @"[In Appendix A: Product Behavior] Implementation does use the MeetingResponse command to tentatively accept a meeting request in the user's Inbox folder. (Exchange 2007 and above follow this behavior.)");
        }
        #endregion

        #region Verify Requirements MS-ASCMD_R5374, MS-ASCMD_R3807
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5374");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5374
        Site.CaptureRequirementIfAreEqual(
            requestID,
            meetingResponseResponse.ResponseData.Result[0].RequestId,
            5374,
            @"[In RequestId] The RequestId element is present in MeetingResponse command responses only if it was present in the corresponding MeetingResponse command request.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3807");

        // Verify MS-ASCMD requirement: MS-ASCMD_R3807
        // If server response contains RequestId element, then MS-ASCMD_3807 is verified.
        Site.CaptureRequirementIfIsNotNull(
            meetingResponseResponse.ResponseData.Result[0].RequestId,
            3807,
            @"[In RequestId] The RequestId element is also returned in the response to the client along with the status of the user's response to the meeting request.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify the MeetingResponse command response has status equals 2, when the request is referencing an item other than a meeting request, e-mail or calendar item.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC05_MeetingResponse_ResponseNonMeeting()
    {
        #region User1 calls SendMail command to send one meeting request to user2

        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region user2 get new added meeting request email
        SwitchUser(User2Information);
        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse to tentatively accept the meeting request with invalid request.
        // Create request with invalid RequestID value
        var meetingResponseRequest = CreateMeetingResponseRequest(2, User2Information.InboxCollectionId, "InvalidValue", string.Empty);
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4183");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4183
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            4183,
            @"[In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The request is referencing an item other than a meeting request, email, or calendar item.");
    }

    /// <summary>
    /// This test case is used to verify if the InstanceId is not a specified meeting request, server should return the status value is 2.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC06_MeetingResponse_NonExistInstanceId()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User1 calls SendMail command to send meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region user2 get new added meeting request email
        SwitchUser(User2Information);
        var serverIDForMeetingResponse = GetItemServerIdFromSpecialFolder(User2Information.InboxCollectionId, meetingRequestSubject);
        GetItemServerIdFromSpecialFolder(User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse to decline the meeting request in Inbox folder with invalid request.
        // Create invalid MeetingResponse request with invalid InstanceID value
        var meetingResponseRequest = CreateMeetingResponseRequest(3, User2Information.InboxCollectionId, serverIDForMeetingResponse, "InvalidValue");
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4186");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4186
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            4186,
            @"[In Status(MeetingResponse)] [In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The InstanceId element specifies a nonexistent instance or is null.");
    }

    /// <summary>
    /// This test case is used to verify if there are more than 100 Request elements listed in the MeetingResponse command request, the server will return status 103.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC07_MeetingResponse_Status103()
    {
        #region User1 calls SendMail command to send one meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region User2 calls Sync command to get new added meeting request email
        // Switch to user2 mailbox
        SwitchUser(User2Information);
        var meetingRequestServerID = GetItemServerIdFromSpecialFolder(User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 creates MeetingResponse command request with 101 Request elements

        var totalCount = 101;
        var serverIDList = new List<string>();
        for (var requestIndex = 0; requestIndex < totalCount; requestIndex++)
        {
            serverIDList.Add(meetingRequestServerID);
        }

        var meetingResponseRequest = CreateMultiMeetingResponseRequest(1, User2Information.InboxCollectionId, serverIDList, string.Empty);

        // Send MeetingResponse command request with 101 Request elements
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        if (Common.IsRequirementEnabled(5672, Site) || Common.IsRequirementEnabled(5670, Site))
        {
            RemoveRecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
            RecordCaseRelativeItems(User2Information, User2Information.DeletedItemsCollectionId, meetingRequestSubject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5670");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5670
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                5670,
                @"[In Appendix A: Product Behavior] Implementation does not limit the number of elements in command requests and not return the specified error if the limit is exceeded. (<118> Section 3.1.5.8: Exchange 2007 SP1 and Exchange 2010 do not limit the number of elements in command requests.) ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5672");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5672
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                5672,
                @"[In Appendix A: Product Behavior] Implementation does not limit the number of elements in command requests. (<119> Section 3.1.5.8: Exchange 2007 SP1 and Exchange 2010 do not limit the number of elements in command requests. )");
        }

        if (Common.IsRequirementEnabled(5671, Site) || Common.IsRequirementEnabled(5673, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5671");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5671
            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                5671,
                @"[In Appendix A: Product Behavior] Implementation does limit the number of elements in command requests and return the specified error if the limit is exceeded. (<118> Section 3.1.5.8: Update Rollup 6 for Exchange 2010 SP2 and Exchange 2013 do limit the number of elements in command requests.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5673");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5673
            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                5673,
                @"[In Appendix A: Product Behavior] Update Rollup 6 for implementation does use the specified limit values by default but can be configured to use different values. (<119> Section 3.1.5.8: Update Rollup 6 for Exchange 2010 SP2 and Exchange 2013 use the specified limit values by default but can be configured to use different values.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5650");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5650
            Site.CaptureRequirementIfAreEqual<int>(
                103,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                5650,
                @"[In Limiting Size of Command Requests] In MeetingResponse (section 2.2.2.9) command request, when the limit value of Request element is bigger than 100 (minimum 1, maximum 2,147,483,647), the error returned by server is Status element (section 2.2.3.162.8) value of 103.");
        }
    }

    /// <summary>
    /// This test case is used to verify if the InstanceId element specifies an email meeting request item, the server returns status code 2.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC08_MeetingResponse_RecurringMeetingInstanceIDInvalid()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User1 calls SendMail command to send one recurring meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
        #endregion

        #region User2 calls Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);

        // Get the meeting request mail from Inbox folder.
        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);

        // Get the calendar item from Calendar folder
        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        var startTime = (string)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);

        // Record relative items for clean up
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to accept the meetingRequest with Instance element referring to email meeting request item in MeetingResponseRequest
        var calendarStartTime = Common.GetNoSeparatorDateTime(startTime).ToUniversalTime();

        // Set invalid instanceID randomly.            
        var instanceID = new Random().ToString();
        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.InboxCollectionId, calendarItemID, instanceID);
            
        // Send MeetingResponseRequest with instanceID specifies a email meeting
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4185");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4185
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            4185,
            @"[In Status(MeetingResponse)] [In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The InstanceId element (section 2.2.3.78.1) specifies an email meeting request item.");
    }

    /// <summary>
    /// This test case is used to verify if the InstanceId element value specifies a non-recurring meeting, the server responds with a Status element value of 146.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC09_MeetingResponse_Status146()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User1 calls SendMail command to send one single meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region User2 calls Sync command to get new added meeting request
        // Switch to user2 mailbox
        SwitchUser(User2Information);

        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);

        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to accept the meetingRequest with InstanceId element
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        var startTime = (string)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);
        var calendarStartTime = Common.GetNoSeparatorDateTime(startTime).ToUniversalTime();
        var instanceID = calendarStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

        // Send MeetingResponseRequest with instanceID specifies a non-recurring meeting
        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.CalendarCollectionId, calendarItemID, instanceID);
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3196");

        // Verify MS-ASCMD requirement: MS-ASCMD_R3196
        Site.CaptureRequirementIfAreEqual<int>(
            146,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            3196,
            @"[In InstanceId(MeetingResponse)] If the InstanceId element value specifies a non-recurring meeting, the server responds with a Status element value of 146.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4919");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4919
        Site.CaptureRequirementIfAreEqual<int>(
            146,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            4919,
            @"[In Common Status Codes] [The meaning of the status value 146 is] The request tried to forward an occurrence of a meeting that has no recurrence.");
    }

    /// <summary>
    /// This test is used to verify implementation does use the MeetingResponse command to tentatively accept a meeting request in the user's Inbox folder
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC10_MeetingResponse_TentativeAcceptMeeting()
    {
        #region User1 calls SendMail command to send one meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region User2 calls Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncInboxResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var inboxItemID = FindServerId(syncInboxResponse, "Subject", meetingRequestSubject);

        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);

        // Get calendar item responseType value before meetingResponse
        // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
        uint responseTypeBeforeMeetingResponse = 0;
        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.0")
        {
            responseTypeBeforeMeetingResponse = (uint)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.ResponseType);
        }

        // Record relative items for clean up
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to tentative accept the meeting
        // Set to tentatively accept one of recurring meeting request instance
        var meetingResponseRequest = CreateMeetingResponseRequest(2, User2Information.InboxCollectionId, inboxItemID, null);
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        Site.Assert.AreEqual<int>(1, int.Parse(meetingResponseResponse.ResponseData.Result[0].Status), "If MeetingResponse command executes successfully, server should return status 1");
        RemoveRecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.DeletedItemsCollectionId, meetingRequestSubject);

        var syncCalendarResponseAfterMeetingResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region Verify Requirements MS-ASCMD_R5790, MS-ASCMD_R5678
        // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.0")
        {
            // Get calendar item responseType value
            var responseTypeAfterMeetingResponse = (uint)GetElementValueFromSyncResponse(syncCalendarResponseAfterMeetingResponse, calendarItemID, Response.ItemsChoiceType8.ResponseType);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5790");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5790
            Site.CaptureRequirementIfAreNotEqual<uint>(
                responseTypeBeforeMeetingResponse,
                responseTypeAfterMeetingResponse,
                5790,
                @"[In Receiving and Accepting Meeting Requests] If the response to the meeting is [accepted or is] tentatively accepted, the server will add or update the corresponding calendar item and return its server ID in the meetingresponse:CalendarId element (section 2.2.3.18) of the response.");
        }

        // If meetingResponse response returns new calendarId element, then MS-ASCMD_R5678 is verified
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5678");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5678
        Site.CaptureRequirementIfIsNotNull(
            meetingResponseResponse.ResponseData.Result[0].CalendarId,
            5678,
            @"[In CalendarId] If the meeting is [accepted or] tentatively accepted, the server adds a new item to the calendar and returns its server ID in the CalendarId element in the response.");
        #endregion
    }

    /// <summary>
    /// This test is used to verify implementation does use the MeetingResponse command to accept a meeting request in the user's Calendar folder.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC11_MeetingResponse_AcceptMeetingInCalendar()
    {
        #region User1 calls SendMail command to send one meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress,null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region User2 calls Sync and FolderSync commands to sync user2 mailbox changes
        SwitchUser(User2Information);
        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var calendarItemID = GetItemServerIdFromSpecialFolder(User2Information.CalendarCollectionId, meetingRequestSubject);

        // Record relative items for clean up
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to accept the meetingRequest
        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.CalendarCollectionId, calendarItemID, null);
        CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion
    }

    /// <summary>
    /// This test case is used to verify if the InstanceId element value specified is not in the proper format, the server responds with a Status value of 104.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC12_MeetingResponse_Status104()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region User1 calls SendMail command to send one recurring meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
        #endregion

        #region User2 calls Sync command to get new added meeting request
        // Switch to user2 mailbox
        SwitchUser(User2Information);
        GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);

        // Record related items that need to be cleaned up
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to accept the meetingRequest with wrong format Instance element in MeetingResponseRequest
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
        var startTime = (string)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);
        var calendarStartTime = Common.GetNoSeparatorDateTime(startTime).ToUniversalTime();

        // Set instanceID using month-day-yearThour:minute:second.milsecondZ format which is different from required format "2010-03-20T22:40:00.000Z".
        var instanceID = calendarStartTime.ToString("MM-dd-yyyyTHH:mm:ss.fffZ");

        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.CalendarCollectionId, calendarItemID, instanceID);

        // Send MeetingResponseRequest with instanceID specifies a non-recurring meeting
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        #endregion

        #region Verify Requirements MS-ASCMD_R3195, MS-ASCMD_R4819
        // If the InstanceId element value is not in the proper format, server returns status code 104 which means the value is in invalid format, then MS-ASCMD_R3195, MS-ASCMD_R4819 are verified.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3195");

        // Verify MS-ASCMD requirement: MS-ASCMD_R3195
        Site.CaptureRequirementIfAreEqual<int>(
            104,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            3195,
            @"[In InstanceId(MeetingResponse)] If the InstanceId element value specified is not in the proper format, the server responds with a Status element (section 2.2.3.162.8) value of 104.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4819");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4819
        Site.CaptureRequirementIfAreEqual<int>(
            104,
            int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
            4819,
            @"[In Common Status Codes] [The meaning of the status value 104 is] The request contains a timestamp that could not be parsed into a valid date and time.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify that server sends a substitute meeting invitation email.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC13_MeetingResponse_SubstituteMeetingInvitationEmail()
    {
        Site.Assume.IsTrue(Common.IsRequirementEnabled(5822, Site), "Exchange Server 2013 and above support Substitute Meeting Invitation.");

        #region User1 calls SendMail command to send one meeting request to user7
        var originalSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User7Information.UserName, User7Information.UserDomain);
        var calendar = CreateCalendar(originalSubject, attendeeEmailAddress,null);

        // Send a meeting request email to user7
        SendMeetingRequest(originalSubject, calendar);
        #endregion

        #region Get new added meeting request email in user7 mailbox
        // Switch to user7 mailbox
        SwitchUser(User7Information);

        // Sync Inbox folder
        GetMailItem(User7Information.InboxCollectionId, originalSubject);

        // Add the debug information.
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R7478");

        // If the protocol version is 16.0, nothing will be done in the method SendMeetingRequest.
        // So when the assertion in GetMailItem succeeds, R7478 can be verified.
        Site.CaptureRequirement(
            7478,
            @"[In Creating a Meeting or Appointment] In protocol version 16.0, the server will send meeting requests to the attendees automatically while processing the Sync command request that creates the meeting.");

        // Sync Calendar folder
        GetMailItem(User7Information.CalendarCollectionId, originalSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User7Information, User7Information.CalendarCollectionId, originalSubject);
        RecordCaseRelativeItems(User7Information, User7Information.InboxCollectionId, originalSubject);
        #endregion

        #region Get substitute invitation email in user8 mailbox
        // Switch to user8 mailbox, user8 is delegate of user7
        SwitchUser(User8Information);

        // Sync Inbox folder
        var substituteSyncResponse = GetSubstituteMailItem(User8Information.InboxCollectionId, originalSubject);
        var substituteInvitationEmailServerId = FindServerId(substituteSyncResponse, "ThreadTopic", originalSubject);
        #endregion

        #region Verify Requirements related receiving and accepting a meeting request
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5613");

        // If the element Body is not null, the response contains the informations of meeting, this requirement can be verified.
        // Verify MS-ASCMD requirement: MS-ASCMD_R5613
        Site.CaptureRequirementIfIsNotNull(
            GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Body),
            5613,
            @"[In Substitute Meeting Invitation Email] The value of element airsyncbase:Body is summary of meeting details.");

        var messageClass = (string)GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.MessageClass);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5822");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5822
        Site.CaptureRequirementIfAreEqual<string>(
            "IPM.Note",
            messageClass,
            5822,
            @"[In Appendix A: Product Behavior] Implementation does return substitute meeting invitation email messages. (Exchange 2013 and above follow this behavior.)");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5614");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5614
        Site.CaptureRequirementIfAreEqual<string>(
            "IPM.Note",
            messageClass,
            5614,
            @"[In Substitute Meeting Invitation Email] The value of element email:MessageClass is set to ""IPM.Note"".");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5604");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5604
        Site.CaptureRequirementIfIsTrue(
            ((string)GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.To)).ToLower(System.Globalization.CultureInfo.InvariantCulture).Contains(User8Information.UserName.ToLower(System.Globalization.CultureInfo.InvariantCulture)),
            5604,
            @"[In Substitute Meeting Invitation Email] The value of element email:To is set to the email address of the delegate.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5605");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5605
        // If Sync response does not contain CC element, then MS-ASCMD_R5605 is verified
        Site.CaptureRequirementIfIsNull(
            (string)GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Cc),
            5605,
            @"[In Substitute Meeting Invitation Email] The value of element email:Cc is blank.");

        var substituteInvitationEmailSubject = (string)GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Subject1);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5607");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5607
        Site.CaptureRequirementIfIsTrue(
            substituteInvitationEmailSubject.Contains(originalSubject),
            5607,
            @"[In Substitute Meeting Invitation Email] The value of element email:Subject is original subject prepended with explanatory text.");
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User8Information, User8Information.InboxCollectionId, substituteInvitationEmailSubject);
        #endregion
    }

    /// <summary>
    /// This test case is used to verify if delegate user forward substitute meeting invitation email, server will append the original meeting request to the forwarded message. 
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC14_MeetingResponse_ForwardSubstituteMeetingInvitationEmail()
    {
        Site.Assume.IsTrue(Common.IsRequirementEnabled(5822, Site), "Exchange Server 2013 and above support Substitute Meeting Invitation.");

        #region User1 calls SendMail command to send one meeting request to user7
        var originalSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User7Information.UserName, User7Information.UserDomain);
        var calendar = CreateCalendar(originalSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user7
        SendMeetingRequest(originalSubject, calendar);
        #endregion

        #region Get new added meeting request email in user7 mailbox
        // Switch to user7 mailbox
        SwitchUser(User7Information);

        // Sync Inbox folder
        GetMailItem(User7Information.InboxCollectionId, originalSubject);

        // Sync Calendar folder
        GetMailItem(User7Information.CalendarCollectionId, originalSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User7Information, User7Information.CalendarCollectionId, originalSubject);
        RecordCaseRelativeItems(User7Information, User7Information.InboxCollectionId, originalSubject);
        #endregion

        #region Get substitute invitation email in user8 mailbox
        // Switch to user8 mailbox
        SwitchUser(User8Information);

        // Sync Inbox folder
        var substituteSyncResponse = GetSubstituteMailItem(User8Information.InboxCollectionId, originalSubject);
        var substituteInvitationEmailServerId = FindServerId(substituteSyncResponse, "ThreadTopic", originalSubject);
        var substituteInvitationEmailSubject = (string)GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Subject1);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User8Information, User8Information.InboxCollectionId, substituteInvitationEmailSubject);
        #endregion

        #region User8 creates SmartForward request which forwards mail to user2
        var forwardFromUser = Common.GetMailAddress(User8Information.UserName, User8Information.UserDomain);
        var forwardToUser = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var forwardContent = Common.GenerateResourceName(Site, "forward:body");
        var smartForwardRequest = CreateSmartForwardRequest(User8Information.InboxCollectionId, substituteInvitationEmailServerId, forwardFromUser, forwardToUser, string.Empty, string.Empty, substituteInvitationEmailSubject, forwardContent);
        #endregion

        #region User8 calls SmartForward command
        var smartForwardResponse = CMDAdapter.SmartForward(smartForwardRequest);
        Site.Assert.IsTrue(string.IsNullOrEmpty(smartForwardResponse.ResponseDataXML), "If SmartForward command execute success, server will return empty");
        #endregion

        #region User2 calls Sync command to get mailbox change
        // Switch to user2 mailbox
        SwitchUser(User2Information);

        // Sync user2 Inbox folder
        GetMailItem(User2Information.InboxCollectionId, substituteInvitationEmailSubject);
        #endregion

        #region Record user name, folder collectionId and item subject that are generated in this case
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, substituteInvitationEmailSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, substituteInvitationEmailSubject);

        // Check the meeting forward notification mail which is sent from server to User1.
        SwitchUser(User1Information);
        var notificationSubject = "Meeting Forward Notification: " + substituteInvitationEmailSubject;
        CheckMeetingForwardNotification(User1Information, notificationSubject);
        #endregion
    }

    /// <summary>
    /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC15_MeetingResponse_AcceptMeetingInCalendarFolder()
    {
        Site.Assume.AreNotEqual<SutVersion>(SutVersion.ExchangeServer2007, Common.GetSutVersion(Site), "Exchange 2007 SP3 does not use MeetingResponse command on Calendar folder.");

        #region User1 calls SendMail command to send one meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region Get new added meeting request email
        // Switch to user2 mailbox
        SwitchUser(User2Information);

        // Sync Inbox folder
        var syncResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var serverIDForMeetingRequest = FindServerId(syncResponse, "Subject", meetingRequestSubject);

        // Sync Calendar folder
        var syncCalendarBeforeMeetingResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemServerID = FindServerId(syncCalendarBeforeMeetingResponse, "Subject", meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse to accept the meeting request in the user2's Calendar folder.
        var meetingResponseRequest = CreateMeetingResponseRequest(1, User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

        // If the user accepts the meeting request, the meeting request mail will be deleted and calendar item will be created.
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);

        var isVerifiedR5707 = meetingResponseResponse.ResponseData.Result[0].CalendarId != null && meetingResponseResponse.ResponseData.Result[0].Status != "0";

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR5707,
            5707,
            @"[In MeetingResponse] The MeetingResponse command is used to accept [, tentatively accept, or decline] a meeting request in the [user's Inbox folder or] Calendar folder.<3>");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC16_MeetingResponse_TentativelyAcceptInCalendarFolder()
    {
        Site.Assume.AreNotEqual<SutVersion>(SutVersion.ExchangeServer2007, Common.GetSutVersion(Site), "Exchange 2007 SP3 does not use MeetingResponse command on Calendar folder.");

        #region User1 calls SendMail command to send one meeting request to user2.
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region User2 calls Sync command to sync user2 mailbox changes
        SwitchUser(User2Information);
        var syncInboxResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var inboxItemID = FindServerId(syncInboxResponse, "Subject", meetingRequestSubject);

        var syncCalendarResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemID = FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);

        // Get calendar item responseType value before meetingResponse
        // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
        uint responseTypeBeforeMeetingResponse = 0;
        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.0")
        {
            responseTypeBeforeMeetingResponse = (uint)GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.ResponseType);
        }

        // Record relative items for clean up
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.CalendarCollectionId, meetingRequestSubject);
        #endregion

        #region User2 calls MeetingResponse command to tentative accept the meeting
        // Set to tentatively accept one of recurring meeting request instance
        var meetingResponseRequest = CreateMeetingResponseRequest(2, User2Information.InboxCollectionId, inboxItemID, null);
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);
        Site.Assert.AreEqual<int>(1, int.Parse(meetingResponseResponse.ResponseData.Result[0].Status), "If MeetingResponse command executes successfully, server should return status 1");
        var itemServerID = GetItemServerIdFromSpecialFolder(User2Information.CalendarCollectionId, meetingRequestSubject);
        RemoveRecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, meetingRequestSubject);
        RecordCaseRelativeItems(User2Information, User2Information.DeletedItemsCollectionId, meetingRequestSubject);

        var isVerifiedR5708 = itemServerID != null && meetingResponseResponse.ResponseData.Result[0].Status == "1";

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR5708,
            5708,
            @"[In MeetingResponse] The MeetingResponse command is used to [accept,] tentatively accept [, or decline] a meeting request in the [user's Inbox folder or] Calendar folder.<3>");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S09_TC17_MeetingResponse_DeclineInCalendarFolder()
    {
        Site.Assume.AreNotEqual<SutVersion>(SutVersion.ExchangeServer2007, Common.GetSutVersion(Site), "Exchange 2007 SP3 does not use MeetingResponse command on Calendar folder.");

        #region User1 calls SendMail command to send one meeting request to user2
        var meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
        var attendeeEmailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        var calendar = CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

        // Send a meeting request email to user2
        SendMeetingRequest(meetingRequestSubject, calendar);
        #endregion

        #region Get new added meeting request email
        // Switch to user2 mailbox
        SwitchUser(User2Information);

        // Sync Inbox folder
        var syncResponse = GetMailItem(User2Information.InboxCollectionId, meetingRequestSubject);
        var serverIDForMeetingRequest = FindServerId(syncResponse, "Subject", meetingRequestSubject);

        // Sync Calendar folder
        var syncCalendarBeforeMeetingResponse = GetMailItem(User2Information.CalendarCollectionId, meetingRequestSubject);
        var calendarItemServerID = FindServerId(syncCalendarBeforeMeetingResponse, "Subject", meetingRequestSubject);
        #endregion

        #region Call method MeetingResponse to decline the meeting request in the user2's Calendar folder.
        var meetingResponseRequest = CreateMeetingResponseRequest(3, User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

        // If the user declines the meeting request, the meeting request mail will be deleted and no calendar item will be created.
        var meetingResponseResponse = CMDAdapter.MeetingResponse(meetingResponseRequest);

        var syncInboxFolder = SyncChanges(User2Information.CalendarCollectionId);
        var itemServerIDInCalendarFolder = FindServerId(syncInboxFolder, "Subject", meetingRequestSubject);

        var isVerifiedR5709 = meetingResponseResponse.ResponseData.Result[0].Status == "1" && itemServerIDInCalendarFolder == null;

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR5709,
            5709,
            @"[In MeetingResponse] The MeetingResponse command is used to [accept, tentatively accept, or] decline a meeting request in the [user's Inbox folder or] Calendar folder.<3>");
        #endregion
    }

    #endregion

    #region Private methods
    /// <summary>
    /// Create a MeetingResponse request.
    /// </summary>
    /// <param name="userResponse">The way the user response the meeting.</param>
    /// <param name="collectionID">The collection id of the folder that contains the meeting request.</param>
    /// <param name="requestIDList">The server ID list of the meeting request message item.</param>
    /// <param name="instanceID">The instance ID of the recurring meeting to be modified.</param>
    /// <returns>The MeetingResponse request.</returns>
    private static MeetingResponseRequest CreateMultiMeetingResponseRequest(byte userResponse, string collectionID, List<string> requestIDList, string instanceID)
    {
        var requestList = new List<Request.MeetingResponseRequest>();
        foreach (var requestID in requestIDList)
        {
            var request = new Request.MeetingResponseRequest
            {
                CollectionId = collectionID,
                RequestId = requestID,
                UserResponse = userResponse
            };

            // Set the instanceId of the meeting request to response
            if (!string.IsNullOrEmpty(instanceID))
            {
                request.InstanceId = instanceID;
            }

            requestList.Add(request);
        }

        return Common.CreateMeetingResponseRequest(requestList.ToArray());
    }

    /// <summary>
    /// Get email with special threadTopic
    /// </summary>
    /// <param name="folderID">The folderID that store mail items</param>
    /// <param name="threadTopic">The thread topic</param>
    /// <returns>Sync result</returns>
    private SyncResponse GetSubstituteMailItem(string folderID, string threadTopic)
    {
        var counter = 0;
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        var retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var syncResult = SyncChanges(folderID);
        var serverID = FindServerId(syncResult, "ThreadTopic", threadTopic);
        while (serverID == null && counter < retryCount)
        {
            Thread.Sleep(waitTime);
            syncResult = SyncChanges(folderID);
            if (syncResult.ResponseDataXML != null)
            {
                serverID = FindServerId(syncResult, "ThreadTopic", threadTopic);
            }

            counter++;
        }

        Site.Assert.IsNotNull(serverID, "The email item with subject '{0}' should be found.", threadTopic);
        Site.Log.Add(LogEntryKind.Debug, "Find item successful Loop count {0}", counter);
        return syncResult;
    }

    /// <summary>
    /// Delete all items from the specified collection
    /// </summary>
    /// <param name="collectionId">The specified collection id</param>
    private void DeleteAll(string collectionId)
    {
        var request = new ItemOperationsRequest
        {
            RequestData = new Request.ItemOperations()
            {
                Items =
                [
                    new Request.ItemOperationsEmptyFolderContents
                    {
                        CollectionId = collectionId,
                        Options = new Request.ItemOperationsEmptyFolderContentsOptions
                        {
                            DeleteSubFolders = string.Empty
                        }
                    }
                ]
            }
        };

        var response = CMDAdapter.ItemOperations(request, DeliveryMethodForFetch.Inline);
        Site.Assert.IsTrue(response.ResponseData != null && response.ResponseData.Status == "1" && response.ResponseData.Response.EmptyFolderContents[0].Status == "1", "All items in the specified collection should be deleted successfully.");
    }
    #endregion
}