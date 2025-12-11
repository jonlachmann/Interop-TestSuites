namespace Microsoft.Protocols.TestSuites.MS_ASCAL;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ItemOperationsStore = Common.DataStructures.ItemOperationsStore;
using Request = Common.Request;
using SearchStore = Common.DataStructures.SearchStore;
using SyncItem = Common.DataStructures.Sync;
using SyncStore = Common.DataStructures.SyncStore;

/// <summary>
/// This scenario is to test Calendar class elements, which are not attached in a Meeting request, including synchronizing
/// the calendar on the server, fetching information of the calendar and searching a specific calendar.
/// </summary>
[TestClass]
public class S01_CalendarElement : TestSuiteBase
{
    #region Test Class initialize and clean up

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

    #region MSASCAL_S01_TC01_AllDayEvent

    /// <summary>
    /// This test case is designed to verify a calendar class with an AllDayEvent element via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC01_AllDayEvent()
    {
        #region Generate calendar subject and record them.
        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subjectWithAllDayEvent0 = Common.GenerateResourceName(Site, "subjectWithAllDayEvent0");
        var subjectWithAllDayEvent1 = Common.GenerateResourceName(Site, "subjectWithAllDayEvent1");

        #endregion

        #region Call Sync command to add a calendar with the element AllDayEvent setting as '0' to the server, and sync calendars from the server.

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithAllDayEvent0);
        calendarItem.Add(Request.ItemsChoiceType8.AllDayEvent, byte.Parse("0"));
        if (!IsActiveSyncProtocolVersion121
            && !IsActiveSyncProtocolVersion140
            && !IsActiveSyncProtocolVersion141)
        {
            var location = new Request.Location();
            location.DisplayName = Location;
            calendarItem.Add(Request.ItemsChoiceType8.Location, location);
        }

        AddSyncCalendar(calendarItem);
        var calendarWithAllDayEvent0 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithAllDayEvent0);
        Site.Assert.IsNotNull(calendarWithAllDayEvent0.Calendar, "The calendar with subject {0} should exist in server.", subjectWithAllDayEvent0);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithAllDayEvent0);
        #endregion

        Site.Assert.IsNotNull(calendarWithAllDayEvent0.Calendar.AllDayEvent, "The AllDayEvent element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R85");

        // Verify MS-ASCAL requirement: MS-ASCAL_R85
        Site.CaptureRequirementIfAreEqual<byte>(
            0,
            calendarWithAllDayEvent0.Calendar.AllDayEvent.Value,
            85,
            @"[In AllDayEvent][The value 0 means AllDayEvent] Is not an all day event.");

        if (IsActiveSyncProtocolVersion121
            || IsActiveSyncProtocolVersion140
            || IsActiveSyncProtocolVersion141)
        {
            #region Call Sync command to add a calendar with the element AllDayEvent setting as '1' and the StartTime and EndTime elements as midnight to midnight values to the server, and sync calendars from the server.

            calendarItem.Clear();
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithAllDayEvent1);
            calendarItem.Add(Request.ItemsChoiceType8.AllDayEvent, byte.Parse("1"));
            calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
            calendarItem.Add(Request.ItemsChoiceType8.EndTime, StartTime.AddDays(1).ToString("yyyyMMddTHHmmssZ"));

            AddSyncCalendar(calendarItem);
            var calendarWithAllDayEvent1 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithAllDayEvent1);
            Site.Assert.IsNotNull(calendarWithAllDayEvent1.Calendar, "The calendar with subject {0} should exist in server.", subjectWithAllDayEvent1);
            RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithAllDayEvent1);

            #endregion

            Site.Assert.IsNotNull(calendarWithAllDayEvent1.Calendar.AllDayEvent, "The AllDayEvent element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R86");

            // Verify MS-ASCAL requirement: MS-ASCAL_R86
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                calendarWithAllDayEvent1.Calendar.AllDayEvent.Value,
                86,
                @"[In AllDayEvent][The value 1 means AllDayEvent] Is an all day event.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R7911");

            // Verify MS-ASCAL requirement: MS-ASCAL_R7911
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                calendarWithAllDayEvent1.Calendar.AllDayEvent.Value,
                7911,
                @"[In AllDayEvent] The AllDayEvent element specifies whether the event represented by the calendar item runs for the entire day.");
        }
    }

    #endregion

    #region MSASCAL_S01_TC02_Sensitivity

    /// <summary>
    /// This test case is to verify a calendar class with a Sensitivity element via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC02_Sensitivity()
    {
        #region Generate calendar subject and record them.
        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subjectWithSensitivity0 = Common.GenerateResourceName(Site, "subjectWithSensitivity0");
        var subjectWithSensitivity1 = Common.GenerateResourceName(Site, "subjectWithSensitivity1");
        var subjectWithSensitivity2 = Common.GenerateResourceName(Site, "subjectWithSensitivity2");
        var subjectWithSensitivity3 = Common.GenerateResourceName(Site, "subjectWithSensitivity3");

        #endregion

        #region Call Sync command to add a calendar with the element Sensitivity setting as '0' to the server, and sync calendars from the server.

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithSensitivity0);
        calendarItem.Add(Request.ItemsChoiceType8.Sensitivity, byte.Parse("0"));

        AddSyncCalendar(calendarItem);

        var calendarWithSensitivity0 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithSensitivity0);
        Site.Assert.IsNotNull(calendarWithSensitivity0.Calendar, "The calendar with subject {0} should exist in server.", subjectWithSensitivity0);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithSensitivity0);
        #endregion

        Site.Assert.IsNotNull(calendarWithSensitivity0.Calendar.Sensitivity, "The Sensitivity element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R610");

        // Verify MS-ASCAL requirement: MS-ASCAL_R610
        Site.CaptureRequirementIfAreEqual<byte>(
            0,
            calendarWithSensitivity0.Calendar.Sensitivity.Value,
            610,
            @"[In Sensitivity] [The value] 0 [means] Normal.");

        #region Call Sync command to add a calendar with the element Sensitivity setting as '1' to the server, and sync calendars from the server.

        calendarItem.Clear();
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithSensitivity1);
        calendarItem.Add(Request.ItemsChoiceType8.Sensitivity, byte.Parse("1"));
        AddSyncCalendar(calendarItem);

        var calendarWithSensitivity1 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithSensitivity1);

        Site.Assert.IsNotNull(calendarWithSensitivity1.Calendar, "The calendar with subject {0} should exist in server.", subjectWithSensitivity1);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithSensitivity1);
        #endregion

        Site.Assert.IsNotNull(calendarWithSensitivity1.Calendar.Sensitivity, "The Sensitivity element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R611");

        // Verify MS-ASCAL requirement: MS-ASCAL_R611
        Site.CaptureRequirementIfAreEqual<byte>(
            1,
            calendarWithSensitivity1.Calendar.Sensitivity.Value,
            611,
            @"[In Sensitivity] [The value] 1 [means] Personal.");

        #region Call Sync command to add a calendar with the element Sensitivity setting as '2' to the server, and sync calendars from the server.

        calendarItem.Clear();
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithSensitivity2);
        calendarItem.Add(Request.ItemsChoiceType8.Sensitivity, byte.Parse("2"));
        AddSyncCalendar(calendarItem);

        var calendarWithSensitivity2 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithSensitivity2);

        Site.Assert.IsNotNull(calendarWithSensitivity2.Calendar, "The calendar with subject {0} should exist in server.", subjectWithSensitivity2);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithSensitivity2);
        #endregion

        Site.Assert.IsNotNull(calendarWithSensitivity2.Calendar.Sensitivity, "The Sensitivity element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R612");

        // Verify MS-ASCAL requirement: MS-ASCAL_R612
        Site.CaptureRequirementIfAreEqual<byte>(
            2,
            calendarWithSensitivity2.Calendar.Sensitivity.Value,
            612,
            @"[In Sensitivity] [The value] 2 [means] Private.");

        #region Call Sync command to add a calendar with the element Sensitivity setting as '3' to the server, and sync calendars from the server.

        calendarItem.Clear();
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithSensitivity3);
        calendarItem.Add(Request.ItemsChoiceType8.Sensitivity, byte.Parse("3"));
        AddSyncCalendar(calendarItem);

        var calendarWithSensitivity3 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithSensitivity3);

        Site.Assert.IsNotNull(calendarWithSensitivity3.Calendar, "The calendar with subject {0} should exist in server.", subjectWithSensitivity3);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithSensitivity3);
        #endregion

        Site.Assert.IsNotNull(calendarWithSensitivity3.Calendar.Sensitivity, "The Sensitivity element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R613");

        // Verify MS-ASCAL requirement: MS-ASCAL_R613
        Site.CaptureRequirementIfAreEqual<byte>(
            3,
            calendarWithSensitivity3.Calendar.Sensitivity.Value,
            613,
            @"[In Sensitivity] [The value] 3 [means]Confidential.");

        // According to above steps, requirement MS-ASCAL_R42111 can be covered directly.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R42111");

        // Verify MS-ASCAL requirement: MS-ASCAL_R42111
        Site.CaptureRequirement(
            42111,
            @"[In Sensitivity] As a top-level element of the Calendar class, the Sensitivity element specifies the recommended privacy policy for the calendar item.");
    }

    #endregion

    #region MSASCAL_S01_TC03_CalendarWithoutOptionalElements

    /// <summary>
    /// This test case is to verify a calendar class without optional elements via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC03_CalendarWithoutOptionalElements()
    {
        #region Call Sync command to add a calendar without optional elements to the server, and sync calendars from the server.
        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
        {
            {
                Request.ItemsChoiceType8.Subject, SubjectName
            }
        };

        // Set Calendar Subject Property
        var addCalendar = new Request.SyncCollectionAddApplicationData
        {
            Items = calendarItem.Values.ToArray<object>(),
            ItemsElementName = calendarItem.Keys.ToArray<Request.ItemsChoiceType8>()
        };

        // Sync to get the SyncKey
        var initializeSyncResponse = InitializeSync(User1Information.CalendarCollectionId, null);

        // Add the calendar item
        var syncRequest = TestSuiteHelper.CreateSyncAddRequest(User1Information.CalendarCollectionId, initializeSyncResponse.SyncKey, addCalendar);
        CALAdapter.Sync(syncRequest);

        var calendarWithoutOptionalElements = GetChangeItem(User1Information.CalendarCollectionId, SubjectName);

        Site.Assert.IsNotNull(calendarWithoutOptionalElements.Calendar, "The calendar with subject {0} should exist in server.", SubjectName);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, SubjectName);

        #endregion
    }

    #endregion

    #region MSASCAL_S01_TC04_MultipleElements

    /// <summary>
    /// This test case is designed to verify a calendar class with multiple elements via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC04_MultipleElements()
    {
        #region Call Sync command to add a calendar with the elements DTStamp and Reminder to the server, and sync calendars from the server.
        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var subjectWithDTStampAndReminder = Common.GenerateResourceName(Site, "subjectWithDTStampAndReminder");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithDTStampAndReminder);

        var reminder = "10";
        calendarItem.Add(Request.ItemsChoiceType8.Reminder, reminder);
        if (IsActiveSyncProtocolVersion121
            || IsActiveSyncProtocolVersion140
            || IsActiveSyncProtocolVersion141)
        {
            calendarItem.Add(Request.ItemsChoiceType8.Location1, Location);
            calendarItem.Add(Request.ItemsChoiceType8.DtStamp, DateTime.Now.ToString("yyyyMMddTHHmmssZ"));
        }

        AddSyncCalendar(calendarItem);

        var calendarWithDTStampAndReminder = GetChangeItem(User1Information.CalendarCollectionId, subjectWithDTStampAndReminder);

        Site.Assert.IsNotNull(calendarWithDTStampAndReminder.Calendar, "The calendar with subject {0} should exist in server.", subjectWithDTStampAndReminder);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithDTStampAndReminder);

        #endregion

        Site.Assert.IsNotNull(calendarWithDTStampAndReminder.Calendar.DtStamp, "The DtStamp element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R21911");

        // Verify MS-ASCAL requirement: MS-ASCAL_R21911
        // If Calendar.DtStamp is not null, it means the server returns the date and time that the calendar item was created or modified in response
        Site.CaptureRequirementIfIsNotNull(
            calendarWithDTStampAndReminder.Calendar.DtStamp.Value,
            21911,
            @"[In DtStamp] As a top-level element of the Calendar class, the DtStamp element specifies the date and time that the calendar item was created or modified [or the date and time at which the exception item was created or modified].");

        if (IsActiveSyncProtocolVersion121
            || IsActiveSyncProtocolVersion140
            || IsActiveSyncProtocolVersion141)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R29411");

            // Verify MS-ASCAL requirement: MS-ASCAL_R29411
            Site.CaptureRequirementIfAreEqual<string>(
                Location,
                calendarWithDTStampAndReminder.Calendar.Location,
                29411,
                @"[In Location] As a top-level element of the Calendar class, the Location element specifies the place where the event specified by the calendar item occurs.");
        }

        var areEqual = string.Compare(Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain), calendarWithDTStampAndReminder.Calendar.OrganizerEmail, StringComparison.CurrentCultureIgnoreCase);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R36211, expected email address is: {0},actually is :{1}", Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain), calendarWithDTStampAndReminder.Calendar.OrganizerEmail);

        // Verify MS-ASCAL requirement: MS-ASCAL_R36211
        Site.CaptureRequirementIfAreEqual<int>(
            0,
            areEqual,
            36211,
            @"[In OrganizerEmail] The OrganizerEmail element specifies the e-mail address of the user who created the calendar item.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R36711");

        // Verify MS-ASCAL requirement: MS-ASCAL_R36711
        Site.CaptureRequirementIfAreEqual<string>(
            User1Information.UserName,
            calendarWithDTStampAndReminder.Calendar.OrganizerName,
            36711,
            @"[In OrganizerName] The OrganizerName element specifies the name of the user who created the calendar item.");

        if (IsActiveSyncProtocolVersion160)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2169, expected email address is: {0},actually is :{1}", Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain), calendarWithDTStampAndReminder.Calendar.OrganizerEmail);

            // Verify MS-ASCAL requirement: MS-ASCAL_R2169
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                areEqual,
                2169,
                @"[In OrganizerEmail] [When protocol version 16.0 is used, the client MUST NOT include the OrganizerEmail element in command requests and] the server will use the email address of the current user.");

        }

        if (IsActiveSyncProtocolVersion161)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2169001, expected email address is: {0},actually is :{1}", Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain), calendarWithDTStampAndReminder.Calendar.OrganizerEmail);

            // Verify MS-ASCAL requirement: MS-ASCAL_R2169001
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                areEqual,
                2169001,
                @"[In OrganizerEmail] [When protocol version 16.1 is used, the client MUST NOT include the OrganizerEmail element in command requests and] the server will use the email address of the current user.");
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R39011");

        // Verify MS-ASCAL requirement: MS-ASCAL_R39011
        Site.CaptureRequirementIfAreEqual<string>(
            reminder,
            calendarWithDTStampAndReminder.Calendar.Reminder.ToString(),
            39011,
            @"[In Reminder] As a top-level element of the Calendar class, the Reminder element specifies the number of minutes before the calendar item's start time to display a reminder notice.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R44511");

        // Verify MS-ASCAL requirement: MS-ASCAL_R44511
        // If Calendar.Timezone is not null, it means the calendar item has Timezone element
        Site.CaptureRequirementIfIsNotNull(
            calendarWithDTStampAndReminder.Calendar.Timezone,
            44511,
            @"[In Timezone] The Timezone element specifies the time zone of the calendar item.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R46011");

        // Verify MS-ASCAL requirement: MS-ASCAL_R46011
        // If Calendar.UID is not null, it means the calendar item has UID element
        Site.CaptureRequirementIfIsNotNull(
            calendarWithDTStampAndReminder.Calendar.UID,
            46011,
            @"[In UID] The UID element specifies an ID that uniquely identifies a single event or recurring series.");

        Site.Assert.IsNotNull(calendarWithDTStampAndReminder.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R310");

        // Verify MS-ASCAL requirement: MS-ASCAL_R310
        // If Calendar.CalendarBody is not null, it means the calendar item has airsyncbase:Body element
        Site.CaptureRequirementIfAreEqual<byte>(
            0,
            calendarWithDTStampAndReminder.Calendar.MeetingStatus.Value,
            310,
            @"[In MeetingStatus][The value 0 means] The event is an appointment which has no attendees.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R33411");

        // Verify MS-ASCAL requirement: MS-ASCAL_R33411
        // If Calendar.NativeBodyType is not null, it means the calendar item has airsyncbase:NativeBodyType element
        Site.CaptureRequirementIfIsNotNull(
            calendarWithDTStampAndReminder.Calendar.NativeBodyType,
            33411,
            @"[In NativeBodyType] The airsyncbase:NativeBodyType element specifies how the body text of the calendar item is stored on the server.");
    }

    #endregion

    #region MSASCAL_S01_TC05_CalendarWithoutStartTimeEndTime

    /// <summary>
    /// This test case is designed to verify a calendar class without StartTime and EndTime elements via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC05_CalendarWithoutStartTimeEndTime()
    {
        #region Call Sync command to add a calendar without StartTime element and EndTime element to the server, and sync calendars from the server.
        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var subjectWithoutStartTimeAndEndTime = Common.GenerateResourceName(Site, "subjectWithoutStartTimeAndEndTime");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithoutStartTimeAndEndTime);

        calendarItem.Add(Request.ItemsChoiceType8.StartTime, null);
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, null);

        AddSyncCalendar(calendarItem);

        var calendarWithoutStartTimeAndEndTime = GetChangeItem(User1Information.CalendarCollectionId, subjectWithoutStartTimeAndEndTime);

        Site.Assert.IsNotNull(calendarWithoutStartTimeAndEndTime.Calendar, "The calendar with subject {0} should exist in server.", subjectWithoutStartTimeAndEndTime);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithoutStartTimeAndEndTime);

        #endregion

        Site.Assert.IsNotNull(calendarWithoutStartTimeAndEndTime.Calendar.StartTime, "The StartTime element should not be null.");
        Site.Assert.IsNotNull(calendarWithoutStartTimeAndEndTime.Calendar.DtStamp, "The DtStamp element should not be null.");
        Site.Assert.IsNotNull(calendarWithoutStartTimeAndEndTime.Calendar.EndTime, "The EndTime element should not be null.");

        // If start time is rounded to the nearest half hour.
        var isRoundedTime = (calendarWithoutStartTimeAndEndTime.Calendar.StartTime.Value.Minute == 0 || calendarWithoutStartTimeAndEndTime.Calendar.StartTime.Value.Minute == 30) && (((TimeSpan)(calendarWithoutStartTimeAndEndTime.Calendar.StartTime - calendarWithoutStartTimeAndEndTime.Calendar.DtStamp)).Minutes <= 30);

        Site.Assert.AreEqual<bool>(true, isRoundedTime, "StartTime should be rounded time");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52511");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52511
        Site.CaptureRequirementIfIsTrue(
            (calendarWithoutStartTimeAndEndTime.Calendar.EndTime.Value - calendarWithoutStartTimeAndEndTime.Calendar.StartTime.Value).Minutes == 30,
            52511,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If the server receives a Sync command request ([MS-ASCMD] section 2.2.2.19) to add a calendar event that is missing either the StartTime element (section 2.2.2.40), the EndTime element (section 2.2.2.18), or both, the server attempts to substitute values based on the current time, rounded to the nearest half hour, for the missing values. ");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52514");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52514
        Site.CaptureRequirementIfIsTrue(
            (calendarWithoutStartTimeAndEndTime.Calendar.EndTime.Value - calendarWithoutStartTimeAndEndTime.Calendar.StartTime.Value).Minutes == 30,
            52514,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If StartTime and EndTime are both absent the server sets the value of the StartTime element to the rounded current time, and sets the value of the EndTime element to the rounded current time plus 30 minutes.");
    }

    #endregion

    #region MSASCAL_S01_TC06_StartTimeAbsentEndTimePast

    /// <summary>
    /// This test case is designed to verify a calendar class with an EndTime element set as past time via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC06_StartTimeAbsentEndTimePast()
    {
        #region Call Sync command to add a calendar with the element EndTime setting as past time to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
        {
            {
                Request.ItemsChoiceType8.StartTime, null
            },
            {
                Request.ItemsChoiceType8.EndTime, DateTime.Now.AddYears(-5).ToString("yyyyMMddTHHmmssZ")
            }
        };

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52515");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52515
        Site.CaptureRequirementIfAreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            52515,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If StartTime is absent and EndTime is in the past the server includes a Status element with a value of 6 in the response, as specified in [MS-ASCMD] section 2.2.3.162.16, indicating an error occurred.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53915");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53915
        Site.CaptureRequirementIfAreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            53915,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The EndTime element (section 2.2.2.18) is included in a request and the StartTime element is not included in the request.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R5251611");

        // Verify MS-ASCAL requirement: MS-ASCAL_R5251611
        Site.CaptureRequirementIfAreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            5251611,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If the rounded current time is after the end time, the server includes a Status element with a value of 6 in the response, indicating an error occurred.");
    }

    #endregion

    #region MSASCAL_S01_TC07_StartTimeAbsentEndTimeFuture

    /// <summary>
    /// This test case is designed to verify a calendar class with an EndTime element set as future time via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC07_StartTimeAbsentEndTimeFuture()
    {
        #region Generate calendar subject and record them.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subjectWithFutureEndTime = Common.GenerateResourceName(Site, "subject");

        #endregion

        #region Call Sync command to add a calendar with the element EndTime setting as future time to the server, and sync calendars from the server.

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithFutureEndTime);
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, null);

        var futureEndtime = FutureTime.AddDays(2);
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, futureEndtime.ToString("yyyyMMddTHHmmssZ"));

        AddSyncCalendar(calendarItem);

        var calendarWithFutureEndTime = GetChangeItem(User1Information.CalendarCollectionId, subjectWithFutureEndTime);

        Site.Assert.IsNotNull(calendarWithFutureEndTime.Calendar, "The calendar with subject {0} should exist in server.", subjectWithFutureEndTime);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithFutureEndTime);
        #endregion

        Site.Assert.IsNotNull(calendarWithFutureEndTime.Calendar.StartTime, "The StartTime element should not be null.");
        Site.Assert.IsNotNull(calendarWithFutureEndTime.Calendar.DtStamp, "The DtStamp element should not be null.");
        Site.Assert.IsNotNull(calendarWithFutureEndTime.Calendar.EndTime, "The EndTime element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52516");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52516
        Site.CaptureRequirementIfIsTrue(
            (calendarWithFutureEndTime.Calendar.StartTime.Value - calendarWithFutureEndTime.Calendar.DtStamp.Value).Minutes <= 30 && calendarWithFutureEndTime.Calendar.EndTime.Value.ToUniversalTime().ToString("yyyyMMddTHHmmssZ").Equals(futureEndtime.ToString("yyyyMMddTHHmmssZ")),
            52516,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If StartTime is absent and EndTime is in the future the server sets the value of the StartTime element to the rounded current time and sets the value of the EndTime element to the value of the EndTime element in the request.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R23311");

        // Verify MS-ASCAL requirement: MS-ASCAL_R23311
        Site.CaptureRequirementIfAreEqual<string>(
            futureEndtime.ToString("yyyyMMddTHHmmssZ"),
            calendarWithFutureEndTime.Calendar.EndTime.Value.ToUniversalTime().ToString("yyyyMMddTHHmmssZ"),
            23311,
            @"[In EndTime] As a top-level element of the Calendar class, the EndTime element specifies the end time of the calendar item.");
    }

    #endregion

    #region MSASCAL_S01_TC08_StartTimePastEndTimeAbsent

    /// <summary>
    /// This test case is designed to verify a calendar class with a StartTime element set as past time via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC08_StartTimePastEndTimeAbsent()
    {
        #region Call Sync command to add a calendar with the element StartTime setting as past time to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var subjectWithPastStartTime = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithPastStartTime);
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, null);
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, PastTime.ToString("yyyyMMddTHHmmssZ"));

        AddSyncCalendar(calendarItem);

        var calendarWithPastStartTime = GetChangeItem(User1Information.CalendarCollectionId, subjectWithPastStartTime);

        Site.Assert.IsNotNull(calendarWithPastStartTime.Calendar, "The calendar with subject {0} should exist in server.", subjectWithPastStartTime);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithPastStartTime);

        #endregion

        Site.Assert.IsNotNull(calendarWithPastStartTime.Calendar.StartTime, "The StartTime element should not be null.");
        Site.Assert.IsNotNull(calendarWithPastStartTime.Calendar.DtStamp, "The DtStamp element should not be null.");
        Site.Assert.IsNotNull(calendarWithPastStartTime.Calendar.EndTime, "The EndTime element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52517");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52517
        Site.CaptureRequirementIfIsTrue(
            (calendarWithPastStartTime.Calendar.EndTime.Value - calendarWithPastStartTime.Calendar.DtStamp.Value.AddMinutes(30)).Minutes <= 30 && calendarWithPastStartTime.Calendar.StartTime.Value.ToUniversalTime().ToString("yyyyMMddTHHmmssZ").Equals(PastTime.ToString("yyyyMMddTHHmmssZ")),
            52517,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If StartTime is in the past and EndTime is absent the server sets the value of the StartTime element to the value of the StartTime element in the request and sets the value of the EndTime element to the rounded current time plus 30 minutes.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R42911");

        // Verify MS-ASCAL requirement: MS-ASCAL_R42911
        Site.CaptureRequirementIfAreEqual<string>(
            PastTime.ToString("yyyyMMddTHHmmssZ"),
            calendarWithPastStartTime.Calendar.StartTime.Value.ToUniversalTime().ToString("yyyyMMddTHHmmssZ"),
            42911,
            @"[In StartTime] As a top-level element of the Calendar class, the StartTime element specifies the start time of the calendar item.");
    }

    #endregion

    #region MSASCAL_S01_TC09_StartTimeFutureEndTimeAbsent

    /// <summary>
    /// This test case is designed to verify a calendar class with a StartTime element set as future time via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC09_StartTimeFutureEndTimeAbsent()
    {
        #region Call Sync command to add a calendar with the element StartTime setting as future time to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
        {
            {
                Request.ItemsChoiceType8.EndTime, null
            }
        };

        calendarItem.Add(Request.ItemsChoiceType8.StartTime, FutureTime.ToString("yyyyMMddTHHmmssZ"));

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52518");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52518
        Site.CaptureRequirementIfAreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            52518,
            @"[In Creating Calendar Events when the StartTime Element or EndTime Element is Absent] If StartTime is in the future and EndTime is absent the server includes a Status element with a value of 6 in the response, indicating an error occurred.");
    }

    #endregion

    #region MSASCAL_S01_TC10_Categories

    /// <summary>
    /// This test case is designed to verify a calendar class with a Categories element via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC10_Categories()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar with the element Categories and one sub-element Category to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var subjectWithCategoriesLessThan300 = Common.GenerateResourceName(Site, "subjectWithCategoriesLessThan300");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithCategoriesLessThan300);

        // Set Calendar StartTime, EndTime elements
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, EndTime.ToString("yyyyMMddTHHmmssZ"));

        // Set Categories element specifies a category that is assigned to the calendar item
        calendarItem.Add(Request.ItemsChoiceType8.Categories, TestSuiteHelper.CreateCalendarCategories(new string[] { Category }));

        // Set Categories element specifies a category that is assigned to the exception item
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(0, 6, 1));

        var categoryNameInException = Category + "InException";

        var exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
        var exceptionList = new List<Request.ExceptionsException>();

        var exceptionWithCategoriesLessThan300 = TestSuiteHelper.CreateExceptionRequired(StartTime.AddDays(2).ToString("yyyyMMddTHHmmssZ"));
        exceptionWithCategoriesLessThan300.Categories = TestSuiteHelper.CreateCalendarCategories(new string[] { categoryNameInException }).Category;

        exceptionList.Add(exceptionWithCategoriesLessThan300);
        exceptions.Exception = exceptionList.ToArray();
        calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);

        AddSyncCalendar(calendarItem);

        var calendarWithCategoriesLessThan300 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithCategoriesLessThan300);

        Site.Assert.IsNotNull(calendarWithCategoriesLessThan300.Calendar, "The calendar with subject {0} should exist in server.", subjectWithCategoriesLessThan300);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithCategoriesLessThan300);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2083");

        // Verify MS-ASCAL requirement: MS-ASCAL_R2083
        // If the DtStamp element is not specified as a child element of an Exception element, the value of the DtStamp element is assumed to be the
        // same as the value of the top-level DtStamp element. So this requirement can be covered if DtStamp for the calendar item is returned.
        Site.CaptureRequirementIfIsNotNull(
            calendarWithCategoriesLessThan300.Calendar.DtStamp.Value,
            2083,
            @"[In DtStamp] As a top-level element of the Calendar class, the DtStamp element specifies [the date and time that the calendar item was created or modified or] the date and time at which the exception item was created or modified..");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R17711");

        // Verify MS-ASCAL requirement: MS-ASCAL_R17711
        Site.CaptureRequirementIfAreEqual<string>(
            Category,
            calendarWithCategoriesLessThan300.Calendar.Categories.Category[0],
            17711,
            @"[In Categories] The Categories element specifies a collection of categories assigned to the calendar item.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R17911");

        // Verify MS-ASCAL requirement: MS-ASCAL_R17911
        Site.CaptureRequirementIfAreEqual<string>(
            categoryNameInException,
            calendarWithCategoriesLessThan300.Calendar.Exceptions.Exception[0].Categories[0],
            17911,
            @"[In Categories] As a child element of the Exception element (section 2.2.2.19), the Categories element specifies the categories for the exception item.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R18011");

        // Verify MS-ASCAL requirement: MS-ASCAL_R18011
        Site.CaptureRequirementIfAreEqual<string>(
            categoryNameInException,
            calendarWithCategoriesLessThan300.Calendar.Exceptions.Exception[0].Categories[0],
            18011,
            @"[In Categories] A command response has a maximum of one Categories child element per Exception element.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R18311");

        // Verify MS-ASCAL requirement: MS-ASCAL_R18311
        Site.CaptureRequirementIfAreEqual<string>(
            Category,
            calendarWithCategoriesLessThan300.Calendar.Categories.Category[0],
            18311,
            @"[In Category] The Category element specifies a category that is assigned to the calendar item.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R18312");

        // Verify MS-ASCAL requirement: MS-ASCAL_R18312
        Site.CaptureRequirementIfAreEqual<string>(
            categoryNameInException,
            calendarWithCategoriesLessThan300.Calendar.Exceptions.Exception[0].Categories[0],
            18312,
            @"[In Category] The Category element specifies a category that is assigned to the exception item.");

        if (Common.IsRequirementEnabled(11026, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R11026");

            // Verify MS-ASCAL requirement: MS-ASCAL_R11026
            Site.CaptureRequirementIfIsTrue(
                calendarWithCategoriesLessThan300.Calendar.Exceptions.Exception[0].Categories.Length >= 0 && calendarWithCategoriesLessThan300.Calendar.Categories.Category.Length <= 300,
                11026,
                @"[In Appendix B: Product Behavior] Implementation command response includes no more than 300 Category child elements per Categories element. (Exchange 2007 SP1 and above follow this behavior.)");
        }

        #region Call Sync command to add a calendar with the element Categories and more than 300 sub-element Category to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set more than 300 sub-element Category
        var categoryList = new List<string>();
        for (var i = 0; i <= 301; i++)
        {
            categoryList.Add(Category);
        }

        // Set Calendar StartTime, EndTime elements
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, EndTime.ToString("yyyyMMddTHHmmssZ"));

        // Set Categories element specifies a category that is assigned to the calendar item
        calendarItem.Add(Request.ItemsChoiceType8.Categories, TestSuiteHelper.CreateCalendarCategories(categoryList.ToArray()));

        // Set Categories element specifies a category that is assigned to the exception item
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(0, 6, 1));

        exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
        exceptionList = new List<Request.ExceptionsException>();

        var exceptionWithCategoriesMoreThan300 = TestSuiteHelper.CreateExceptionRequired(StartTime.AddDays(2).ToString("yyyyMMddTHHmmssZ"));
        exceptionWithCategoriesMoreThan300.Categories = TestSuiteHelper.CreateCalendarCategories(categoryList.ToArray()).Category;

        exceptionList.Add(exceptionWithCategoriesMoreThan300);
        exceptions.Exception = exceptionList.ToArray();
        calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.IsFalse(addCalendarResponse.AddResponses[0].Status.Equals(1), "Command request can not includes more than 300 Category child elements per Categories element.");

        #endregion
    }

    #endregion

    #region MSASCAL_S01_TC11_RecurrenceWithType0

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Type set as recurs daily via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC11_RecurrenceWithType0()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
           
        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("0");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var calendars = new List<SyncItem>();

        var subjectWithType0AndOccurrences = Common.GenerateResourceName(Site, "subjectWithType0AndOccurrences");
        var subjectWithType0AndUntil = Common.GenerateResourceName(Site, "subjectWithType0AndUntil");
        var subjectWithType0Only = Common.GenerateResourceName(Site, "subjectWithType0Only");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '0' and Occurrences sub-element to the server, and sync calendars from the server.

        // Add Calendar Recurrence element including Occurrences sub-element
        var occurrences = 3;
        var interval = 2;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(recurrenceType, occurrences, interval));
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType0AndOccurrences);

        AddSyncCalendar(calendarItem);
        var calendarWithType0AndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType0AndOccurrences);

        Site.Assert.IsNotNull(calendarWithType0AndOccurrences.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType0AndOccurrences);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType0AndOccurrences);
        calendars.Add(calendarWithType0AndOccurrences);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '0' and Until sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Until sub-element
        var untilTime = FutureTime.AddDays(2).ToString("yyyyMMddTHHmmssZ");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 0, interval);
        recurrence.Until = untilTime;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType0AndUntil);

        AddSyncCalendar(calendarItem);
        var calendarWithType0AndUntil = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType0AndUntil);

        Site.Assert.IsNotNull(calendarWithType0AndUntil.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType0AndUntil);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType0AndUntil);
        calendars.Add(calendarWithType0AndUntil);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence only including Type '0' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without Occurrences and Until sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence.OccurrencesSpecified = false;
        recurrence.Until = null;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType0Only);

        AddSyncCalendar(calendarItem);
        var calendarWithType0Only = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType0Only);

        Site.Assert.IsNotNull(calendarWithType0Only.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType0Only);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType0Only);
        calendars.Add(calendarWithType0Only);

        #endregion

        foreach (var response in calendars)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R37211");

            // Verify MS-ASCAL requirement: MS-ASCAL_R37211
            // If Calendar.Recurrence is not null, it means the element Recurrence is returned in response
            Site.CaptureRequirementIfIsNotNull(
                response.Calendar.Recurrence,
                37211,
                @"[In Recurrence] The Recurrence element specifies the recurrence pattern for the calendar item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19511");

            // Verify MS-ASCAL requirement: MS-ASCAL_R19511
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.DayOfMonthSpecified,
                19511,
                @"[In DayOfMonth] The DayOfMonth element MUST NOT be included in responses when the Type element value is zero (0).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32811");

            // Verify MS-ASCAL requirement: MS-ASCAL_R32811
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.MonthOfYearSpecified,
                32811,
                @"[In MonthOfYear] The MonthOfYear element MUST NOT be included in responses when the Type element value is zero (0).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R454");

            // Verify MS-ASCAL requirement: MS-ASCAL_R454
            Site.CaptureRequirementIfAreEqual<byte>(
                recurrenceType,
                response.Calendar.Recurrence.Type,
                454,
                @"[In Type] [The value 0 means] Recurs daily.");
        }
    }

    #endregion

    #region MSASCAL_S01_TC12_RecurrenceWithType1

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Type set as recurs weekly via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC12_RecurrenceWithType1()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
           
        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("1");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var calendars = new List<SyncItem>();

        var subjectWithType1AndOccurrences = Common.GenerateResourceName(Site, "subjectWithType1AndOccurrences");
        var subjectWithType1AndUntil = Common.GenerateResourceName(Site, "subjectWithType1AndUntil");
        var subjectWithType1Only = Common.GenerateResourceName(Site, "subjectWithType1Only");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '1' and Occurrences sub-element to the server, and sync calendars from the server.

        // Add Calendar Recurrence element including Occurrences sub-element
        var occurrences = 4;
        var interval = 2;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(recurrenceType, occurrences, interval));
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType1AndOccurrences);

        AddSyncCalendar(calendarItem);
        var calendarWithType1AndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType1AndOccurrences);

        Site.Assert.IsNotNull(calendarWithType1AndOccurrences.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType1AndOccurrences);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType1AndOccurrences);
        calendars.Add(calendarWithType1AndOccurrences);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '1' and Until sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Until sub-element
        var untilTime = FutureTime.AddDays(14).ToString("yyyyMMddTHHmmssZ");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 0, interval);
        recurrence.Until = untilTime;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType1AndUntil);

        AddSyncCalendar(calendarItem);
        var calendarWithType1AndUntil = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType1AndUntil);

        Site.Assert.IsNotNull(calendarWithType1AndUntil.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType1AndUntil);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType1AndUntil);
        calendars.Add(calendarWithType1AndUntil);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence only including Type '1' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without Occurrences and Until sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence.OccurrencesSpecified = false;
        recurrence.Until = null;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType1Only);

        AddSyncCalendar(calendarItem);
        var calendarWithType1Only = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType1Only);

        Site.Assert.IsNotNull(calendarWithType1Only.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType1Only);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType1Only);
        calendars.Add(calendarWithType1Only);

        #endregion

        foreach (var response in calendars)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19512");

            // Verify MS-ASCAL requirement: MS-ASCAL_R19512
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.DayOfMonthSpecified,
                19512,
                @"[In DayOfMonth] The DayOfMonth element MUST NOT be included in responses when the Type element value is 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32812");

            // Verify MS-ASCAL requirement: MS-ASCAL_R32812
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.MonthOfYearSpecified,
                32812,
                @"[In MonthOfYear] The MonthOfYear element MUST NOT be included in responses when the Type element value is 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R455");

            // Verify MS-ASCAL requirement: MS-ASCAL_R455
            Site.CaptureRequirementIfAreEqual<byte>(
                recurrenceType,
                response.Calendar.Recurrence.Type,
                455,
                @"[In Type] [The value 1 means]  Recurs weekly.");
        }
    }

    #endregion

    #region MSASCAL_S01_TC13_RecurrenceWithType2

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Type set as recurs monthly via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC13_RecurrenceWithType2()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("2");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var calendars = new List<SyncItem>();

        var subjectWithType2AndOccurrences = Common.GenerateResourceName(Site, "subjectWithType2AndOccurrences");
        var subjectWithType2AndUntil = Common.GenerateResourceName(Site, "subjectWithType2AndUntil");
        var subjectWithType2Only = Common.GenerateResourceName(Site, "subjectWithType2Only");
        var subjectWithType2AndCalendarType1 = Common.GenerateResourceName(Site, "subjectWithType2AndCalendarType1");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '2' and Occurrences sub-element to the server, and sync calendars from the server.

        // Add Calendar Recurrence element including Occurrences sub-element
        var occurrences = 5;
        var interval = 2;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(recurrenceType, occurrences, interval));
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType2AndOccurrences);

        AddSyncCalendar(calendarItem);
        var calendarWithType2AndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType2AndOccurrences);

        Site.Assert.IsNotNull(calendarWithType2AndOccurrences.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType2AndOccurrences);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType2AndOccurrences);
        calendars.Add(calendarWithType2AndOccurrences);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '2' and Until sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Until sub-element
        var untilTime = FutureTime.AddMonths(2).ToString("yyyyMMddTHHmmssZ");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 0, interval);
        recurrence.Until = untilTime;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType2AndUntil);

        AddSyncCalendar(calendarItem);
        var calendarWithType2AndUntil = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType2AndUntil);

        Site.Assert.IsNotNull(calendarWithType2AndUntil.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType2AndUntil);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType2AndUntil);
        calendars.Add(calendarWithType2AndUntil);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence only including Type '2' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without Occurrences and Until sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence.OccurrencesSpecified = false;
        recurrence.Until = null;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType2Only);

        AddSyncCalendar(calendarItem);
        var calendarWithType2Only = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType2Only);

        Site.Assert.IsNotNull(calendarWithType2Only.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType2Only);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType2Only);
        calendars.Add(calendarWithType2Only);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '2' and CalendarType sub-element setting as "1" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "1".
        var calendarType = byte.Parse("1");
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), calendarType);
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType2AndCalendarType1);

        AddSyncCalendar(calendarItem);
        var calendarWithType2AndCalendarType1 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType2AndCalendarType1);

        Site.Assert.IsNotNull(calendarWithType2AndCalendarType1.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType2AndCalendarType1);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType2AndCalendarType1);
        calendars.Add(calendarWithType2AndCalendarType1);

        #endregion

        foreach (var response in calendars)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19011");

            // Verify MS-ASCAL requirement: MS-ASCAL_R19011
            Site.CaptureRequirementIfIsTrue(
                response.Calendar.Recurrence.DayOfMonthSpecified,
                19011,
                @"[In DayOfMonth] A command response has a minimum of one DayOfMonth child element per Recurrence element when the value of the Type element (section 2.2.2.43) is 2[or 5].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R456");

            // Verify MS-ASCAL requirement: MS-ASCAL_R456
            Site.CaptureRequirementIfAreEqual<byte>(
                recurrenceType,
                response.Calendar.Recurrence.Type,
                456,
                @"[In Type] [The value 2 means]  Recurs monthly.");

            if (!IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R143");

                // Verify MS-ASCAL requirement: MS-ASCAL_R143
                Site.CaptureRequirementIfIsTrue(
                    response.Calendar.Recurrence.CalendarTypeSpecified,
                    143,
                    @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element value is 2.");
            }

            Site.Assert.IsFalse(response.Calendar.Recurrence.MonthOfYearSpecified, "The MonthOfYear element MUST NOT be included in responses when the Type element value is 2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32813");

            // Verify MS-ASCAL requirement: MS-ASCAL_R32813
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.MonthOfYearSpecified,
                32813,
                @"[In MonthOfYear] The MonthOfYear element MUST NOT be included in responses when the Type element value is 2.");
        }
    }

    #endregion

    #region MSASCAL_S01_TC14_RecurrenceWithType3

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Type set as recurs monthly on the nth day via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC14_RecurrenceWithType3()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("3");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var calendars = new List<SyncItem>();

        var subjectWithType3AndOccurrences = Common.GenerateResourceName(Site, "subjectWithType3AndOccurrences");
        var subjectWithType3AndUntil = Common.GenerateResourceName(Site, "subjectWithType3AndUntil");
        var subjectWithType3Only = Common.GenerateResourceName(Site, "subjectWithType3Only");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '3' and Occurrences sub-element to the server, and sync calendars from the server.

        // Add Calendar Recurrence element including Occurrences sub-element
        var occurrences = 6;
        var interval = 2;
        var recurrence1 = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence1.DayOfWeek = 62;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence1);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType3AndOccurrences);

        AddSyncCalendar(calendarItem);
        var calendarWithType3AndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType3AndOccurrences);

        Site.Assert.IsNotNull(calendarWithType3AndOccurrences.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType3AndOccurrences);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType3AndOccurrences);
        calendars.Add(calendarWithType3AndOccurrences);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '3' and Until sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Until sub-element
        var untilTime = FutureTime.AddMonths(2).ToString("yyyyMMddTHHmmssZ");
        var recurrence2 = CreateCalendarRecurrence(recurrenceType, 0, interval);
        recurrence2.Until = untilTime;
        recurrence2.DayOfWeek = 65;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence2);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType3AndUntil);

        AddSyncCalendar(calendarItem);
        var calendarWithType3AndUntil = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType3AndUntil);

        Site.Assert.IsNotNull(calendarWithType3AndUntil.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType3AndUntil);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType3AndUntil);
        calendars.Add(calendarWithType3AndUntil);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence only including Type '3' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without Occurrences and Until sub-element
        var recurrence3 = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence3.OccurrencesSpecified = false;
        recurrence3.Until = null;
        recurrence3.DayOfWeek = 127;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence3);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType3Only);

        AddSyncCalendar(calendarItem);
        var calendarWithType3Only = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType3Only);

        Site.Assert.IsNotNull(calendarWithType3Only.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType3Only);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType3Only);
        calendars.Add(calendarWithType3Only);

        #endregion

        foreach (var response in calendars)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19513");

            // Verify MS-ASCAL requirement: MS-ASCAL_R19513
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.DayOfMonthSpecified,
                19513,
                @"[In DayOfMonth] The DayOfMonth element MUST NOT be included in responses when the Type element value is 3.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32814");

            // Verify MS-ASCAL requirement: MS-ASCAL_R32814
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.MonthOfYearSpecified,
                32814,
                @"[In MonthOfYear] The MonthOfYear element MUST NOT be included in responses when the Type element value is 3.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R457");

            // Verify MS-ASCAL requirement: MS-ASCAL_R457
            Site.CaptureRequirementIfAreEqual<byte>(
                recurrenceType,
                response.Calendar.Recurrence.Type,
                457,
                @"[In Type] [The value 3 means]  Recurs monthly on the nth day.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R47311");

            // Verify MS-ASCAL requirement: MS-ASCAL_R47311
            Site.CaptureRequirementIfIsTrue(
                response.Calendar.Recurrence.WeekOfMonthSpecified,
                47311,
                @"[In WeekOfMonth] A command response has a minimum of one WeekOfMonth child element per Recurrence element when the value of the Type element (section 2.2.2.43) is 3.");

            if (!IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R14311");

                // Verify MS-ASCAL requirement: MS-ASCAL_R14311
                Site.CaptureRequirementIfIsTrue(
                    response.Calendar.Recurrence.CalendarTypeSpecified,
                    14311,
                    @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element value is 3.");
            }
        }
    }

    #endregion

    #region MSASCAL_S01_TC15_RecurrenceWithType5

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Type set as recurs yearly via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC15_RecurrenceWithType5()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("5");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var calendars = new List<SyncItem>();

        var subjectWithType5AndOccurrences = Common.GenerateResourceName(Site, "subjectWithType5AndOccurrences");
        var subjectWithType5AndUntil = Common.GenerateResourceName(Site, "subjectWithType5AndUntil");
        var subjectWithType5Only = Common.GenerateResourceName(Site, "subjectWithType5Only");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '5' and Occurrences sub-element to the server, and sync calendars from the server.

        // Add Calendar Recurrence element including Occurrences sub-element
        var occurrences = 7;
        var interval = 2;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(recurrenceType, occurrences, interval));
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType5AndOccurrences);

        AddSyncCalendar(calendarItem);
        var calendarWithType5AndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType5AndOccurrences);

        Site.Assert.IsNotNull(calendarWithType5AndOccurrences.Calendar, "The calendar with subject name {0} should be found in server.", subjectWithType5AndOccurrences);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType5AndOccurrences);

        calendars.Add(calendarWithType5AndOccurrences);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '5' and Until sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Until sub-element
        var recurrence = CreateCalendarRecurrence(recurrenceType, 0, interval);
        recurrence.Until = FutureTime.AddYears(5).ToString("yyyyMMddTHHmmssZ");
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType5AndUntil);

        AddSyncCalendar(calendarItem);
        var calendarWithType5AndUntil = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType5AndUntil);

        Site.Assert.IsNotNull(calendarWithType5AndUntil.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType5AndUntil);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType5AndUntil);
        calendars.Add(calendarWithType5AndUntil);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence only including Type '5' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without Occurrences and Until sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, 1, interval);
        recurrence.OccurrencesSpecified = false;
        recurrence.Until = null;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType5Only);

        AddSyncCalendar(calendarItem);
        var calendarWithType5Only = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType5Only);

        Site.Assert.IsNotNull(calendarWithType5Only.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType5Only);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType5Only);
        calendars.Add(calendarWithType5Only);

        #endregion

        foreach (var response in calendars)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19012");

            // Verify MS-ASCAL requirement: MS-ASCAL_R19012
            Site.CaptureRequirementIfIsTrue(
                response.Calendar.Recurrence.DayOfMonthSpecified,
                19012,
                @"[In DayOfMonth] A command response has a minimum of one DayOfMonth child element per Recurrence element when the value of the Type element (section 2.2.2.43) is [2 or] 5.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32311");

            // Verify MS-ASCAL requirement: MS-ASCAL_R32311
            Site.CaptureRequirementIfIsTrue(
                response.Calendar.Recurrence.MonthOfYearSpecified,
                32311,
                @"[In MonthOfYear] A command response has a minimum of one MonthOfYear child element per Recurrence element if the value of the Type element (section 2.2.2.43) is 5.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R458");

            // Verify MS-ASCAL requirement: MS-ASCAL_R458
            Site.CaptureRequirementIfAreEqual<byte>(
                recurrenceType,
                response.Calendar.Recurrence.Type,
                458,
                @"[In Type] [The value 5 means]  Recurs yearly.");

            if (!IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R14312");

                // Verify MS-ASCAL requirement: MS-ASCAL_R14312
                Site.CaptureRequirementIfIsTrue(
                    response.Calendar.Recurrence.CalendarTypeSpecified,
                    14312,
                    @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element value is 5.");
            }
        }
    }

    #endregion

    #region MSASCAL_S01_TC16_RecurrenceWithType6

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Type set as recurs yearly on the nth day via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC16_RecurrenceWithType6()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("6");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var calendars = new List<SyncItem>();

        var subjectWithType6AndOccurrences = Common.GenerateResourceName(Site, "subjectWithType6AndOccurrences");
        var subjectWithType6AndUntil = Common.GenerateResourceName(Site, "subjectWithType6AndUntil");
        var subjectWithType6Only = Common.GenerateResourceName(Site, "subjectWithType6Only");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '6' and Occurrences sub-element to the server, and sync calendars from the server.

        // Add Calendar Recurrence element including Occurrences sub-element
        var occurrences = 3;
        var interval = 1;
        var recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence.WeekOfMonth = 5;
        recurrence.WeekOfMonthSpecified = true;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType6AndOccurrences);

        AddSyncCalendar(calendarItem);
        var calendarWithType6AndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType6AndOccurrences);

        Site.Assert.IsNotNull(calendarWithType6AndOccurrences.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType6AndOccurrences);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType6AndOccurrences);
        calendars.Add(calendarWithType6AndOccurrences);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '6' and Until sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Until sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, 0, interval);
        recurrence.Until = FutureTime.AddYears(5).ToString("yyyyMMddTHHmmssZ");
        recurrence.WeekOfMonth = 5;
        recurrence.WeekOfMonthSpecified = true;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType6AndUntil);

        AddSyncCalendar(calendarItem);
        var calendarWithType6AndUntil = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType6AndUntil);

        Site.Assert.IsNotNull(calendarWithType6AndUntil.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType6AndUntil);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType6AndUntil);
        calendars.Add(calendarWithType6AndUntil);

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence only including Type '6' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without Occurrences and Until sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, 1, interval);
        recurrence.OccurrencesSpecified = false;
        recurrence.Until = null;
        recurrence.WeekOfMonth = 5;
        recurrence.WeekOfMonthSpecified = true;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType6Only);

        AddSyncCalendar(calendarItem);
        var calendarWithType6Only = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType6Only);

        Site.Assert.IsNotNull(calendarWithType6Only.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType6Only);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType6Only);
        calendars.Add(calendarWithType6Only);

        #endregion

        foreach (var response in calendars)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R19514");

            // Verify MS-ASCAL requirement: MS-ASCAL_R19514
            Site.CaptureRequirementIfIsFalse(
                response.Calendar.Recurrence.DayOfMonthSpecified,
                19514,
                @"[In DayOfMonth] The DayOfMonth element MUST NOT be included in responses when the Type element value is 6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R32312");

            // Verify MS-ASCAL requirement: MS-ASCAL_R32312
            Site.CaptureRequirementIfIsTrue(
                response.Calendar.Recurrence.MonthOfYearSpecified,
                32312,
                @"[In MonthOfYear] A command response has a minimum of one MonthOfYear child element per Recurrence element if the value of the Type element (section 2.2.2.43) is 6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R47312");

            // Verify MS-ASCAL requirement: MS-ASCAL_R47312
            Site.CaptureRequirementIfIsTrue(
                response.Calendar.Recurrence.WeekOfMonthSpecified,
                47312,
                @"[In WeekOfMonth] A command response has a minimum of one WeekOfMonth child element per Recurrence element when the value of the Type element (section 2.2.2.43) is 6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R459");

            // Verify MS-ASCAL requirement: MS-ASCAL_R459
            Site.CaptureRequirementIfAreEqual<byte>(
                recurrenceType,
                response.Calendar.Recurrence.Type,
                459,
                @"[In Type] [The value 6 means]  Recurs yearly on the nth day.");

            if (!IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R14313");

                // Verify MS-ASCAL requirement: MS-ASCAL_R14313
                Site.CaptureRequirementIfIsTrue(
                    response.Calendar.Recurrence.CalendarTypeSpecified,
                    14313,
                    @"[In CalendarType] A command response has a minimum of one CalendarType child element per Recurrence element when the Type element value is 6.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R477");

            // Verify MS-ASCAL requirement: MS-ASCAL_R477
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                response.Calendar.Recurrence.WeekOfMonth,
                477,
                @"[In WeekOfMonth] The value of 5 specifies the last week of the month.");
        }
    }

    #endregion

    #region MSASCAL_S01_TC17_FirstDayOfWeek

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element including the element FirstDayOfWeek.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC17_FirstDayOfWeek()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subjectWithType1AndFirstDayOfWeek = Common.GenerateResourceName(Site, "subjectWithType1AndFirstDayOfWeek");
        var subjectWithoutType1AndFirstDayOfWeek = Common.GenerateResourceName(Site, "subjectWithoutType1AndFirstDayOfWeek");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '1' and FirstDayOfWeek sub-element to the server, and sync calendars from the server.

        var recurrenceType = byte.Parse("1");

        // Set Calendar Recurrence element including Type '1' and FirstDayOfWeek sub-element
        var firstDayofWeek = 4;
        var recurrence = CreateCalendarRecurrence(recurrenceType, 7, 5);
        recurrence.FirstDayOfWeek = byte.Parse(firstDayofWeek.ToString());
        recurrence.FirstDayOfWeekSpecified = true;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType1AndFirstDayOfWeek);

        AddSyncCalendar(calendarItem);
        var calendarWithType1AndFirstDayOfWeek = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType1AndFirstDayOfWeek);

        Site.Assert.IsNotNull(calendarWithType1AndFirstDayOfWeek.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType1AndFirstDayOfWeek);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType1AndFirstDayOfWeek);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R272");

        // Verify MS-ASCAL requirement: MS-ASCAL_R272
        Site.CaptureRequirementIfAreEqual<byte>(
            (byte)firstDayofWeek,
            calendarWithType1AndFirstDayOfWeek.Calendar.Recurrence.FirstDayOfWeek,
            272,
            @"[In FirstDayOfWeek] A command response has a maximum of one FirstDayOfWeek child element per Recurrence element.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52523");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52523
        Site.CaptureRequirementIfIsTrue(
            calendarWithType1AndFirstDayOfWeek.Calendar.Recurrence.Type == recurrenceType && calendarWithType1AndFirstDayOfWeek.Calendar.Recurrence.FirstDayOfWeekSpecified,
            52523,
            @"[In Message Processing Events and Sequencing Rules][The following information pertains to all command responses:] The server MUST return a FirstDayOfWeek element when the value of the Type element (section 2.2.2.43) is 1.");

        #region Call Sync command to add a calendar with the element Recurrence including Type '1' and without FirstDayOfWeek sub-element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Type '1' and FirstDayOfWeek sub-element
        recurrence = CreateCalendarRecurrence(recurrenceType, 7, 5);
        recurrence.FirstDayOfWeekSpecified = false;
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithoutType1AndFirstDayOfWeek);

        AddSyncCalendar(calendarItem);
        var calendarWithoutType1AndFirstDayOfWeek = GetChangeItem(User1Information.CalendarCollectionId, subjectWithoutType1AndFirstDayOfWeek);

        Site.Assert.IsNotNull(calendarWithoutType1AndFirstDayOfWeek.Calendar, "The calendar with subject {0} should exist in server.", subjectWithoutType1AndFirstDayOfWeek);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithoutType1AndFirstDayOfWeek);

        #endregion

        if (Common.IsRequirementEnabled(11028, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R11028");

            // Verify MS-ASCAL requirement: MS-ASCAL_R11028
            Site.CaptureRequirementIfIsTrue(
                calendarWithoutType1AndFirstDayOfWeek.Calendar.Recurrence.FirstDayOfWeekSpecified,
                11028,
                @"[In Appendix B: Product Behavior] The implementation identifies the first day of the week for any recurrence according to the preconfigured options of the user creating the calendar item, if the FirstDayOfWeek element is not included in the client request. (Exchange 2013 and above follow this behavior.)");
        }
    }

    #endregion

    #region MSASCAL_S01_TC18_WrongFormatEmailElement

    /// <summary>
    /// This test case is designed to verify wrong-format Email element related requirements.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC18_WrongFormatEmailElement()
    {
        #region Call Sync command to add a calendar with the element Attendees with wrong-formatted email address to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        // Set Calendar Attendee element with wrong-formatted email address
        var wrongFormatEmailAddress = "wrongFormatEmail";
        calendarItem.Add(Request.ItemsChoiceType8.Attendees, TestSuiteHelper.CreateAttendeesRequired(new string[] { wrongFormatEmailAddress }, new string[] { User2Information.UserName }));
        var subjectWithWrongEmailAddress = Common.GenerateResourceName(Site, "subjectWithWrongEmailAddress");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithWrongEmailAddress);

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        var calendarWithWrongEmailAddress = GetChangeItem(User1Information.CalendarCollectionId, subjectWithWrongEmailAddress);

        Site.Assert.IsNotNull(calendarWithWrongEmailAddress.Calendar, "The calendar with subject {0} should exist in server.", subjectWithWrongEmailAddress);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithWrongEmailAddress);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52521");

        // Verify MS-ASCAL requirement: MS-ASCAL_R52521
        Site.CaptureRequirementIfIsTrue(
            addCalendarResponse.CollectionStatus.Equals((byte)1),
            52521,
            @"[In Message Processing Events and Sequencing Rules][The following information pertains to all command responses:] A server MUST recognize when the value of the Email element is not formatted as specified in [MS-ASDTYPE] section 2.6.2");

        if (Common.IsRequirementEnabled(52529, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52529");

            // Verify MS-ASCAL requirement: MS-ASCAL_R52529
            Site.CaptureRequirementIfAreEqual<string>(
                "wrongFormatEmail",
                calendarWithWrongEmailAddress.Calendar.Attendees.Attendee[0].Email,
                52529,
                @"[In Appendix B: Product Behavior] The implementation does not replace it [Email element] with suitable placeholder text if the value of the Email element is not formatted as specified in [MS-ASDTYPE] section 2.6.2. (Exchange Server 2007 SP1 and above follow this behavior.)");
        }
    }

    #endregion

    #region MSASCAL_S01_TC19_OccurrencesAndUntilBothSet

    /// <summary>
    /// This test case is designed to verify Recurrence element when Occurrences and Until both are set.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC19_OccurrencesAndUntilBothSet()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subjectWithUntilAndOccurrences = Common.GenerateResourceName(Site, "subjectWithUntilAndOccurrences");
        var subjectWithOccurrences999 = Common.GenerateResourceName(Site, "subjectWithOccurrences999");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Until and Occurrences sub-element to the server, and sync calendars from the server.

        var recurrenceType = byte.Parse("0");

        // Set Calendar Recurrence element including Until and Occurrences sub-element.
        var occurrences = 5;
        var recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, 1);

        // Set Until element
        recurrence.Until = FutureTime.AddYears(2).ToString("yyyyMMddTHHmmssZ");
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithUntilAndOccurrences);

        AddSyncCalendar(calendarItem);

        var calendarWithUntilAndOccurrences = GetChangeItem(User1Information.CalendarCollectionId, subjectWithUntilAndOccurrences);

        Site.Assert.IsNotNull(calendarWithUntilAndOccurrences.Calendar, "The calendar with subject {0} should exist in server.", subjectWithUntilAndOccurrences);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithUntilAndOccurrences);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53911");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53911
        Site.CaptureRequirementIfIsTrue(
            calendarWithUntilAndOccurrences.Calendar.Recurrence.OccurrencesSpecified && calendarWithUntilAndOccurrences.Calendar.Recurrence.Occurrences == occurrences && calendarWithUntilAndOccurrences.Calendar.Recurrence.Until == null,
            53911,
            @"[In Sync Command Response] If both the Occurrences element (section 2.2.2.30) and the Until element (section 2.2.2.45) are included in a Sync command request, then the server MUST respect the value of the Occurrences element [and ignore the value of the Until element].");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R467");

        // Verify MS-ASCAL requirement: MS-ASCAL_R467
        Site.CaptureRequirementIfIsTrue(
            calendarWithUntilAndOccurrences.Calendar.Recurrence.OccurrencesSpecified && calendarWithUntilAndOccurrences.Calendar.Recurrence.Until == null,
            467,
            @"[In Until] The Until element and the Occurrences element (section 2.2.2.30) are mutually exclusive.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R341");

        // Verify MS-ASCAL requirement: MS-ASCAL_R341
        Site.CaptureRequirementIfIsTrue(
            calendarWithUntilAndOccurrences.Calendar.Recurrence.OccurrencesSpecified && calendarWithUntilAndOccurrences.Calendar.Recurrence.Until == null,
            341,
            @"[In Occurrences] The Occurrences element and the Until element (section 2.2.2.45) are mutually exclusive.");

        #region Call Sync command to add a calendar with the element Recurrence including Occurrences sub-element which is set as '999' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element including Occurrences sub-element which is set as '999'.
        recurrence = CreateCalendarRecurrence(recurrenceType, 999, 1);
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithOccurrences999);

        AddSyncCalendar(calendarItem);
        var calendarWithOccurrences999 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithOccurrences999);

        Site.Assert.IsNotNull(calendarWithOccurrences999.Calendar, "The calendar with subject {0} should exist in server.", subjectWithOccurrences999);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithOccurrences999);

        #endregion

        calendarItem.Clear();

        #region Call Sync command to add a calendar with the element Recurrence including Occurrences sub-element which is set as more than '999' to the server, and sync calendars from the server.

        // Set Calendar Recurrence element including Occurrences sub-element which is set as more than '999'.
        recurrence = CreateCalendarRecurrence(recurrenceType, 1000, 1);

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R345");

        // Verify MS-ASCAL requirement: MS-ASCAL_R341
        Site.CaptureRequirementIfIsTrue(
            calendarWithOccurrences999.Calendar.Recurrence.OccurrencesSpecified
            && calendarWithOccurrences999.Calendar.Recurrence.Occurrences <= 999
            && addCalendarResponse.AddResponses[0].Status.Equals("6"),
            345,
            @"[In Occurrences] The maximum value is 999.");
    }

    #endregion

    #region MSASCAL_S01_TC20_IsLeapMonth

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element including the element IsLeapMonth.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC20_IsLeapMonth()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The IsLeapMonth element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subjectWithIsLeapMonth0 = Common.GenerateResourceName(Site, "subjectWithIsLeapMonth0");
        var subjectWithCalendarTypeWithoutIsLeapMonth = Common.GenerateResourceName(Site, "subjectWithCalendarTypeWithoutIsLeapMonth");
        var subjectWithCalendarTypeAndIsLeapMonth = Common.GenerateResourceName(Site, "subjectWithCalendarTypeAndIsLeapMonth");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including IsLeapMonth sub-element is '0' to the server, and sync calendars from the server.

        var occurrences = 5;
        var interval = 1;

        // Set Calendar Recurrence element including IsLeapMonth sub-element
        var recurrenceType = byte.Parse("5");
        var recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);

        // IsLeapMonth is set to 0, the recurrence of the appointment doesn't takes place on the embolismic (leap) month
        recurrence.IsLeapMonth = 0;
        recurrence.IsLeapMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithIsLeapMonth0);

        AddSyncCalendar(calendarItem);
        var calendarWithIsLeapMonth0 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithIsLeapMonth0);

        Site.Assert.IsNotNull(calendarWithIsLeapMonth0.Calendar, "The calendar with subject {0} should exist in server.", subjectWithIsLeapMonth0);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithIsLeapMonth0);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R608");

        // Verify MS-ASCAL requirement: MS-ASCAL_R608
        Site.CaptureRequirementIfAreEqual<byte>(
            (byte)0,
            calendarWithIsLeapMonth0.Calendar.Recurrence.IsLeapMonth,
            608,
            @"[In IsLeapMonth] [The value] 0 [means] False.");

        #region Call Sync command to add a calendar with the element Recurrence without IsLeapMonth but including the CalendarType element to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without IsLeapMonth sub-element
        recurrenceType = byte.Parse("5");
        recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);

        // CalendarType is set to 15, Chinese Lunar calendar system is used by the recurrence
        recurrence.CalendarTypeSpecified = true;
        recurrence.CalendarType = 15;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithCalendarTypeWithoutIsLeapMonth);

        AddSyncCalendar(calendarItem);
        var calendarWithCalendarTypeWithoutIsLeapMonth = GetChangeItem(User1Information.CalendarCollectionId, subjectWithCalendarTypeWithoutIsLeapMonth);

        Site.Assert.IsNotNull(calendarWithCalendarTypeWithoutIsLeapMonth.Calendar, "The calendar with subject {0} should exist in server.", subjectWithCalendarTypeWithoutIsLeapMonth);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithCalendarTypeWithoutIsLeapMonth);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R28411");

        // Verify MS-ASCAL requirement: MS-ASCAL_R28411
        // If Sync response can get a non-null IsLeapMonth, then we can capture this requirement
        Site.CaptureRequirementIfIsNotNull(
            calendarWithCalendarTypeWithoutIsLeapMonth.Calendar.Recurrence.IsLeapMonth,
            28411,
            @"[In IsLeapMonth] The IsLeapMonth element<11> specifies whether the recurrence of the appointment takes place on the embolismic (leap) month.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R288");

        // Verify MS-ASCAL requirement: MS-ASCAL_R288
        // If Sync response can get a non-null IsLeapMonth, then we can capture this requirement
        Site.CaptureRequirementIfIsNotNull(
            calendarWithCalendarTypeWithoutIsLeapMonth.Calendar.Recurrence.IsLeapMonth,
            288,
            @"[In IsLeapMonth] This element only applies when the CalendarType element (section 2.2.2.9) specifies a calendar system that incorporates an embolismic (leap) month.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R293");

        // Verify MS-ASCAL requirement: MS-ASCAL_R293
        Site.CaptureRequirementIfAreEqual<byte>(
            (byte)0,
            calendarWithCalendarTypeWithoutIsLeapMonth.Calendar.Recurrence.IsLeapMonth,
            293,
            @"[In IsLeapMonth] The default value of the IsLeapMonth element is 0 (FALSE).");

        #region Call Sync command to add a calendar with the element Recurrence including IsLeapMonth sub-element setting as '1' and CalendarType sub-element setting as "Gregorian" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without IsLeapMonth sub-element
        recurrenceType = byte.Parse("5");
        recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence.MonthOfYear = byte.Parse("6");

        // CalendarType set to "Gregorian"
        recurrence.CalendarTypeSpecified = true;
        recurrence.CalendarType = 1;

        // IsLeapMonth is set to 1, the recurrence of the appointment takes place on the embolismic (leap) month
        recurrence.IsLeapMonth = 1;
        recurrence.IsLeapMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithCalendarTypeAndIsLeapMonth);
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, new DateTime(2017, 1, 1, 1, 0, 0).ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, new DateTime(2017, 1, 1, 2, 0, 0).ToString("yyyyMMddTHHmmssZ"));

        AddSyncCalendar(calendarItem);
        var calendarWithCalendarTypeAndIsLeapMonth = GetChangeItem(User1Information.CalendarCollectionId, subjectWithCalendarTypeAndIsLeapMonth);

        Site.Assert.IsNotNull(calendarWithCalendarTypeAndIsLeapMonth.Calendar, "The calendar with subject {0} should exist in server.", subjectWithCalendarTypeAndIsLeapMonth);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithCalendarTypeAndIsLeapMonth);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R290");

        // Verify MS-ASCAL requirement: MS-ASCAL_R290
        Site.CaptureRequirementIfIsTrue(
            calendarWithCalendarTypeAndIsLeapMonth.Calendar.Recurrence.IsLeapMonthSpecified && calendarWithCalendarTypeAndIsLeapMonth.Calendar.Recurrence.IsLeapMonth != 1,
            290,
            @"[In IsLeapMonth] This element[IsLeapMonth] has no effect when specified in conjunction with the Gregorian calendar.");

        #region Call Sync command to add a calendar with the element Recurrence including IsLeapMonth sub-element setting as '1' and CalendarType sub-element setting as "Chinese Lunar" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element without IsLeapMonth sub-element
        recurrenceType = byte.Parse("5");
        recurrence = CreateCalendarRecurrence(recurrenceType, occurrences, interval);
        recurrence.MonthOfYear = byte.Parse("6");

        // CalendarType set to "Chinese Lunar"
        recurrence.CalendarTypeSpecified = true;
        recurrence.CalendarType = 15;

        // IsLeapMonth is set to 1, the recurrence of the appointment takes place on the embolismic (leap) month
        recurrence.IsLeapMonth = 1;
        recurrence.IsLeapMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, new DateTime(2017, 1, 1, 1, 0, 0).ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, new DateTime(2017, 1, 1, 2, 0, 0).ToString("yyyyMMddTHHmmssZ"));

        AddSyncCalendar(calendarItem);
        var calendar = GetChangeItem(User1Information.CalendarCollectionId, SubjectName);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", SubjectName);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, SubjectName);
        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R609");

        // Verify MS-ASCAL requirement: MS-ASCAL_R609
        Site.CaptureRequirementIfIsTrue(
            calendar.Calendar.Recurrence.IsLeapMonthSpecified && calendar.Calendar.Recurrence.IsLeapMonth == 1,
            609,
            @"[In IsLeapMonth] [The value] 1 [means]True.");
    }

    #endregion

    #region MSASCAL_S01_TC21_Status6WithMultiCalendarType

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including more than one CalendarType elements.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC21_Status6WithMultiCalendarType()
    {
        #region Create calendars with different Type element value.

        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        // Add a calendar with more than one CalendarType when type is 2.
        byte type = 2;
        var syncAddResponse1 = AddCalendarWithMultipleCalendarType(type.ToString());

        // Add a calendar with more than one CalendarType when type is 3.
        type = 3;
        var syncAddResponse2 = AddCalendarWithMultipleCalendarType(type.ToString());

        // Add a calendar with more than one CalendarType when type is 5.
        type = 5;
        var syncAddResponse3 = AddCalendarWithMultipleCalendarType(type.ToString());

        // Add a calendar with more than one CalendarType when type is 6.
        type = 6;
        var syncAddResponse4 = AddCalendarWithMultipleCalendarType(type.ToString());

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53912");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53912
        Site.CaptureRequirementIfIsTrue(
            syncAddResponse1.AddResponses[0].Status == "6" && syncAddResponse2.AddResponses[0].Status == "6" && syncAddResponse3.AddResponses[0].Status == "6" && syncAddResponse4.AddResponses[0].Status == "6",
            53912,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] A command request has more than one CalendarType element (section 2.2.2.9) per Recurrence element (section 2.2.2.35) when the Type element (section 2.2.2.43) value is 2, 3, 5, or 6.");
    }

    #endregion

    #region MSASCAL_S01_TC22_Status6WithSpecifiedCalendarType

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including CalendarType with specified value.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC22_Status6WithSpecifiedCalendarType()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Define common variables.

        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The CalendarType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
        var recurrenceType = byte.Parse("2");
        var occurrences = 3;
        var interval = 3;

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "13" to the server, and sync calendars from the server.

        // Set Calendar Recurrence element, CalendarType is set to "13".
        var recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("13"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 13.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "16" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "16".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("16"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 16.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "17" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "17".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("17"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 17.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "18" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "18".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("18"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 18.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "19" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "19".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("19"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 19.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "21" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "21".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("21"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 21.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "22" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "22".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("22"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 22.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including CalendarType sub-element which is set as "23" to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element, CalendarType is set to "23".
        recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), byte.Parse("23"));
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        addCalendarResponse = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the CalendarType element is set to 23.");

        #endregion

        // According to above steps, requirement MS-ASCAL_R53913 can be covered directly.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53913");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53913
        Site.CaptureRequirement(
            53913,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The CalendarType element is set to one of the following values in the request: 13, 16, 17, 18, 19, 21, 22, or 23.");
    }

    #endregion

    #region MSASCAL_S01_TC23_Status6WithOutsideRangeFirstDayOfWeek

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including a FirstDayOfWeek element with out-of-ranged.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC23_Status6WithOutsideRangeFirstDayOfWeek()
    {
        #region Call Sync command to add a calendar with the element Recurrence including FirstDayOfWeek sub-element out of range to the server, and sync calendars from the server.

        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The FirstDayOfWeek element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        // Set Calendar Recurrence element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set FirstDayOfWeek outside the range 0 (zero) through 6 (inclusive)
        recurrence.FirstDayOfWeek = 7;
        recurrence.FirstDayOfWeekSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse = AddSyncCalendar(calendarItem);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53914");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53914
        Site.CaptureRequirementIfAreEqual<string>(
            "6",
            addCalendarResponse.AddResponses[0].Status,
            53914,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The value of the FirstDayOfWeek element (section 2.2.2.22) is outside the range 0 (zero) through 6 (inclusive).");
    }

    #endregion

    #region MSASCAL_S01_TC24_Status6WithDayOfMonth

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including a DayOfMonth element.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC24_Status6WithDayOfMonth()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar with the element Recurrence including DayOfMonth sub-element when Type is '0' to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        // Set Calendar Recurrence element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set DayOfMonth
        recurrence.DayOfMonth = 10;
        recurrence.DayOfMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse1 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse1.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the DayOfMonth element is included in a request and the Type element value is set to 0.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including DayOfMonth sub-element when Type is '1' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("1");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set DayOfMonth
        recurrence.DayOfMonth = 10;
        recurrence.DayOfMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse2 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse2.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the DayOfMonth element is included in a request and the Type element value is set to 1.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including DayOfMonth sub-element when Type is '3' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("3");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set DayOfMonth
        recurrence.DayOfMonth = 10;
        recurrence.DayOfMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse3 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse3.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the DayOfMonth element is included in a request and the Type element value is set to 3.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including DayOfMonth sub-element when Type is '6' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("6");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set DayOfMonth
        recurrence.DayOfMonth = 10;
        recurrence.DayOfMonthSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse4 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse4.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the DayOfMonth element is included in a request and the Type element value is set to 6.");

        #endregion

        // According to above steps, requirement MS-ASCAL_R53916 can be covered directly.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53916");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53916
        Site.CaptureRequirement(
            53916,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The DayOfMonth element (section 2.2.2.12) is included in a request when the value of the Type element is not 2 or 5.");
    }

    #endregion

    #region MSASCAL_S01_TC25_Status6WithDayOfWeek

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including a DayOfWeek element.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC25_Status6WithDayOfWeek()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar with the element Recurrence including DayOfWeek sub-element when Type is '2' to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        // Set Calendar Recurrence element
        var recurrenceType = byte.Parse("2");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set DayOfWeek
        recurrence.DayOfWeek = 1;
        recurrence.DayOfWeekSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse1 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse1.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the DayOfWeek element is included in a request and the Type element value is set to 2.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including DayOfWeek sub-element when Type is '5' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("5");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set DayOfWeek
        recurrence.DayOfWeek = 1;
        recurrence.DayOfWeekSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse2 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse2.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the DayOfWeek element is included in a request and the Type element value is set to 5.");

        #endregion

        // According to above steps, requirement MS-ASCAL_R53917 can be covered directly.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53917");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53917
        Site.CaptureRequirement(
            53917,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The DayOfWeek element (section 2.2.2.13) is included in a request when the value of the Type element is not 0 (zero), 1, 3, or 6.");
    }

    #endregion

    #region MSASCAL_S01_TC26_Status6WithMonthOfYear

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including a MonthOfYear element.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC26_Status6WithMonthOfYear()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar with the element Recurrence including MonthOfYear sub-element when Type is '0' to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        // Set Calendar Recurrence element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set MonthOfYear
        recurrence.MonthOfYear = 10;
        recurrence.MonthOfYearSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse1 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse1.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the MonthOfYear element is included in a request and the Type element value is set to 0.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including MonthOfYear sub-element when Type is '1' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("1");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set MonthOfYear
        recurrence.MonthOfYear = 10;
        recurrence.MonthOfYearSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse2 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse2.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the MonthOfYear element is included in a request and the Type element value is set to 1.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including MonthOfYear sub-element when Type is '2' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("2");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set MonthOfYear
        recurrence.MonthOfYear = 10;
        recurrence.MonthOfYearSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse3 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse3.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the MonthOfYear element is included in a request and the Type element value is set to 2.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including MonthOfYear sub-element when Type is '3' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("3");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set MonthOfYear
        recurrence.MonthOfYear = 10;
        recurrence.MonthOfYearSpecified = true;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse4 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse4.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the MonthOfYear element is included in a request and the Type element value is set to 3.");

        #endregion

        // According to above steps, requirement MS-ASCAL_R53918 can be covered directly.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53918");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53918
        Site.CaptureRequirement(
            53918,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The MonthOfYear element (section 2.2.2.27) is included in a request when the value of the Type element is not 5 or 6.");
    }

    #endregion

    #region MSASCAL_S01_TC27_Status6WithWeekOfMonth

    /// <summary>
    /// This test case is designed to verify the server will respond with status code 6 when a calendar class with a Recurrence element including a WeekOfMonth element.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC27_Status6WithWeekOfMonth()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar with the element Recurrence including WeekOfMonth sub-element when Type is '0' to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        // Set Calendar Recurrence element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set WeekOfMonth
        recurrence.WeekOfMonthSpecified = true;
        recurrence.WeekOfMonth = 3;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse1 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse1.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the WeekOfMonth element is included in a request and the Type element value is set to 0.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including WeekOfMonth sub-element when Type is '1' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("1");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set WeekOfMonth
        recurrence.WeekOfMonthSpecified = true;
        recurrence.WeekOfMonth = 3;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse2 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse2.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the WeekOfMonth element is included in a request and the Type element value is set to 1.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including WeekOfMonth sub-element when Type is '2' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("2");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set WeekOfMonth
        recurrence.WeekOfMonthSpecified = true;
        recurrence.WeekOfMonth = 3;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse3 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse3.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the WeekOfMonth element is included in a request and the Type element value is set to 2.");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including WeekOfMonth sub-element when Type is '5' to the server, and sync calendars from the server.

        calendarItem.Clear();

        // Set Calendar Recurrence element
        recurrenceType = byte.Parse("5");
        recurrence = CreateCalendarRecurrence(recurrenceType, 3, 3);

        // Set WeekOfMonth
        recurrence.WeekOfMonthSpecified = true;
        recurrence.WeekOfMonth = 3;

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        var addCalendarResponse4 = AddSyncCalendar(calendarItem);

        Site.Assert.AreEqual<string>(
            "6",
            addCalendarResponse4.AddResponses[0].Status,
            "The Sync command response should contain an airsync:Status element with a value of 6 when the WeekOfMonth element is included in a request and the Type element value is set to 5.");

        #endregion

        // According to above steps, requirement MS-ASCAL_R53919 can be covered directly.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R53919");

        // Verify MS-ASCAL requirement: MS-ASCAL_R53919
        Site.CaptureRequirement(
            53919,
            @"[In Sync Command Response][The Sync command response contains an airsync:Status element ([MS-ASCMD] section 2.2.3.162.16) with a value of 6 in the following cases:] The WeekOfMonth element (section 2.2.2.46) is included in a request when the value of the Type element is not 3 or 6.");
    }

    #endregion

    #region MSASCAL_S01_TC28_GhostedElements

    /// <summary>
    /// This test case is designed to verify ghosted elements.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC28_GhostedElements()
    {
        #region Call Sync command to add a calendar to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var exceptionStartTime = StartTime.AddDays(3);
        var startTimeInException = exceptionStartTime.AddMinutes(15);
        var endTimeInException = startTimeInException.AddHours(2);

        // Set Calendar StartTime, EndTime elements
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, EndTime.ToString("yyyyMMddTHHmmssZ"));

        // Set Calendar Recurrence element including Occurrence sub-element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 6, 1);

        // Set Calendar Exceptions element
        var exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
        var exceptionList = new List<Request.ExceptionsException>();

        // Set ExceptionStartTime element in exception
        var exception = TestSuiteHelper.CreateExceptionRequired(exceptionStartTime.ToString("yyyyMMddTHHmmssZ"));

        exception.StartTime = startTimeInException.ToString("yyyyMMddTHHmmssZ");
        exception.EndTime = endTimeInException.ToString("yyyyMMddTHHmmssZ");

        exception.Subject = "Calendar Exception";
        exception.Location = "Room 666";
        exceptionList.Add(exception);
        exceptions.Exception = exceptionList.ToArray();

        if (IsActiveSyncProtocolVersion121
            || IsActiveSyncProtocolVersion140
            || IsActiveSyncProtocolVersion141)
        {
            calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
            calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);
            calendarItem.Add(Request.ItemsChoiceType8.Location1, Location);
        }

        // Set elements which can be ghosted
        var emailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        calendarItem.Add(Request.ItemsChoiceType8.Attendees, TestSuiteHelper.CreateAttendeesRequired(new string[] { emailAddress }, new string[] { User2Information.UserName }));
        calendarItem.Add(Request.ItemsChoiceType8.MeetingStatus, (byte)1);
        if (!IsActiveSyncProtocolVersion121)
        {
            calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            calendarItem.Add(Request.ItemsChoiceType8.DisallowNewTimeProposal, true);
        }

        var subject = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

        AddSyncCalendar(calendarItem);

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

        SyncChanges(User1Information.CalendarCollectionId);

        #endregion

        #region Call Sync command to change a calendar element, and sync calendars from the server.

        // To support ghosted elements of Calendar, following elements must be included in Supported element.
        Request.Supported supportedElement = null;

        // All Calendar class properties are ghosted by default when protocol version 16.0 is used.
        if (IsActiveSyncProtocolVersion121
            || IsActiveSyncProtocolVersion140
            || IsActiveSyncProtocolVersion141)
        {
            supportedElement = new Request.Supported();
            var supportedItem = new Dictionary<Request.ItemsChoiceType, object>
            {
                {
                    Request.ItemsChoiceType.Exceptions, exceptions
                },
                {
                    Request.ItemsChoiceType.DtStamp, string.Empty
                },
                {
                    Request.ItemsChoiceType.Categories, TestSuiteHelper.CreateCalendarCategories(new string[] { "Categories" })
                },
                {
                    Request.ItemsChoiceType.Sensitivity, (byte)1
                },
                {
                    Request.ItemsChoiceType.BusyStatus, (byte)1
                },
                {
                    Request.ItemsChoiceType.UID, string.Empty
                },
                {
                    Request.ItemsChoiceType.Timezone, string.Empty
                },
                {
                    Request.ItemsChoiceType.StartTime, string.Empty
                },
                {
                    Request.ItemsChoiceType.EndTime, string.Empty
                },
                {
                    Request.ItemsChoiceType.Subject, string.Empty
                },
                {
                    Request.ItemsChoiceType.Location, string.Empty
                },
                {
                    Request.ItemsChoiceType.Recurrence, recurrence
                },
                {
                    Request.ItemsChoiceType.AllDayEvent, (byte)1
                },
                {
                    Request.ItemsChoiceType.Reminder, string.Empty
                }
            };

            supportedElement.Items = supportedItem.Values.ToArray<object>();
            supportedElement.ItemsElementName = supportedItem.Keys.ToArray<Request.ItemsChoiceType>();
        }

        // Sync calendars with supported element
        var syncResponse1 = InitializeSync(User1Information.CalendarCollectionId, supportedElement);
        var syncRequest = TestSuiteHelper.CreateSyncRequest(User1Information.CalendarCollectionId, syncResponse1.SyncKey, true);
        var syncResponse2 = CALAdapter.Sync(syncRequest);

        // Update Subject value
        var changeItem = new Dictionary<Request.ItemsChoiceType7, object>();

        var newSubject = Common.GenerateResourceName(Site, "newSubject");
        changeItem.Add(Request.ItemsChoiceType7.Subject, newSubject);

        UpdateCalendarProperty(calendar.ServerId, User1Information.CalendarCollectionId, syncResponse2.SyncKey, changeItem);

        var newCalendar = GetChangeItem(User1Information.CalendarCollectionId, newSubject);

        if (newCalendar.Calendar != null)
        {
            RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, newSubject);
        }
        else
        {
            RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);
            Site.Assert.IsNotNull(newCalendar.Calendar, "The calendar with subject {0} should exist in server.", newSubject);
        }

        #endregion

        #region Verify Requirements.

        Site.Assert.IsNotNull(newCalendar.Calendar.Body, "The Body element should not be null.");
        Site.Assert.IsNotNull(calendar.Calendar.Body, "The Body element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R128");

        // Verify MS-ASCAL requirement: MS-ASCAL_R128
        Site.CaptureRequirementIfIsTrue(
            newCalendar.Calendar.Body.Type == calendar.Calendar.Body.Type && newCalendar.Calendar.Body.Data == calendar.Calendar.Body.Data,
            128,
            @"[In Body (AirSyncBase Namespace)] The top-level airsyncbase:Body element can be ghosted.");

        if (!IsActiveSyncProtocolVersion121)
        {
            Site.Assert.IsNotNull(calendar.Calendar.ResponseRequested, "The ResponseRequested element should not be null.");
            Site.Assert.IsNotNull(newCalendar.Calendar.ResponseRequested, "The ResponseRequested element should not be null.");
            Site.Assert.IsNotNull(calendar.Calendar.ResponseType, "The ResponseType element should not be null.");
            Site.Assert.IsNotNull(newCalendar.Calendar.ResponseType, "The ResponseType element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R399");

            // Verify MS-ASCAL requirement: MS-ASCAL_R399
            Site.CaptureRequirementIfAreEqual<bool>(
                calendar.Calendar.ResponseRequested.Value,
                newCalendar.Calendar.ResponseRequested.Value,
                399,
                @"[In ResponseRequested] The ResponseRequested element can be ghosted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R419");

            // Verify MS-ASCAL requirement: MS-ASCAL_R419
            Site.CaptureRequirementIfAreEqual<uint>(
                calendar.Calendar.ResponseType.Value,
                newCalendar.Calendar.ResponseType.Value,
                419,
                @"[In ResponseType] The top-level ResponseType element can be ghosted.");
        }

        Site.Assert.IsNotNull(calendar.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");
        Site.Assert.IsNotNull(newCalendar.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R319");

        // Verify MS-ASCAL requirement: MS-ASCAL_R319
        Site.CaptureRequirementIfAreEqual<byte>(
            calendar.Calendar.MeetingStatus.Value,
            newCalendar.Calendar.MeetingStatus.Value,
            319,
            @"[In MeetingStatus] The top-level MeetingStatus element can be ghosted.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R365");

        // Verify MS-ASCAL requirement: MS-ASCAL_R365
        Site.CaptureRequirementIfAreEqual<string>(
            calendar.Calendar.OrganizerEmail.ToLower(CultureInfo.CurrentCulture),
            newCalendar.Calendar.OrganizerEmail.ToLower(CultureInfo.CurrentCulture),
            365,
            @"[In OrganizerEmail] The OrganizerEmail element can be ghosted.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R370");

        // Verify MS-ASCAL requirement: MS-ASCAL_R370
        Site.CaptureRequirementIfAreEqual<string>(
            calendar.Calendar.OrganizerName.ToLower(CultureInfo.CurrentCulture),
            newCalendar.Calendar.OrganizerName.ToLower(CultureInfo.CurrentCulture),
            370,
            @"[In OrganizerName] The OrganizerName element can be ghosted.");

        // There is only one attendee.
        Site.Assert.AreEqual<int>(
            1,
            newCalendar.Calendar.Attendees.Attendee.Length,
            "The Attendees element should be ghosted");

        var isR111Verified = newCalendar.Calendar.Attendees.Attendee[0].AttendeeStatusSpecified == calendar.Calendar.Attendees.Attendee[0].AttendeeStatusSpecified &&
                             newCalendar.Calendar.Attendees.Attendee[0].AttendeeStatus == calendar.Calendar.Attendees.Attendee[0].AttendeeStatus &&
                             newCalendar.Calendar.Attendees.Attendee[0].AttendeeTypeSpecified == calendar.Calendar.Attendees.Attendee[0].AttendeeTypeSpecified &&
                             newCalendar.Calendar.Attendees.Attendee[0].AttendeeType == calendar.Calendar.Attendees.Attendee[0].AttendeeType &&
                             newCalendar.Calendar.Attendees.Attendee[0].Email == calendar.Calendar.Attendees.Attendee[0].Email &&
                             newCalendar.Calendar.Attendees.Attendee[0].Name == calendar.Calendar.Attendees.Attendee[0].Name;

        // Add the debug information
        Site.Log.Add(
            LogEntryKind.Debug,
            "Verify MS-ASCAL_R111.\n" + "The AttendeeStatus is {0};\n" + "The AttendeeType is {1};\n" + "The Email is {2};\n" + "The Name is {3}.",
            newCalendar.Calendar.Attendees.Attendee[0].AttendeeStatus,
            newCalendar.Calendar.Attendees.Attendee[0].AttendeeType,
            newCalendar.Calendar.Attendees.Attendee[0].Email,
            newCalendar.Calendar.Attendees.Attendee[0].Name);

        // Verify MS-ASCAL requirement: MS-ASCAL_R111
        Site.CaptureRequirementIfIsTrue(
            isR111Verified,
            111,
            @"[In Attendees] The top-level Attendees element can be ghosted.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R217");

        // Verify MS-ASCAL requirement: MS-ASCAL_R217
        Site.CaptureRequirementIfAreEqual<bool?>(
            calendar.Calendar.DisallowNewTimeProposal,
            newCalendar.Calendar.DisallowNewTimeProposal,
            217,
            @"[In DisallowNewTimeProposal] The DisallowNewTimeProposal element can be ghosted.");

        if (IsActiveSyncProtocolVersion121 || IsActiveSyncProtocolVersion140 || IsActiveSyncProtocolVersion141)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R301");

            // Verify MS-ASCAL requirement: MS-ASCAL_R301
            Site.CaptureRequirementIfIsNull(
                newCalendar.Calendar.Location,
                301,
                @"[In Location] The top-level Location element cannot be ghosted.");
        }

        #endregion
    }

    #endregion

    #region MSASCAL_S01_TC29_ItemOperations

    /// <summary>
    /// This test case is designed to when the client uses ItemOperations command in the default inline way to Fetch the calendar, the server responds with part element instead of data element in the Calendar's body.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC29_ItemOperations()
    {
        #region Call Sync command to add a calendar to the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subject = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

        // Add a default calendar
        AddSyncCalendar(calendarItem);

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);

        #endregion

        #region Call method ItemOperations to fetch all the information about calendars using ServerIds and get the expected success response.

        // To verify MS-ASCAL_R530, just include subject element in schema.
        var schema = new Request.Schema();
        var elements = new List<object> { string.Empty };

        var names = new List<Request.ItemsChoiceType4>
        {
            Request.ItemsChoiceType4.Subject
        };

        schema.Items = elements.ToArray();
        schema.ItemsElementName = names.ToArray();

        // The server id of Calendar
        var serverIds = new List<string> { calendar.ServerId };
        var itemOperationsRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(User1Information.CalendarCollectionId, serverIds, schema);
        var fetchResponse = CALAdapter.ItemOperations(itemOperationsRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R530");

        // Verify MS-ASEMAIL requirement: MS-ASCAL_R530
        Site.CaptureRequirementIfIsTrue(
            TestSuiteHelper.IsOnlySpecifiedElement((XmlElement)CALAdapter.LastRawResponseXml, "Properties", "Subject"),
            530,
            @"[In ItemOperations Command Response] If an airsync:Schema element ([MS-ASCMD] section 2.2.3.145) is included in the ItemOperations command request, the elements returned in the ItemOperations command response MUST be restricted to the elements that were included as child elements of the airsync:Schema element in the command request.");

        // Verify ItemOperations response
        Site.Assert.AreEqual<string>(
            "1",
            fetchResponse.Status,
            "If the ItemOperations command executes successfully, the Status in response should be 1.");

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R525");

        // Verify MS-ASCAL requirement: MS-ASCAL_R525
        // If ItemOperations response can get a non-null ItemOperationsStore.CalendarItems, it means the client had fetched the calendar,
        // then we can capture this requirement
        Site.CaptureRequirementIfIsNotNull(
            fetchResponse.Items,
            525,
            @"[In Retrieving Details for One or More Calendar Items] The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R528");

        // Verify MS-ASCAL requirement: MS-ASCAL_R528
        // If ItemOperations response can get a non-null ItemOperationsStore.CalendarItems, it means the client had fetched the calendar,
        // then we can capture this requirement
        Site.CaptureRequirementIfIsNotNull(
            fetchResponse.Items,
            528,
            @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.2.8), as specified in section 3.1.5.1, to retrieve data from the server for one or more specific calendar items, the server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");
    }

    #endregion

    #region MSASCAL_S01_TC30_Search

    /// <summary>
    /// This test case is designed to verify the client calls Search command request to search calendars using the given keyword text, the calendar which satisfies the condition returned.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC30_Search()
    {
        #region Call Sync command to add a calendar to the server.

        // Add a default calendar
        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var subject = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

        // Add a default calendar
        AddSyncCalendar(calendarItem);

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);

        #endregion

        #region Call method Search to Search calendars using the given keyword text.

        // Wait for a period after calendars are created for the search command to get results.
        var waitTime = Convert.ToInt32(Common.GetConfigurationPropertyValue("WaitTime", Site));
        var retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var counter = 0;

        var searchRequest = TestSuiteHelper.CreateSearchRequest(SearchName.Mailbox.ToString(), subject, User1Information.CalendarCollectionId);
        SearchStore searchResponse;
        do
        {
            Thread.Sleep(waitTime);

            // Search the Calendar
            searchResponse = CALAdapter.Search(searchRequest);
            counter++;
        }
        while (searchResponse.Total == 0 && counter < retryCount);

        // Verify search response
        Site.Assert.AreEqual<string>(
            "1",
            searchResponse.Status,
            "If the Search command executes successfully, the Status in response should be 1.");

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R533");

        // Verify MS-ASCAL requirement: MS-ASCAL_R533
        // If Search response can get a non-null SearchStore.Results, it means the client had searched the calendar which satisfies the condition,
        // then we can capture this requirement
        Site.CaptureRequirementIfIsNotNull(
            searchResponse.Results,
            533,
            @"[In Search Command Response] When a client uses the Search command request ([MS-ASCMD] section 2.2.2.14), as specified in section 3.1.5.2, to retrieve Calendar class items from the server that match the criteria specified by the client, the server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R523");

        // Verify MS-ASCAL requirement: MS-ASCAL_R523
        // If Search response can get a non-null SearchStore.Results, it means the client had searched the calendar which satisfies the condition,
        // then we can capture this requirement
        Site.CaptureRequirementIfIsNotNull(
            searchResponse.Results,
            523,
            @"[In Searching for Calendar Data] The server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");
    }

    #endregion

    #region MSASCAL_S01_TC31_UnchangedExceptions

    /// <summary>
    /// This test case is designed to verify exceptions via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC31_UnchangedExceptions()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Calls Sync command to add a calendar to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var exceptionStartTime = StartTime.AddDays(3);
        var startTimeInException = exceptionStartTime.AddMinutes(15);
        var endTimeInException = startTimeInException.AddHours(2);

        // Set Calendar StartTime, EndTime elements
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, EndTime.ToString("yyyyMMddTHHmmssZ"));

        // Set Calendar Recurrence element including Occurrence sub-element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 6, 1);
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

        // Set Calendar Exceptions element
        var exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
        var exceptionList = new List<Request.ExceptionsException>();

        // Set ExceptionStartTime element in exception
        var exception = TestSuiteHelper.CreateExceptionRequired(exceptionStartTime.ToString("yyyyMMddTHHmmssZ"));

        exception.StartTime = startTimeInException.ToString("yyyyMMddTHHmmssZ");
        exception.EndTime = endTimeInException.ToString("yyyyMMddTHHmmssZ");

        exception.Subject = "Calendar Exception";
        exception.Location = "Room 666";
        exceptionList.Add(exception);
        exceptions.Exception = exceptionList.ToArray();
        calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);

        var subject = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

        AddSyncCalendar(calendarItem);

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);

        #endregion

        #region Calls Sync command to Sync calendars from the server.

        // Set Supported Element
        var supportedElement = new Request.Supported();

        var supportedItem = new Dictionary<Request.ItemsChoiceType, object>
        {
            {
                Request.ItemsChoiceType.Exceptions, exceptions
            },
            {
                Request.ItemsChoiceType.DtStamp, string.Empty
            },
            {
                Request.ItemsChoiceType.Categories, TestSuiteHelper.CreateCalendarCategories(new string[] { "Categories" })
            },
            {
                Request.ItemsChoiceType.Sensitivity, (byte)1
            },
            {
                Request.ItemsChoiceType.BusyStatus, (byte)1
            },
            {
                Request.ItemsChoiceType.UID, string.Empty
            },
            {
                Request.ItemsChoiceType.Timezone, string.Empty
            },
            {
                Request.ItemsChoiceType.StartTime, string.Empty
            },
            {
                Request.ItemsChoiceType.EndTime, string.Empty
            },
            {
                Request.ItemsChoiceType.Subject, string.Empty
            },
            {
                Request.ItemsChoiceType.Location, string.Empty
            },
            {
                Request.ItemsChoiceType.Recurrence, recurrence
            },
            {
                Request.ItemsChoiceType.AllDayEvent, (byte)1
            },
            {
                Request.ItemsChoiceType.Reminder, string.Empty
            }
        };

        supportedElement.Items = supportedItem.Values.ToArray<object>();
        supportedElement.ItemsElementName = supportedItem.Keys.ToArray<Request.ItemsChoiceType>();

        // Sync calendars with supported element
        var syncResponse1 = InitializeSync(User1Information.CalendarCollectionId, supportedElement);

        var syncRequest = TestSuiteHelper.CreateSyncRequest(User1Information.CalendarCollectionId, syncResponse1.SyncKey, true);
        var syncResponse2 = CALAdapter.Sync(syncRequest);
        var createdCalendar = new SyncItem();

        foreach (var item in syncResponse2.AddElements)
        {
            if (item.Calendar.Subject == subject)
            {
                createdCalendar = item;
                break;
            }
        }

        Site.Assert.IsNotNull(createdCalendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

        syncRequest = TestSuiteHelper.CreateSyncRequest(User1Information.CalendarCollectionId, syncResponse2.SyncKey, false);
        var syncResponse3 = CALAdapter.Sync(syncRequest);

        Site.Assert.AreEqual<int>(
            0,
            syncResponse3.AddResponses.Count,
            "This Sync command response should be null.");

        var updatedCalendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(updatedCalendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

        #endregion

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R54111");

        // Verify MS-ASCAL requirement: MS-ASCAL_R54111
        Site.CaptureRequirementIfIsTrue(
            createdCalendar.Calendar.Exceptions.Exception[0].ExceptionStartTime == calendar.Calendar.Exceptions.Exception[0].ExceptionStartTime
            && createdCalendar.Calendar.Exceptions.Exception[0].StartTime == calendar.Calendar.Exceptions.Exception[0].StartTime
            && createdCalendar.Calendar.Exceptions.Exception[0].EndTime == calendar.Calendar.Exceptions.Exception[0].EndTime
            && createdCalendar.Calendar.Exceptions.Exception[0].Subject == calendar.Calendar.Exceptions.Exception[0].Subject
            && createdCalendar.Calendar.Exceptions.Exception[0].Location == calendar.Calendar.Exceptions.Exception[0].Location,
            54111,
            @"[In Removing Exceptions] [If an Exceptions element (section 2.2.2.20) is not specified in a Sync command request ([MS-ASCMD] section 2.2.2.19.2), then] any exceptions previously defined are unchanged, even if the client included the Exceptions element as a child of the Supported element, as specified in [MS-ASCMD] section 2.2.3.164.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R541");

        // Verify MS-ASCAL requirement: MS-ASCAL_R541
        Site.CaptureRequirementIfIsTrue(
            updatedCalendar.Calendar.Exceptions.Exception[0].ExceptionStartTime == calendar.Calendar.Exceptions.Exception[0].ExceptionStartTime
            && updatedCalendar.Calendar.Exceptions.Exception[0].StartTime == calendar.Calendar.Exceptions.Exception[0].StartTime
            && updatedCalendar.Calendar.Exceptions.Exception[0].EndTime == calendar.Calendar.Exceptions.Exception[0].EndTime
            && updatedCalendar.Calendar.Exceptions.Exception[0].Subject == calendar.Calendar.Exceptions.Exception[0].Subject
            && updatedCalendar.Calendar.Exceptions.Exception[0].Location == calendar.Calendar.Exceptions.Exception[0].Location,
            541,
            @"[In Removing Exceptions] If an Exceptions element (section 2.2.2.20) is not specified in a Sync command request ([MS-ASCMD] section 2.2.2.19), then any exceptions previously defined are unchanged. ");
    }

    #endregion

    #region MSASCAL_S01_TC32_WithoutUID

    /// <summary>
    /// This test case is designed to verify server behavior when the UID element is not included in the command request.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC32_WithoutUID()
    {
        #region Call Sync command to add a calendar without element UID to the server, and sync calendars from the server.
        var calendarItem = CreateDefaultCalendar();
        calendarItem.Remove(Request.ItemsChoiceType8.UID);
        var addCalendar = new Request.SyncCollectionAddApplicationData
        {
            Items = calendarItem.Values.ToArray<object>(),
            ItemsElementName = calendarItem.Keys.ToArray<Request.ItemsChoiceType8>()
        };

        // Sync to get the SyncKey
        var initializeSyncResponse = InitializeSync(CurrentUserInformation.CalendarCollectionId, null);

        // Add the calendar item
        var syncRequest = TestSuiteHelper.CreateSyncAddRequest(CurrentUserInformation.CalendarCollectionId, initializeSyncResponse.SyncKey, addCalendar);
        var syncCalendarResponse = CALAdapter.Sync(syncRequest);

        // Verify sync response, if the Sync command executes successfully, the Status in response should be 1.
        Site.Assert.AreEqual<byte>(
            1,
            syncCalendarResponse.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, SubjectName);
        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", SubjectName);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, SubjectName);
        #endregion

        if (!IsActiveSyncProtocolVersion121
            && !IsActiveSyncProtocolVersion140
            && !IsActiveSyncProtocolVersion141)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2223");

            // Verify MS-ASCAL requirement: MS-ASCAL_R2223
            Site.CaptureRequirementIfIsNotNull(
                calendar.Calendar.UID,
                2223,
                @"[In UID] When a calendar item is created, the server will generate a unique identifier for the calendar item and return the identifier in the UID element of the Sync command response ([MS-ASCMD] section 2.2.2.20) for an add operation.");
        }
    }

    #endregion

    #region MSASCAL_S01_TC33_DeletePropertyOfException

    /// <summary>
    /// This test case is designed to verify server transmits empty element in response if property of an exception
    /// for recurring calendar item has been deleted.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC33_DeletePropertyOfException()
    {
        Site.Assume.IsTrue(Common.IsRequirementEnabled(2242, Site), "Exchange 2007 does not support deleting elements of a recurring calendar item in an Exception element.");
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var exceptionStartTime = StartTime.AddDays(3);

        // Set Calendar StartTime, EndTime elements
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, EndTime.ToString("yyyyMMddTHHmmssZ"));

        // Set Calendar Recurrence element including Occurrence sub-element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 6, 1);

        // Set Calendar Exceptions element
        var exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
        var exceptionList = new List<Request.ExceptionsException>();

        // Set ExceptionStartTime element in exception
        var exception = TestSuiteHelper.CreateExceptionRequired(exceptionStartTime.ToString("yyyyMMddTHHmmssZ"));

        exception.Subject = "Calendar Exception";
        exception.Location = "Room 666";
        exceptionList.Add(exception);
        exceptions.Exception = exceptionList.ToArray();

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);
        calendarItem.Add(Request.ItemsChoiceType8.Location1, Location);

        var emailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        calendarItem.Add(Request.ItemsChoiceType8.Attendees, TestSuiteHelper.CreateAttendeesRequired(new string[] { emailAddress }, new string[] { User2Information.UserName }));
        calendarItem.Add(Request.ItemsChoiceType8.MeetingStatus, (byte)1);
        if (!IsActiveSyncProtocolVersion121)
        {
            calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            calendarItem.Add(Request.ItemsChoiceType8.DisallowNewTimeProposal, true);
        }

        var subject = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

        AddSyncCalendar(calendarItem);

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);
        #endregion

        #region Call Sync command to delete the Location property of the exception to change the calendar, and sync calendars from the server.

        var syncResponse1 = InitializeSync(User1Information.CalendarCollectionId, null);
        var syncRequest = TestSuiteHelper.CreateSyncRequest(User1Information.CalendarCollectionId, syncResponse1.SyncKey, true);
        var syncResponse2 = CALAdapter.Sync(syncRequest);

        // Delete Location property of the Exception
        var changeItem = new Dictionary<Request.ItemsChoiceType7, object>();
        exception.Location = null;
        changeItem.Add(Request.ItemsChoiceType7.Exceptions, exceptions);
        changeItem.Add(Request.ItemsChoiceType7.Recurrence, recurrence);
        changeItem.Add(Request.ItemsChoiceType7.Subject, subject);
        var syncChangeData = new Request.SyncCollectionChangeApplicationData
        {
            ItemsElementName = changeItem.Keys.ToArray<Request.ItemsChoiceType7>(),
            Items = changeItem.Values.ToArray<object>()
        };

        var syncChange = new Request.SyncCollectionChange
        {
            ApplicationData = syncChangeData,
            ServerId = calendar.ServerId
        };

        var syncChangeRequest = new SyncRequest
        {
            RequestData = new Request.Sync { Collections = new Request.SyncCollection[1] }
        };

        syncChangeRequest.RequestData.Collections[0] = new Request.SyncCollection
        {
            Commands = new object[] { syncChange },
            SyncKey = syncResponse2.SyncKey,
            CollectionId = User1Information.CalendarCollectionId
        };

        // If an element in a recurring calendar item has been deleted in an Exception element, sends an empty tag
        // for this element to remove the inherited value from the server.
        var syncXmlRequest = syncChangeRequest.GetRequestDataSerializedXML();
        var changedSyncXmlRequest = syncXmlRequest.Insert(syncXmlRequest.IndexOf("</Exception>", StringComparison.CurrentCulture), "<Location />");
        var result = CALAdapter.SendStringRequest(changedSyncXmlRequest);

        #endregion

        #region Call Sync command to get the changed calendar.

        var initializeSyncResponse = InitializeSync(User1Information.CalendarCollectionId, null);
        syncRequest = TestSuiteHelper.CreateSyncRequest(User1Information.CalendarCollectionId, initializeSyncResponse.SyncKey, true);
        result = CALAdapter.SendStringRequest(syncRequest.GetRequestDataSerializedXML());

        var doc = new XmlDocument();
        doc.LoadXml(result.ResponseDataXML);
        var nameSpaceManager = new XmlNamespaceManager(doc.NameTable);
        nameSpaceManager.AddNamespace("e", "AirSync");
        var nodes = doc.SelectNodes("//e:Collections/e:Collection/e:Commands/e:Add/e:ApplicationData", nameSpaceManager);
        var isEmptyLocationContained = false;
        foreach (XmlNode node in nodes)
        {
            var isFound = false;
            var subNodes = node.ChildNodes;
            foreach (XmlNode subNode in subNodes)
            {
                if (subNode.Name.Equals("Subject") && subNode.InnerText != null && subNode.InnerText.Equals(subject))
                {
                    isFound = true;
                }
                if (isFound && subNode.Name.Equals("Exceptions"))
                {
                    isEmptyLocationContained = subNode.InnerXml.Contains("<Location />");
                    break;
                }
            }
            if (isEmptyLocationContained)
            {
                break;
            }
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R2242");

        // Verify MS-ASCAL requirement: MS-ASCAL_R2242
        Site.CaptureRequirementIfIsTrue(
            isEmptyLocationContained,
            2242,
            @"[In Appendix B: Product Behavior]  If an element in a recurring calendar item has been deleted in an Exception element (section 2.2.2.19), the client sends an empty tag for this element to remove the inherited value from the implementation. (Exchange 2010 and above follow this behavior.)");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R539");

        // Verify MS-ASCAL requirement: MS-ASCAL_R539
        Site.CaptureRequirementIfIsTrue(
            isEmptyLocationContained,
            539,
            @"[In Sync Command Response] If one or more properties of an exception for recurring calendar item (that is, any child elements of the Exception element (section 2.2.2.19)) have been deleted, the server MUST transmit an empty element in the Sync command response to indicate that this property is not inherited from the recurrence.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R543");

        // Verify MS-ASCAL requirement: MS-ASCAL_R543
        Site.CaptureRequirementIfIsTrue(
            isEmptyLocationContained,
            543,
            @"[In Indicating Deleted Elements in Exceptions] If an element of a recurring calendar item has been deleted in an Exception element (section 2.2.2.19), the server MUST send an empty tag for this element in the Sync command response ([MS-ASCMD] section 2.2.2.19).");

        #endregion
    }

    #endregion

    #region MSASCAL_S01_TC34_ExcludePropertyOfException

    /// <summary>
    /// This test case is designed to verify if a particular Exception element is excluded in a Sync command request,
    /// then that particular exception remains unchanged.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC34_ExcludePropertyOfException()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            
        #region Call Sync command to add a calendar to the server, and sync calendars from the server.

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

        var exceptionStartTime = StartTime.AddDays(3);

        // Set Calendar StartTime, EndTime elements
        calendarItem.Add(Request.ItemsChoiceType8.StartTime, StartTime.ToString("yyyyMMddTHHmmssZ"));
        calendarItem.Add(Request.ItemsChoiceType8.EndTime, EndTime.ToString("yyyyMMddTHHmmssZ"));

        // Set Calendar Recurrence element including Occurrence sub-element
        var recurrenceType = byte.Parse("0");
        var recurrence = CreateCalendarRecurrence(recurrenceType, 6, 1);

        // Set Calendar Exceptions element
        var exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
        var exceptionList = new List<Request.ExceptionsException>();

        // Set ExceptionStartTime element in exception
        var exception = TestSuiteHelper.CreateExceptionRequired(exceptionStartTime.ToString("yyyyMMddTHHmmssZ"));

        exception.Subject = "Calendar Exception";
        exception.Location = "Room 666";
        exceptionList.Add(exception);
        exceptions.Exception = exceptionList.ToArray();

        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);
        calendarItem.Add(Request.ItemsChoiceType8.Location1, Location);

        var emailAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        calendarItem.Add(Request.ItemsChoiceType8.Attendees, TestSuiteHelper.CreateAttendeesRequired(new string[] { emailAddress }, new string[] { User2Information.UserName }));
        calendarItem.Add(Request.ItemsChoiceType8.MeetingStatus, (byte)1);
        if (!IsActiveSyncProtocolVersion121)
        {
            calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            calendarItem.Add(Request.ItemsChoiceType8.DisallowNewTimeProposal, true);
        }

        var subject = Common.GenerateResourceName(Site, "subject");
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

        AddSyncCalendar(calendarItem);

        var calendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);
        #endregion

        #region Call Sync command to change the Exception element of calendar by excluding the Location property, and sync calendars from the server.

        var syncResponse1 = InitializeSync(User1Information.CalendarCollectionId, null);
        var syncRequest = TestSuiteHelper.CreateSyncRequest(User1Information.CalendarCollectionId, syncResponse1.SyncKey, true);
        var syncResponse2 = CALAdapter.Sync(syncRequest);

        // Exclude Location property of the Exception
        var changeItem = new Dictionary<Request.ItemsChoiceType7, object>();
        exception.Location = null;
        changeItem.Add(Request.ItemsChoiceType7.Exceptions, exceptions);
        changeItem.Add(Request.ItemsChoiceType7.Recurrence, recurrence);
        changeItem.Add(Request.ItemsChoiceType7.Subject, subject);
        UpdateCalendarProperty(calendar.ServerId, User1Information.CalendarCollectionId, syncResponse2.SyncKey, changeItem);

        var newCalendar = GetChangeItem(User1Information.CalendarCollectionId, subject);

        var isUnChanged = newCalendar.Calendar.Exceptions.Exception[0].Subject == exception.Subject
                          && newCalendar.Calendar.Exceptions.Exception[0].Location == "Room 666";

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R542");

        // Verify MS-ASCAL requirement: MS-ASCAL_R542
        Site.CaptureRequirementIfIsTrue(
            isUnChanged,
            542,
            @"[In Removing Exceptions] If a particular Exception element (section 2.2.2.19) is excluded in a Sync command request, then that particular exception remains unchanged.");

        #endregion
    }

    #endregion

    #region MSASCAL_S01_TC35_RecurrenceWithInterval0

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when Interval set as 0 via invoking Sync command.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC35_RecurrenceWithInterval0()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command to add a calendar with the element Recurrence including Type '0' and Occurrences sub-element to the server, and sync calendars from the server.

        var calendarItem = CreateDefaultCalendar();
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, CreateCalendarRecurrence(byte.Parse("0"), 3, 0));
        var addCalendar = new Request.SyncCollectionAddApplicationData
        {
            Items = calendarItem.Values.ToArray<object>(),
            ItemsElementName = calendarItem.Keys.ToArray<Request.ItemsChoiceType8>()
        };

        // Sync to get the SyncKey
        var initializeSyncResponse = InitializeSync(CurrentUserInformation.CalendarCollectionId, null);

        // Add the calendar item
        var syncRequest = TestSuiteHelper.CreateSyncAddRequest(CurrentUserInformation.CalendarCollectionId, initializeSyncResponse.SyncKey, addCalendar);
        var syncCalendarResponse = CALAdapter.Sync(syncRequest);

        if (Common.IsRequirementEnabled(4, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R4");

            // Verify MS-ASCAL requirement: MS-ASCAL_R4
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                syncCalendarResponse.AddResponses[0].Status,
                4,
                @"[In Appendix B: Product Behavior] <2> Section 2.2.2.25:  If Interval is set to 0 in command request, Microsoft Exchange Server 2007 SP1 and Exchange 2010 returns Status value 6;");
        }

        if (Common.IsRequirementEnabled(5, Site))
        {
            var calendar = GetChangeItem(User1Information.CalendarCollectionId, SubjectName);

            Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", SubjectName);
            RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, SubjectName);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R5");

            // Verify MS-ASCAL requirement: MS-ASCAL_R5
            Site.CaptureRequirementIfAreEqual<ushort>(
                1,
                calendar.Calendar.Recurrence.Interval,
                5,
                @"[In Appendix B: Product Behavior] [<2> Section 2.2.2.25:  If Interval is set to 0 in command request,] Microsoft Exchange Server 2010 Service Pack 1 (SP1), Exchange 2013, Exchange 2016 and Exchange 2019 return Interval value 1");
        }

        #endregion
    }

    #endregion

    #region MSASCAL_S01_TC36_RecurrenceWithCalendarType1

    /// <summary>
    /// This test case is designed to verify a calendar class with a Recurrence element when CalendarType set 1.
    /// </summary>
    [TestCategory("MSASCAL"), TestMethod]
    public void MSASCAL_S01_TC36_RecurrenceWithCalendarType1()
    {
        Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The element CalendarType is not supported when protocol version is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Generate calendar subject and record them.

        var recurrenceType = byte.Parse("2");

        var calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();
            
        var subjectWithType2AndCalendarType1 = Common.GenerateResourceName(Site, "subjectWithType2AndCalendarType1");

        #endregion

        #region Call Sync command to add a calendar with the element Recurrence including Type '2' and CalendarType sub-element setting as "1" to the server, and sync calendars from the server.
        var occurrences = 5;
        var interval = 2;

        // Set Calendar Recurrence element, CalendarType is set to "1".
        var calendarType = byte.Parse("1");
        var recurrence = CreateRecurrenceIncludingCalendarType(CreateCalendarRecurrence(recurrenceType, occurrences, interval), calendarType);
        calendarItem.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
        calendarItem.Add(Request.ItemsChoiceType8.Subject, subjectWithType2AndCalendarType1);

        AddSyncCalendar(calendarItem);

        var calendarWithType2AndCalendarType1 = GetChangeItem(User1Information.CalendarCollectionId, subjectWithType2AndCalendarType1);

        Site.Assert.IsNotNull(calendarWithType2AndCalendarType1.Calendar, "The calendar with subject {0} should exist in server.", subjectWithType2AndCalendarType1);
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subjectWithType2AndCalendarType1);

        if (Common.IsRequirementEnabled(3, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                calendarWithType2AndCalendarType1.Calendar.Recurrence.CalendarType,
                3,
                @"[In Appendix B: Product Behavior] The implementation return a value of 0 (Default) when a client specifies a value of 1 (Gregorian). (<1> Section 2.2.2.10:  Microsoft Exchange Server 2013 Service Pack 1 (SP1) returns a value of 0 when a client specifies a value of 1 (Gregorian).)");
        }

        if (Common.IsRequirementEnabled(2239, Site))
        {
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                calendarWithType2AndCalendarType1.Calendar.Recurrence.CalendarType,
                2239,
                @"[In Appendix B: Product Behavior] The implementation does not return a value of 0 (Default) when a client specifies a value of 1 (Gregorian). (Microsoft Exchange Server 2010 follows this behavior.)");
        }
        #endregion
    }

    #endregion
    #endregion
}