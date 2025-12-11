namespace Microsoft.Protocols.TestSuites.MS_ASWBXML;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using TestTools;

/// <summary>
/// The implementation of MS-ASWBXML
/// </summary>
public class MS_ASWBXML
{
    /// <summary>
    /// WBXML version
    /// </summary>
    private const byte VersionByte = 0x03;

    /// <summary>
    /// Public Identifier
    /// </summary>
    private const byte PublicIdentifierByte = 0x01;

    /// <summary>
    /// Encoding. 0x6A == UTF-8
    /// </summary>
    private const byte CharsetByte = 0x6A;

    /// <summary>
    /// String table length. This is not used in MS-ASWBXML, so this value is always 0.
    /// </summary>
    private const byte StringTableLengthByte = 0x00;

    /// <summary>
    /// XmlDocument that contain the xml
    /// </summary>
    private XmlDocument xmlDoc = new XmlDocument();

    /// <summary>
    /// Code pages.
    /// </summary>
    private CodePage[] codePages;

    /// <summary>
    /// Current code page.
    /// </summary>
    private int currentCodePage = 0;

    /// <summary>
    /// Default code page.
    /// </summary>
    private int defaultCodePage = -1;

    /// <summary>
    /// The DataCollection in encoding process
    /// </summary>
    private Dictionary<string, int> encodeDataCollection = new Dictionary<string, int>();

    /// <summary>
    /// The DataCollection in decoding process
    /// </summary>
    private Dictionary<string, int> decodeDataCollection = new Dictionary<string, int>();

    /// <summary>
    /// An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.
    /// </summary>
    private ITestSite site;

    /// <summary>
    /// Initializes a new instance of the MS_ASWBXML class.
    /// </summary>
    /// <param name="site">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
    public MS_ASWBXML(ITestSite site)
    {
        this.site = site;

        // Loads up code pages. There are 26 code pages as per MS-ASWBXML
        codePages = new CodePage[26];

        // Code Page 0: AirSync
        codePages[0] = new CodePage { Namespace = "AirSync", Xmlns = "airsync" };

        codePages[0].AddToken(0x05, "Sync");
        codePages[0].AddToken(0x06, "Responses");
        codePages[0].AddToken(0x07, "Add");
        codePages[0].AddToken(0x08, "Change");
        codePages[0].AddToken(0x09, "Delete");
        codePages[0].AddToken(0x0A, "Fetch");
        codePages[0].AddToken(0x0B, "SyncKey");
        codePages[0].AddToken(0x0C, "ClientId");
        codePages[0].AddToken(0x0D, "ServerId");
        codePages[0].AddToken(0x0E, "Status");
        codePages[0].AddToken(0x0F, "Collection");
        codePages[0].AddToken(0x10, "Class");
        codePages[0].AddToken(0x12, "CollectionId");
        codePages[0].AddToken(0x13, "GetChanges");
        codePages[0].AddToken(0x14, "MoreAvailable");
        codePages[0].AddToken(0x15, "WindowSize");
        codePages[0].AddToken(0x16, "Commands");
        codePages[0].AddToken(0x17, "Options");
        codePages[0].AddToken(0x18, "FilterType");
        codePages[0].AddToken(0x1B, "Conflict");
        codePages[0].AddToken(0x1C, "Collections");
        codePages[0].AddToken(0x1D, "ApplicationData");
        codePages[0].AddToken(0x1E, "DeletesAsMoves");
        codePages[0].AddToken(0x20, "Supported");
        codePages[0].AddToken(0x21, "SoftDelete");
        codePages[0].AddToken(0x22, "MIMESupport");
        codePages[0].AddToken(0x23, "MIMETruncation");
        codePages[0].AddToken(0x24, "Wait");
        codePages[0].AddToken(0x25, "Limit");
        codePages[0].AddToken(0x26, "Partial");
        codePages[0].AddToken(0x27, "ConversationMode");
        codePages[0].AddToken(0x28, "MaxItems");
        codePages[0].AddToken(0x29, "HeartbeatInterval");

        // Code Page 1: Contacts
        codePages[1] = new CodePage { Namespace = "Contacts", Xmlns = "contacts" };

        codePages[1].AddToken(0x05, "Anniversary");
        codePages[1].AddToken(0x06, "AssistantName");
        codePages[1].AddToken(0x07, "AssistantPhoneNumber");
        codePages[1].AddToken(0x08, "Birthday");
        codePages[1].AddToken(0x0C, "Business2PhoneNumber");
        codePages[1].AddToken(0x0D, "BusinessAddressCity");
        codePages[1].AddToken(0x0E, "BusinessAddressCountry");
        codePages[1].AddToken(0x0F, "BusinessAddressPostalCode");
        codePages[1].AddToken(0x10, "BusinessAddressState");
        codePages[1].AddToken(0x11, "BusinessAddressStreet");
        codePages[1].AddToken(0x12, "BusinessFaxNumber");
        codePages[1].AddToken(0x13, "BusinessPhoneNumber");
        codePages[1].AddToken(0x14, "CarPhoneNumber");
        codePages[1].AddToken(0x15, "Categories");
        codePages[1].AddToken(0x16, "Category");
        codePages[1].AddToken(0x17, "Children");
        codePages[1].AddToken(0x18, "Child");
        codePages[1].AddToken(0x19, "CompanyName");
        codePages[1].AddToken(0x1A, "Department");
        codePages[1].AddToken(0x1B, "Email1Address");
        codePages[1].AddToken(0x1C, "Email2Address");
        codePages[1].AddToken(0x1D, "Email3Address");
        codePages[1].AddToken(0x1E, "FileAs");
        codePages[1].AddToken(0x1F, "FirstName");
        codePages[1].AddToken(0x20, "Home2PhoneNumber");
        codePages[1].AddToken(0x21, "HomeAddressCity");
        codePages[1].AddToken(0x22, "HomeAddressCountry");
        codePages[1].AddToken(0x23, "HomeAddressPostalCode");
        codePages[1].AddToken(0x24, "HomeAddressState");
        codePages[1].AddToken(0x25, "HomeAddressStreet");
        codePages[1].AddToken(0x26, "HomeFaxNumber");
        codePages[1].AddToken(0x27, "HomePhoneNumber");
        codePages[1].AddToken(0x28, "JobTitle");
        codePages[1].AddToken(0x29, "LastName");
        codePages[1].AddToken(0x2A, "MiddleName");
        codePages[1].AddToken(0x2B, "MobilePhoneNumber");
        codePages[1].AddToken(0x2C, "OfficeLocation");
        codePages[1].AddToken(0x2D, "OtherAddressCity");
        codePages[1].AddToken(0x2E, "OtherAddressCountry");
        codePages[1].AddToken(0x2F, "OtherAddressPostalCode");
        codePages[1].AddToken(0x30, "OtherAddressState");
        codePages[1].AddToken(0x31, "OtherAddressStreet");
        codePages[1].AddToken(0x32, "PagerNumber");
        codePages[1].AddToken(0x33, "RadioPhoneNumber");
        codePages[1].AddToken(0x34, "Spouse");
        codePages[1].AddToken(0x35, "Suffix");
        codePages[1].AddToken(0x36, "Title");
        codePages[1].AddToken(0x37, "WebPage");
        codePages[1].AddToken(0x38, "YomiCompanyName");
        codePages[1].AddToken(0x39, "YomiFirstName");
        codePages[1].AddToken(0x3A, "YomiLastName");
        codePages[1].AddToken(0x3C, "Picture");
        codePages[1].AddToken(0x3D, "Alias");
        codePages[1].AddToken(0x3E, "WeightedRank");

        // Code Page 2: Email
        codePages[2] = new CodePage { Namespace = "Email", Xmlns = "email" };
        codePages[2].AddToken(0x0F, "DateReceived");
        codePages[2].AddToken(0x11, "DisplayTo");
        codePages[2].AddToken(0x12, "Importance");
        codePages[2].AddToken(0x13, "MessageClass");
        codePages[2].AddToken(0x14, "Subject");
        codePages[2].AddToken(0x15, "Read");
        codePages[2].AddToken(0x16, "To");
        codePages[2].AddToken(0x17, "Cc");
        codePages[2].AddToken(0x18, "From");
        codePages[2].AddToken(0x19, "ReplyTo");
        codePages[2].AddToken(0x1A, "AllDayEvent");
        codePages[2].AddToken(0x1B, "Categories");
        codePages[2].AddToken(0x1C, "Category");
        codePages[2].AddToken(0x1D, "DtStamp");
        codePages[2].AddToken(0x1E, "EndTime");
        codePages[2].AddToken(0x1F, "InstanceType");
        codePages[2].AddToken(0x20, "BusyStatus");
        codePages[2].AddToken(0x21, "Location");
        codePages[2].AddToken(0x22, "MeetingRequest");
        codePages[2].AddToken(0x23, "Organizer");
        codePages[2].AddToken(0x24, "RecurrenceId");
        codePages[2].AddToken(0x25, "Reminder");
        codePages[2].AddToken(0x26, "ResponseRequested");
        codePages[2].AddToken(0x27, "Recurrences");
        codePages[2].AddToken(0x28, "Recurrence");
        codePages[2].AddToken(0x29, "Type");
        codePages[2].AddToken(0x2A, "Until");
        codePages[2].AddToken(0x2B, "Occurrences");
        codePages[2].AddToken(0x2C, "Interval");
        codePages[2].AddToken(0x2D, "DayOfWeek");
        codePages[2].AddToken(0x2E, "DayOfMonth");
        codePages[2].AddToken(0x2F, "WeekOfMonth");
        codePages[2].AddToken(0x30, "MonthOfYear");
        codePages[2].AddToken(0x31, "StartTime");
        codePages[2].AddToken(0x32, "Sensitivity");
        codePages[2].AddToken(0x33, "TimeZone");
        codePages[2].AddToken(0x34, "GlobalObjId");
        codePages[2].AddToken(0x35, "ThreadTopic");
        codePages[2].AddToken(0x39, "InternetCPID");
        codePages[2].AddToken(0x3A, "Flag");
        codePages[2].AddToken(0x3B, "Status");
        codePages[2].AddToken(0x3C, "ContentClass");
        codePages[2].AddToken(0x3D, "FlagType");
        codePages[2].AddToken(0x3E, "CompleteTime");
        codePages[2].AddToken(0x3F, "DisallowNewTimeProposal");

        // Code Page 3: AirNotify
        codePages[3] = new CodePage { Namespace = string.Empty, Xmlns = string.Empty };

        // Code Page 4: Calendar
        codePages[4] = new CodePage { Namespace = "Calendar", Xmlns = "calendar" };

        codePages[4].AddToken(0x05, "Timezone");
        codePages[4].AddToken(0x06, "AllDayEvent");
        codePages[4].AddToken(0x07, "Attendees");
        codePages[4].AddToken(0x08, "Attendee");
        codePages[4].AddToken(0x09, "Email");
        codePages[4].AddToken(0x0A, "Name");
        codePages[4].AddToken(0x0D, "BusyStatus");
        codePages[4].AddToken(0x0E, "Categories");
        codePages[4].AddToken(0x0F, "Category");
        codePages[4].AddToken(0x11, "DtStamp");
        codePages[4].AddToken(0x12, "EndTime");
        codePages[4].AddToken(0x13, "Exception");
        codePages[4].AddToken(0x14, "Exceptions");
        codePages[4].AddToken(0x15, "Deleted");
        codePages[4].AddToken(0x16, "ExceptionStartTime");
        codePages[4].AddToken(0x17, "Location");
        codePages[4].AddToken(0x18, "MeetingStatus");
        codePages[4].AddToken(0x19, "OrganizerEmail");
        codePages[4].AddToken(0x1A, "OrganizerName");
        codePages[4].AddToken(0x1B, "Recurrence");
        codePages[4].AddToken(0x1C, "Type");
        codePages[4].AddToken(0x1D, "Until");
        codePages[4].AddToken(0x1E, "Occurrences");
        codePages[4].AddToken(0x1F, "Interval");
        codePages[4].AddToken(0x20, "DayOfWeek");
        codePages[4].AddToken(0x21, "DayOfMonth");
        codePages[4].AddToken(0x22, "WeekOfMonth");
        codePages[4].AddToken(0x23, "MonthOfYear");
        codePages[4].AddToken(0x24, "Reminder");
        codePages[4].AddToken(0x25, "Sensitivity");
        codePages[4].AddToken(0x26, "Subject");
        codePages[4].AddToken(0x27, "StartTime");
        codePages[4].AddToken(0x28, "UID");
        codePages[4].AddToken(0x29, "AttendeeStatus");
        codePages[4].AddToken(0x2A, "AttendeeType");
        codePages[4].AddToken(0x33, "DisallowNewTimeProposal");
        codePages[4].AddToken(0x34, "ResponseRequested");
        codePages[4].AddToken(0x35, "AppointmentReplyTime");
        codePages[4].AddToken(0x36, "ResponseType");
        codePages[4].AddToken(0x37, "CalendarType");
        codePages[4].AddToken(0x38, "IsLeapMonth");
        codePages[4].AddToken(0x39, "FirstDayOfWeek");
        codePages[4].AddToken(0x3A, "OnlineMeetingConfLink");
        codePages[4].AddToken(0x3B, "OnlineMeetingExternalLink");
        codePages[4].AddToken(0x3C, "ClientUid");

        // Code Page 5: Move
        codePages[5] = new CodePage { Namespace = "Move", Xmlns = "move" };

        codePages[5].AddToken(0x05, "MoveItems");
        codePages[5].AddToken(0x06, "Move");
        codePages[5].AddToken(0x07, "SrcMsgId");
        codePages[5].AddToken(0x08, "SrcFldId");
        codePages[5].AddToken(0x09, "DstFldId");
        codePages[5].AddToken(0x0A, "Response");
        codePages[5].AddToken(0x0B, "Status");
        codePages[5].AddToken(0x0C, "DstMsgId");

        // Code Page 6: GetItemEstimate
        codePages[6] = new CodePage { Namespace = "GetItemEstimate", Xmlns = "getitemestimate" };

        codePages[6].AddToken(0x05, "GetItemEstimate");
        codePages[6].AddToken(0x07, "Collections");
        codePages[6].AddToken(0x08, "Collection");
        codePages[6].AddToken(0x09, "Class");
        codePages[6].AddToken(0x0A, "CollectionId");
        codePages[6].AddToken(0x0C, "Estimate");
        codePages[6].AddToken(0x0D, "Response");
        codePages[6].AddToken(0x0E, "Status");

        // Code Page 7: FolderHierarchy
        codePages[7] = new CodePage { Namespace = "FolderHierarchy", Xmlns = "folderhierarchy" };

        codePages[7].AddToken(0x05, "Folders");
        codePages[7].AddToken(0x06, "Folder");
        codePages[7].AddToken(0x07, "DisplayName");
        codePages[7].AddToken(0x08, "ServerId");
        codePages[7].AddToken(0x09, "ParentId");
        codePages[7].AddToken(0x0A, "Type");
        codePages[7].AddToken(0x0C, "Status");
        codePages[7].AddToken(0x0E, "Changes");
        codePages[7].AddToken(0x0F, "Add");
        codePages[7].AddToken(0x10, "Delete");
        codePages[7].AddToken(0x11, "Update");
        codePages[7].AddToken(0x12, "SyncKey");
        codePages[7].AddToken(0x13, "FolderCreate");
        codePages[7].AddToken(0x14, "FolderDelete");
        codePages[7].AddToken(0x15, "FolderUpdate");
        codePages[7].AddToken(0x16, "FolderSync");
        codePages[7].AddToken(0x17, "Count");

        // Code Page 8: MeetingResponse
        codePages[8] = new CodePage { Namespace = "MeetingResponse", Xmlns = "meetingresponse" };

        codePages[8].AddToken(0x05, "CalendarId");
        codePages[8].AddToken(0x06, "CollectionId");
        codePages[8].AddToken(0x07, "MeetingResponse");
        codePages[8].AddToken(0x08, "RequestId");
        codePages[8].AddToken(0x09, "Request");
        codePages[8].AddToken(0x0A, "Result");
        codePages[8].AddToken(0x0B, "Status");
        codePages[8].AddToken(0x0C, "UserResponse");
        codePages[8].AddToken(0x0E, "InstanceId");
        codePages[8].AddToken(0x10, "ProposedStartTime");
        codePages[8].AddToken(0x11, "ProposedEndTime");
        codePages[8].AddToken(0x12, "SendResponse");

        // Code Page 9: Tasks
        codePages[9] = new CodePage { Namespace = "Tasks", Xmlns = "tasks" };

        codePages[9].AddToken(0x08, "Categories");
        codePages[9].AddToken(0x09, "Category");
        codePages[9].AddToken(0x0A, "Complete");
        codePages[9].AddToken(0x0B, "DateCompleted");
        codePages[9].AddToken(0x0C, "DueDate");
        codePages[9].AddToken(0x0D, "UtcDueDate");
        codePages[9].AddToken(0x0E, "Importance");
        codePages[9].AddToken(0x0F, "Recurrence");
        codePages[9].AddToken(0x10, "Type");
        codePages[9].AddToken(0x11, "Start");
        codePages[9].AddToken(0x12, "Until");
        codePages[9].AddToken(0x13, "Occurrences");
        codePages[9].AddToken(0x14, "Interval");
        codePages[9].AddToken(0x15, "DayOfMonth");
        codePages[9].AddToken(0x16, "DayOfWeek");
        codePages[9].AddToken(0x17, "WeekOfMonth");
        codePages[9].AddToken(0x18, "MonthOfYear");
        codePages[9].AddToken(0x19, "Regenerate");
        codePages[9].AddToken(0x1A, "DeadOccur");
        codePages[9].AddToken(0x1B, "ReminderSet");
        codePages[9].AddToken(0x1C, "ReminderTime");
        codePages[9].AddToken(0x1D, "Sensitivity");
        codePages[9].AddToken(0x1E, "StartDate");
        codePages[9].AddToken(0x1F, "UtcStartDate");
        codePages[9].AddToken(0x20, "Subject");
        codePages[9].AddToken(0x22, "OrdinalDate");
        codePages[9].AddToken(0x23, "SubOrdinalDate");
        codePages[9].AddToken(0x24, "CalendarType");
        codePages[9].AddToken(0x25, "IsLeapMonth");
        codePages[9].AddToken(0x26, "FirstDayOfWeek");

        // Code Page 10: ResolveRecipients
        codePages[10] = new CodePage { Namespace = "ResolveRecipients", Xmlns = "resolverecipients" };

        codePages[10].AddToken(0x05, "ResolveRecipients");
        codePages[10].AddToken(0x06, "Response");
        codePages[10].AddToken(0x07, "Status");
        codePages[10].AddToken(0x08, "Type");
        codePages[10].AddToken(0x09, "Recipient");
        codePages[10].AddToken(0x0A, "DisplayName");
        codePages[10].AddToken(0x0B, "EmailAddress");
        codePages[10].AddToken(0x0C, "Certificates");
        codePages[10].AddToken(0x0D, "Certificate");
        codePages[10].AddToken(0x0E, "MiniCertificate");
        codePages[10].AddToken(0x0F, "Options");
        codePages[10].AddToken(0x10, "To");
        codePages[10].AddToken(0x11, "CertificateRetrieval");
        codePages[10].AddToken(0x12, "RecipientCount");
        codePages[10].AddToken(0x13, "MaxCertificates");
        codePages[10].AddToken(0x14, "MaxAmbiguousRecipients");
        codePages[10].AddToken(0x15, "CertificateCount");
        codePages[10].AddToken(0x16, "Availability");
        codePages[10].AddToken(0x17, "StartTime");
        codePages[10].AddToken(0x18, "EndTime");
        codePages[10].AddToken(0x19, "MergedFreeBusy");
        codePages[10].AddToken(0x1A, "Picture");
        codePages[10].AddToken(0x1B, "MaxSize");
        codePages[10].AddToken(0x1C, "Data");
        codePages[10].AddToken(0x1D, "MaxPictures");

        // Code Page 11: ValidateCert
        codePages[11] = new CodePage { Namespace = "ValidateCert", Xmlns = "ValidateCert" };

        codePages[11].AddToken(0x05, "ValidateCert");
        codePages[11].AddToken(0x06, "Certificates");
        codePages[11].AddToken(0x07, "Certificate");
        codePages[11].AddToken(0x08, "CertificateChain");
        codePages[11].AddToken(0x09, "CheckCrl");
        codePages[11].AddToken(0x0A, "Status");

        // Code Page 12: Contacts2
        codePages[12] = new CodePage { Namespace = "Contacts2", Xmlns = "contacts2" };

        codePages[12].AddToken(0x05, "CustomerId");
        codePages[12].AddToken(0x06, "GovernmentId");
        codePages[12].AddToken(0x07, "IMAddress");
        codePages[12].AddToken(0x08, "IMAddress2");
        codePages[12].AddToken(0x09, "IMAddress3");
        codePages[12].AddToken(0x0A, "ManagerName");
        codePages[12].AddToken(0x0B, "CompanyMainPhone");
        codePages[12].AddToken(0x0C, "AccountName");
        codePages[12].AddToken(0x0D, "NickName");
        codePages[12].AddToken(0x0E, "MMS");

        // Code Page 13: Ping
        codePages[13] = new CodePage { Namespace = "Ping", Xmlns = "ping" };

        codePages[13].AddToken(0x05, "Ping");
        codePages[13].AddToken(0x07, "Status");
        codePages[13].AddToken(0x08, "HeartbeatInterval");
        codePages[13].AddToken(0x09, "Folders");
        codePages[13].AddToken(0x0A, "Folder");
        codePages[13].AddToken(0x0B, "Id");
        codePages[13].AddToken(0x0C, "Class");
        codePages[13].AddToken(0x0D, "MaxFolders");

        // Code Page 14: Provision
        codePages[14] = new CodePage { Namespace = "Provision", Xmlns = "provision" };

        codePages[14].AddToken(0x05, "Provision");
        codePages[14].AddToken(0x06, "Policies");
        codePages[14].AddToken(0x07, "Policy");
        codePages[14].AddToken(0x08, "PolicyType");
        codePages[14].AddToken(0x09, "PolicyKey");
        codePages[14].AddToken(0x0A, "Data");
        codePages[14].AddToken(0x0B, "Status");
        codePages[14].AddToken(0x0C, "RemoteWipe");
        codePages[14].AddToken(0x0D, "EASProvisionDoc");
        codePages[14].AddToken(0x0E, "DevicePasswordEnabled");
        codePages[14].AddToken(0x0F, "AlphanumericDevicePasswordRequired");
        codePages[14].AddToken(0x10, "RequireStorageCardEncryption");
        codePages[14].AddToken(0x11, "PasswordRecoveryEnabled");
        codePages[14].AddToken(0x13, "AttachmentsEnabled");
        codePages[14].AddToken(0x14, "MinDevicePasswordLength");
        codePages[14].AddToken(0x15, "MaxInactivityTimeDeviceLock");
        codePages[14].AddToken(0x16, "MaxDevicePasswordFailedAttempts");
        codePages[14].AddToken(0x17, "MaxAttachmentSize");
        codePages[14].AddToken(0x18, "AllowSimpleDevicePassword");
        codePages[14].AddToken(0x19, "DevicePasswordExpiration");
        codePages[14].AddToken(0x1A, "DevicePasswordHistory");
        codePages[14].AddToken(0x1B, "AllowStorageCard");
        codePages[14].AddToken(0x1C, "AllowCamera");
        codePages[14].AddToken(0x1D, "RequireDeviceEncryption");
        codePages[14].AddToken(0x1E, "AllowUnsignedApplications");
        codePages[14].AddToken(0x1F, "AllowUnsignedInstallationPackages");
        codePages[14].AddToken(0x20, "MinDevicePasswordComplexCharacters");
        codePages[14].AddToken(0x21, "AllowWiFi");
        codePages[14].AddToken(0x22, "AllowTextMessaging");
        codePages[14].AddToken(0x23, "AllowPOPIMAPEmail");
        codePages[14].AddToken(0x24, "AllowBluetooth");
        codePages[14].AddToken(0x25, "AllowIrDA");
        codePages[14].AddToken(0x26, "RequireManualSyncWhenRoaming");
        codePages[14].AddToken(0x27, "AllowDesktopSync");
        codePages[14].AddToken(0x28, "MaxCalendarAgeFilter");
        codePages[14].AddToken(0x29, "AllowHTMLEmail");
        codePages[14].AddToken(0x2A, "MaxEmailAgeFilter");
        codePages[14].AddToken(0x2B, "MaxEmailBodyTruncationSize");
        codePages[14].AddToken(0x2C, "MaxEmailHTMLBodyTruncationSize");
        codePages[14].AddToken(0x2D, "RequireSignedSMIMEMessages");
        codePages[14].AddToken(0x2E, "RequireEncryptedSMIMEMessages");
        codePages[14].AddToken(0x2F, "RequireSignedSMIMEAlgorithm");
        codePages[14].AddToken(0x30, "RequireEncryptionSMIMEAlgorithm");
        codePages[14].AddToken(0x31, "AllowSMIMEEncryptionAlgorithmNegotiation");
        codePages[14].AddToken(0x32, "AllowSMIMESoftCerts");
        codePages[14].AddToken(0x33, "AllowBrowser");
        codePages[14].AddToken(0x34, "AllowConsumerEmail");
        codePages[14].AddToken(0x35, "AllowRemoteDesktop");
        codePages[14].AddToken(0x36, "AllowInternetSharing");
        codePages[14].AddToken(0x37, "UnapprovedInROMApplicationList");
        codePages[14].AddToken(0x38, "ApplicationName");
        codePages[14].AddToken(0x39, "ApprovedApplicationList");
        codePages[14].AddToken(0x3A, "Hash");
        codePages[14].AddToken(0x3B, "AccountOnlyRemoteWipe");

        // Code Page 15: Search
        codePages[15] = new CodePage { Namespace = "Search", Xmlns = "search" };

        codePages[15].AddToken(0x05, "Search");
        codePages[15].AddToken(0x07, "Store");
        codePages[15].AddToken(0x08, "Name");
        codePages[15].AddToken(0x09, "Query");
        codePages[15].AddToken(0x0A, "Options");
        codePages[15].AddToken(0x0B, "Range");
        codePages[15].AddToken(0x0C, "Status");
        codePages[15].AddToken(0x0D, "Response");
        codePages[15].AddToken(0x0E, "Result");
        codePages[15].AddToken(0x0F, "Properties");
        codePages[15].AddToken(0x10, "Total");
        codePages[15].AddToken(0x11, "EqualTo");
        codePages[15].AddToken(0x12, "Value");
        codePages[15].AddToken(0x13, "And");
        codePages[15].AddToken(0x14, "Or");
        codePages[15].AddToken(0x15, "FreeText");
        codePages[15].AddToken(0x17, "DeepTraversal");
        codePages[15].AddToken(0x18, "LongId");
        codePages[15].AddToken(0x19, "RebuildResults");
        codePages[15].AddToken(0x1A, "LessThan");
        codePages[15].AddToken(0x1B, "GreaterThan");
        codePages[15].AddToken(0x1E, "UserName");
        codePages[15].AddToken(0x1F, "Password");
        codePages[15].AddToken(0x20, "ConversationId");
        codePages[15].AddToken(0x21, "Picture");
        codePages[15].AddToken(0x22, "MaxSize");
        codePages[15].AddToken(0x23, "MaxPictures");

        // Code Page 16: GAL
        codePages[16] = new CodePage { Namespace = "GAL", Xmlns = "gal" };

        codePages[16].AddToken(0x05, "DisplayName");
        codePages[16].AddToken(0x06, "Phone");
        codePages[16].AddToken(0x07, "Office");
        codePages[16].AddToken(0x08, "Title");
        codePages[16].AddToken(0x09, "Company");
        codePages[16].AddToken(0x0A, "Alias");
        codePages[16].AddToken(0x0B, "FirstName");
        codePages[16].AddToken(0x0C, "LastName");
        codePages[16].AddToken(0x0D, "HomePhone");
        codePages[16].AddToken(0x0E, "MobilePhone");
        codePages[16].AddToken(0x0F, "EmailAddress");
        codePages[16].AddToken(0x10, "Picture");
        codePages[16].AddToken(0x11, "Status");
        codePages[16].AddToken(0x12, "Data");

        // Code Page 17: AirSyncBase
        codePages[17] = new CodePage { Namespace = "AirSyncBase", Xmlns = "airsyncbase" };

        codePages[17].AddToken(0x05, "BodyPreference");
        codePages[17].AddToken(0x06, "Type");
        codePages[17].AddToken(0x07, "TruncationSize");
        codePages[17].AddToken(0x08, "AllOrNone");
        codePages[17].AddToken(0x0A, "Body");
        codePages[17].AddToken(0x0B, "Data");
        codePages[17].AddToken(0x0C, "EstimatedDataSize");
        codePages[17].AddToken(0x0D, "Truncated");
        codePages[17].AddToken(0x0E, "Attachments");
        codePages[17].AddToken(0x0F, "Attachment");
        codePages[17].AddToken(0x10, "DisplayName");
        codePages[17].AddToken(0x11, "FileReference");
        codePages[17].AddToken(0x12, "Method");
        codePages[17].AddToken(0x13, "ContentId");
        codePages[17].AddToken(0x14, "ContentLocation");
        codePages[17].AddToken(0x15, "IsInline");
        codePages[17].AddToken(0x16, "NativeBodyType");
        codePages[17].AddToken(0x17, "ContentType");
        codePages[17].AddToken(0x18, "Preview");
        codePages[17].AddToken(0x19, "BodyPartPreference");
        codePages[17].AddToken(0x1A, "BodyPart");
        codePages[17].AddToken(0x1B, "Status");
        codePages[17].AddToken(0x1C, "Add");
        codePages[17].AddToken(0x1D, "Delete");
        codePages[17].AddToken(0x1E, "ClientId");
        codePages[17].AddToken(0x1F, "Content");
        codePages[17].AddToken(0x20, "Location");
        codePages[17].AddToken(0x21, "Annotation");
        codePages[17].AddToken(0x22, "Street");
        codePages[17].AddToken(0x23, "City");
        codePages[17].AddToken(0x24, "State");
        codePages[17].AddToken(0x25, "Country");
        codePages[17].AddToken(0x26, "PostalCode");
        codePages[17].AddToken(0x27, "Latitude");
        codePages[17].AddToken(0x28, "Longitude");
        codePages[17].AddToken(0x29, "Accuracy");
        codePages[17].AddToken(0x2A, "Altitude");
        codePages[17].AddToken(0x2B, "AltitudeAccuracy");
        codePages[17].AddToken(0x2C, "LocationUri");
        codePages[17].AddToken(0x2D, "InstanceId");


        // Code Page 18: Settings
        codePages[18] = new CodePage { Namespace = "Settings", Xmlns = "settings" };

        codePages[18].AddToken(0x05, "Settings");
        codePages[18].AddToken(0x06, "Status");
        codePages[18].AddToken(0x07, "Get");
        codePages[18].AddToken(0x08, "Set");
        codePages[18].AddToken(0x09, "Oof");
        codePages[18].AddToken(0x0A, "OofState");
        codePages[18].AddToken(0x0B, "StartTime");
        codePages[18].AddToken(0x0C, "EndTime");
        codePages[18].AddToken(0x0D, "OofMessage");
        codePages[18].AddToken(0x0E, "AppliesToInternal");
        codePages[18].AddToken(0x0F, "AppliesToExternalKnown");
        codePages[18].AddToken(0x10, "AppliesToExternalUnknown");
        codePages[18].AddToken(0x11, "Enabled");
        codePages[18].AddToken(0x12, "ReplyMessage");
        codePages[18].AddToken(0x13, "BodyType");
        codePages[18].AddToken(0x14, "DevicePassword");
        codePages[18].AddToken(0x15, "Password");
        codePages[18].AddToken(0x16, "DeviceInformation");
        codePages[18].AddToken(0x17, "Model");
        codePages[18].AddToken(0x18, "IMEI");
        codePages[18].AddToken(0x19, "FriendlyName");
        codePages[18].AddToken(0x1A, "OS");
        codePages[18].AddToken(0x1B, "OSLanguage");
        codePages[18].AddToken(0x1C, "PhoneNumber");
        codePages[18].AddToken(0x1D, "UserInformation");
        codePages[18].AddToken(0x1E, "EmailAddresses");
        codePages[18].AddToken(0x1F, "SMTPAddress");
        codePages[18].AddToken(0x20, "UserAgent");
        codePages[18].AddToken(0x21, "EnableOutboundSMS");
        codePages[18].AddToken(0x22, "MobileOperator");
        codePages[18].AddToken(0x23, "PrimarySmtpAddress");
        codePages[18].AddToken(0x24, "Accounts");
        codePages[18].AddToken(0x25, "Account");
        codePages[18].AddToken(0x26, "AccountId");
        codePages[18].AddToken(0x27, "AccountName");
        codePages[18].AddToken(0x28, "UserDisplayName");
        codePages[18].AddToken(0x29, "SendDisabled");
        codePages[18].AddToken(0x2B, "RightsManagementInformation");

        // Code Page 19: DocumentLibrary
        codePages[19] = new CodePage { Namespace = "DocumentLibrary", Xmlns = "documentlibrary" };

        codePages[19].AddToken(0x05, "LinkId");
        codePages[19].AddToken(0x06, "DisplayName");
        codePages[19].AddToken(0x07, "IsFolder");
        codePages[19].AddToken(0x08, "CreationDate");
        codePages[19].AddToken(0x09, "LastModifiedDate");
        codePages[19].AddToken(0x0A, "IsHidden");
        codePages[19].AddToken(0x0B, "ContentLength");
        codePages[19].AddToken(0x0C, "ContentType");

        // Code Page 20: ItemOperations
        codePages[20] = new CodePage { Namespace = "ItemOperations", Xmlns = "itemoperations" };

        codePages[20].AddToken(0x05, "ItemOperations");
        codePages[20].AddToken(0x06, "Fetch");
        codePages[20].AddToken(0x07, "Store");
        codePages[20].AddToken(0x08, "Options");
        codePages[20].AddToken(0x09, "Range");
        codePages[20].AddToken(0x0A, "Total");
        codePages[20].AddToken(0x0B, "Properties");
        codePages[20].AddToken(0x0C, "Data");
        codePages[20].AddToken(0x0D, "Status");
        codePages[20].AddToken(0x0E, "Response");
        codePages[20].AddToken(0x0F, "Version");
        codePages[20].AddToken(0x10, "Schema");
        codePages[20].AddToken(0x11, "Part");
        codePages[20].AddToken(0x12, "EmptyFolderContents");
        codePages[20].AddToken(0x13, "DeleteSubFolders");
        codePages[20].AddToken(0x14, "UserName");
        codePages[20].AddToken(0x15, "Password");
        codePages[20].AddToken(0x16, "Move");
        codePages[20].AddToken(0x17, "DstFldId");
        codePages[20].AddToken(0x18, "ConversationId");
        codePages[20].AddToken(0x19, "MoveAlways");

        // Code Page 21: ComposeMail
        codePages[21] = new CodePage { Namespace = "ComposeMail", Xmlns = "composemail" };

        codePages[21].AddToken(0x05, "SendMail");
        codePages[21].AddToken(0x06, "SmartForward");
        codePages[21].AddToken(0x07, "SmartReply");
        codePages[21].AddToken(0x08, "SaveInSentItems");
        codePages[21].AddToken(0x09, "ReplaceMime");
        codePages[21].AddToken(0x0B, "Source");
        codePages[21].AddToken(0x0C, "FolderId");
        codePages[21].AddToken(0x0D, "ItemId");
        codePages[21].AddToken(0x0E, "LongId");
        codePages[21].AddToken(0x0F, "InstanceId");
        codePages[21].AddToken(0x10, "Mime");
        codePages[21].AddToken(0x11, "ClientId");
        codePages[21].AddToken(0x12, "Status");
        codePages[21].AddToken(0x13, "AccountId");
        codePages[21].AddToken(0x15, "Forwardees");
        codePages[21].AddToken(0x16, "Forwardee");
        codePages[21].AddToken(0x17, "ForwardeeName");
        codePages[21].AddToken(0x18, "ForwardeeEmail");

        // Code Page 22: Email2
        codePages[22] = new CodePage { Namespace = "Email2", Xmlns = "email2" };

        codePages[22].AddToken(0x05, "UmCallerID");
        codePages[22].AddToken(0x06, "UmUserNotes");
        codePages[22].AddToken(0x07, "UmAttDuration");
        codePages[22].AddToken(0x08, "UmAttOrder");
        codePages[22].AddToken(0x09, "ConversationId");
        codePages[22].AddToken(0x0A, "ConversationIndex");
        codePages[22].AddToken(0x0B, "LastVerbExecuted");
        codePages[22].AddToken(0x0C, "LastVerbExecutionTime");
        codePages[22].AddToken(0x0D, "ReceivedAsBcc");
        codePages[22].AddToken(0x0E, "Sender");
        codePages[22].AddToken(0x0F, "CalendarType");
        codePages[22].AddToken(0x10, "IsLeapMonth");
        codePages[22].AddToken(0x11, "AccountId");
        codePages[22].AddToken(0x12, "FirstDayOfWeek");
        codePages[22].AddToken(0x13, "MeetingMessageType");
        codePages[22].AddToken(0x15, "IsDraft");
        codePages[22].AddToken(0x16, "Bcc");
        codePages[22].AddToken(0x17, "Send");

        // Code Page 23: Notes
        codePages[23] = new CodePage { Namespace = "Notes", Xmlns = "notes" };

        codePages[23].AddToken(0x05, "Subject");
        codePages[23].AddToken(0x06, "MessageClass");
        codePages[23].AddToken(0x07, "LastModifiedDate");
        codePages[23].AddToken(0x08, "Categories");
        codePages[23].AddToken(0x09, "Category");

        // Code Page 24: RightsManagement
        codePages[24] = new CodePage { Namespace = "RightsManagement", Xmlns = "rightsmanagement" };

        codePages[24].AddToken(0x05, "RightsManagementSupport");
        codePages[24].AddToken(0x06, "RightsManagementTemplates");
        codePages[24].AddToken(0x07, "RightsManagementTemplate");
        codePages[24].AddToken(0x08, "RightsManagementLicense");
        codePages[24].AddToken(0x09, "EditAllowed");
        codePages[24].AddToken(0x0A, "ReplyAllowed");
        codePages[24].AddToken(0x0B, "ReplyAllAllowed");
        codePages[24].AddToken(0x0C, "ForwardAllowed");
        codePages[24].AddToken(0x0D, "ModifyRecipientsAllowed");
        codePages[24].AddToken(0x0E, "ExtractAllowed");
        codePages[24].AddToken(0x0F, "PrintAllowed");
        codePages[24].AddToken(0x10, "ExportAllowed");
        codePages[24].AddToken(0x11, "ProgrammaticAccessAllowed");
        codePages[24].AddToken(0x12, "Owner");
        codePages[24].AddToken(0x13, "ContentExpiryDate");
        codePages[24].AddToken(0x14, "TemplateID");
        codePages[24].AddToken(0x15, "TemplateName");
        codePages[24].AddToken(0x16, "TemplateDescription");
        codePages[24].AddToken(0x17, "ContentOwner");
        codePages[24].AddToken(0x18, "RemoveRightsManagementProtection");

        //Code page 25: Find
        codePages[25] = new CodePage { Namespace = "Find", Xmlns = "Find" };
        codePages[25].AddToken(0x05, "Find");
        codePages[25].AddToken(0x06, "SearchId");
        codePages[25].AddToken(0x07, "ExecuteSearch");
        codePages[25].AddToken(0x08, "MailBoxSearchCriterion");
        codePages[25].AddToken(0x09, "Query");
        codePages[25].AddToken(0x0A, "Status");
        codePages[25].AddToken(0x0B, "FreeText");
        codePages[25].AddToken(0x0C, "Options");
        codePages[25].AddToken(0x0D, "Range");
        codePages[25].AddToken(0x0E, "DeepTraversal");
        codePages[25].AddToken(0x11, "Response");
        codePages[25].AddToken(0x12, "Result");
        codePages[25].AddToken(0x13, "Properties");
        codePages[25].AddToken(0x14, "Preview");
        codePages[25].AddToken(0x15, "HasAttachments");
        codePages[25].AddToken(0x16, "Total");
        codePages[25].AddToken(0x17, "DisplayCc");
        codePages[25].AddToken(0x18, "DisplayBcc");
        codePages[25].AddToken(0x19, "GALSearchCriterion");
        codePages[25].AddToken(0x20, "MaxPictures");
        codePages[25].AddToken(0x21, "MaxSize");
        codePages[25].AddToken(0x22, "Picture");

    }

    /// <summary>
    /// Gets the DataCollection in encoding process
    /// </summary>
    public Dictionary<string, int> EncodeDataCollection
    {
        get { return encodeDataCollection; }
    }

    /// <summary>
    /// Gets the DataCollection in decoding process
    /// </summary>
    public Dictionary<string, int> DecodeDataCollection
    {
        get { return decodeDataCollection; }
    }

    /// <summary>
    /// Loads byte array and decode to xml string.
    /// </summary>
    /// <param name="byteWBXML">The bytes to be decoded</param>
    /// <returns>The decoded xml string.</returns>
    public string DecodeToXml(byte[] byteWBXML)
    {
        xmlDoc = new XmlDocument();

        var bytes = new ByteQueue(byteWBXML);

        // Remove the version from bytes
        bytes.Dequeue();

        // Remove public identifier from bytes
        bytes.DequeueMultibyteInt();

        // Gets the Character set from bytes
        var charset = bytes.DequeueMultibyteInt();
        if (charset != 0x6A)
        {
            return string.Empty;
        }

        // String table length. MS-ASWBXML does not use string tables, it should be 0.
        var stringTableLength = bytes.DequeueMultibyteInt();
        site.Assert.AreEqual<int>(0, stringTableLength, "MS-ASWBXML does not use string tables, therefore String table length should be 0.");

        // Initializes the DecodeDataCollection and begins to record
        if (null == decodeDataCollection)
        {
            decodeDataCollection = new Dictionary<string, int>();
        }
        else
        {
            decodeDataCollection.Clear();
        }

        // Adds the declaration
        var xmlDec = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
        xmlDoc.InsertBefore(xmlDec, null);

        XmlNode currentNode = xmlDoc;

        while (bytes.Count > 0)
        {
            var currentByte = bytes.Dequeue();

            switch ((GlobalTokens)currentByte)
            {
                case GlobalTokens.SWITCH_PAGE:
                    var newCodePage = (int)bytes.Dequeue();
                    if (newCodePage >= 0 && newCodePage < 26)
                    {
                        currentCodePage = newCodePage;
                    }
                    else
                    {
                        site.Assert.Fail("Code page value which defined in MS-ASWBXML should be between 0-25, the actual value is : {0}.", newCodePage);
                    }

                    break;
                case GlobalTokens.END:
                    if (currentNode.ParentNode != null)
                    {
                        currentNode = currentNode.ParentNode;
                    }
                    else
                    {
                        return string.Empty;
                    }

                    break;
                case GlobalTokens.OPAQUE:
                    var cdataLength = bytes.DequeueMultibyteInt();
                    XmlCDataSection newOpaqueNode;
                    if (currentNode.Name == "ConversationId"
                        || currentNode.Name == "ConversationIndex")
                    {
                        newOpaqueNode = xmlDoc.CreateCDataSection(bytes.DequeueBase64String(cdataLength));
                    }
                    else
                    {
                        newOpaqueNode = xmlDoc.CreateCDataSection(bytes.DequeueString(cdataLength));
                    }

                    currentNode.AppendChild(newOpaqueNode);
                    break;
                case GlobalTokens.STR_I:
                    XmlNode newTextNode = xmlDoc.CreateTextNode(bytes.DequeueString());
                    currentNode.AppendChild(newTextNode);
                    break;

                case GlobalTokens.ENTITY:
                case GlobalTokens.EXT_0:
                case GlobalTokens.EXT_1:
                case GlobalTokens.EXT_2:
                case GlobalTokens.EXT_I_0:
                case GlobalTokens.EXT_I_1:
                case GlobalTokens.EXT_I_2:
                case GlobalTokens.EXT_T_0:
                case GlobalTokens.EXT_T_1:
                case GlobalTokens.EXT_T_2:
                case GlobalTokens.LITERAL:
                case GlobalTokens.LITERAL_A:
                case GlobalTokens.LITERAL_AC:
                case GlobalTokens.LITERAL_C:
                case GlobalTokens.PI:
                case GlobalTokens.STR_T:
                    return string.Empty;

                default:
                    var hasAttributes = (currentByte & 0x80) > 0;
                    var hasContent = (currentByte & 0x40) > 0;

                    var token = (byte)(currentByte & 0x3F);

                    if (hasAttributes)
                    {
                        return string.Empty;
                    }

                    var strTag = codePages[currentCodePage].GetTag(token) ?? string.Format(CultureInfo.CurrentCulture, "UNKNOWN_TAG_{0,2:X}", token);

                    XmlNode newNode;
                    try
                    {
                        newNode = xmlDoc.CreateElement(codePages[currentCodePage].Xmlns, strTag, codePages[currentCodePage].Namespace);
                    }
                    catch (XmlException)
                    {
                        return string.Empty;
                    }

                    try
                    {
                        var codepageName = codePages[currentCodePage].Xmlns;
                        var combinedTagAndToken = string.Format(CultureInfo.CurrentCulture, @"{0}|{1}|{2}|{3}", decodeDataCollection.Count, codepageName, strTag, token);
                        decodeDataCollection.Add(combinedTagAndToken, currentCodePage);
                    }
                    catch (ArgumentException)
                    {
                        return string.Empty;
                    }

                    newNode.Prefix = string.Empty;
                    currentNode.AppendChild(newNode);

                    if (hasContent)
                    {
                        currentNode = newNode;
                    }

                    break;
            }
        }

        using (var stringWriter = new StringWriter())
        {
            var textWriter = new XmlTextWriter(stringWriter) { Formatting = Formatting.Indented };
            xmlDoc.WriteTo(textWriter);
            textWriter.Flush();

            return stringWriter.ToString();
        }
    }

    /// <summary>
    /// Loads xml string and encodes to bytes.
    /// </summary>
    /// <param name="xmlValue">The xml string.</param>
    /// <returns>The encoded bytes.</returns>
    public byte[] EncodeToWBXML(string xmlValue)
    {
        xmlDoc.LoadXml(xmlValue);

        var byteList = new List<byte>();

        // Initializes the EncodeDataCollection
        if (encodeDataCollection == null)
        {
            encodeDataCollection = new Dictionary<string, int>();
        }
        else
        {
            encodeDataCollection.Clear();
        }

        byteList.Add(VersionByte);
        byteList.Add(PublicIdentifierByte);
        byteList.Add(CharsetByte);
        byteList.Add(StringTableLengthByte);

        foreach (XmlNode node in xmlDoc.ChildNodes)
        {
            byteList.AddRange(EncodeNode(node));
        }

        return byteList.ToArray();
    }

    /// <summary>
    /// Encodes a string.
    /// </summary>
    /// <param name="value">The string to encode.</param>
    /// <returns>The encoded bytes.</returns>
    private static byte[] EncodeString(string value)
    {
        var byteList = new List<byte>();

        var charArray = value.ToCharArray();

        for (var i = 0; i < charArray.Length; i++)
        {
            byteList.Add((byte)charArray[i]);
        }

        byteList.Add(0x00);

        return byteList.ToArray();
    }

    /// <summary>
    /// Encodes multi byte integer
    /// </summary>
    /// <param name="value">Then integer to encode.</param>
    /// <returns>The encoded bytes</returns>
    private static byte[] EncodeMultibyteInteger(int value)
    {
        var byteList = new List<byte>();

        while (value > 0)
        {
            var addByte = (byte)(value & 0x7F);

            if (byteList.Count > 0)
            {
                addByte |= 0x80;
            }

            byteList.Insert(0, addByte);

            value >>= 7;
        }

        return byteList.ToArray();
    }

    /// <summary>
    /// Encodes opaque data.
    /// </summary>
    /// <param name="opaqueBytes">The opaque data</param>
    /// <returns>The encoded bytes.</returns>
    private static byte[] EncodeOpaque(byte[] opaqueBytes)
    {
        var byteList = new List<byte>();

        byteList.AddRange(EncodeMultibyteInteger(opaqueBytes.Length));
        byteList.AddRange(opaqueBytes);

        return byteList.ToArray();
    }

    /// <summary>
    /// Encodes a node.
    /// </summary>
    /// <param name="node">The node need to encode.</param>
    /// <returns>The encoded bytes.</returns>
    private byte[] EncodeNode(XmlNode node)
    {
        var byteList = new List<byte>();
        switch (node.NodeType)
        {
            case XmlNodeType.Element:
                if (node.Attributes != null && node.Attributes.Count > 0)
                {
                    ParseXmlnsAttributes(node);
                }

                if (SetCodePageByXmlns(node.NamespaceURI))
                {
                    byteList.Add((byte)GlobalTokens.SWITCH_PAGE);
                    byteList.Add((byte)currentCodePage);
                }

                // Gets token in this.codePages
                var wbxmlMapToken = codePages[currentCodePage].GetToken(node.LocalName);
                var token = wbxmlMapToken;
                if (node.HasChildNodes)
                {
                    token |= 0x40;
                }

                byteList.Add(token);

                var codepageName = codePages[currentCodePage].Xmlns;
                var combinedTagAndToken = string.Format(CultureInfo.CurrentCulture, @"{0}|{1}|{2}|{3}", encodeDataCollection.Count, codepageName, node.LocalName, wbxmlMapToken);
                encodeDataCollection.Add(combinedTagAndToken, currentCodePage);

                if (node.HasChildNodes)
                {
                    foreach (XmlNode child in node.ChildNodes)
                    {
                        byteList.AddRange(EncodeNode(child));
                    }

                    byteList.Add((byte)GlobalTokens.END);
                }

                break;
            case XmlNodeType.Text:
                byteList.Add((byte)GlobalTokens.STR_I);
                byteList.AddRange(EncodeString(node.Value));
                break;
            case XmlNodeType.CDATA:
                byteList.Add((byte)GlobalTokens.OPAQUE);

                var cdataValue = System.Text.Encoding.ASCII.GetBytes(node.Value);
                if (node.ParentNode.Name == "ConversationId"
                    || node.ParentNode.Name == "ConversationIndex")
                {
                    cdataValue = Convert.FromBase64String(node.Value);
                }

                byteList.AddRange(EncodeOpaque(cdataValue));
                break;
        }

        return byteList.ToArray();
    }

    /// <summary>
    /// Gets code page index by namespace.
    /// </summary>
    /// <param name="nameSpace">The namespace</param>
    /// <returns>The index of code page</returns>
    private int GetCodePageByNamespace(string nameSpace)
    {
        for (var i = 0; i < codePages.Length; i++)
        {
            if (string.Equals(codePages[i].Namespace, nameSpace, StringComparison.CurrentCultureIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Switches to the code page by prefix.
    /// </summary>
    /// <param name="namespaceUri">The prefix</param>
    /// <returns>True, if successful.</returns>
    private bool SetCodePageByXmlns(string namespaceUri)
    {
        if (string.IsNullOrEmpty(namespaceUri))
        {
            if (currentCodePage != defaultCodePage)
            {
                currentCodePage = defaultCodePage;
                return true;
            }

            return false;
        }

        if (string.Equals(codePages[currentCodePage].Xmlns, namespaceUri, StringComparison.CurrentCultureIgnoreCase))
        {
            return false;
        }

        for (var i = 0; i < codePages.Length; i++)
        {
            if (string.Equals(codePages[i].Namespace, namespaceUri, StringComparison.CurrentCultureIgnoreCase))
            {
                currentCodePage = i;
                return true;
            }
        }

        throw new InvalidDataException($"Unknown Xmlns: {namespaceUri}.");
    }

    /// <summary>
    /// Parses namespaceUri attribute
    /// </summary>
    /// <param name="node">The xml node to parse</param>
    private void ParseXmlnsAttributes(XmlNode node)
    {
        if (node.Attributes == null)
        {
            return;
        }

        foreach (XmlAttribute attribute in node.Attributes)
        {
            var codePage = GetCodePageByNamespace(attribute.Value);

            if (!string.IsNullOrEmpty(attribute.Value) && (attribute.Value.StartsWith("http://www.w3.org/2001/XMLSchema-instance", StringComparison.CurrentCultureIgnoreCase) || attribute.Value.StartsWith("http://www.w3.org/2001/XMLSchema", StringComparison.CurrentCultureIgnoreCase)))
            {
                break;
            }

            if (string.Equals(attribute.Name, "XMLNS", StringComparison.CurrentCultureIgnoreCase))
            {
                defaultCodePage = codePage;
            }
            else if (string.Equals(attribute.Prefix, "XMLNS", StringComparison.CurrentCultureIgnoreCase))
            {
                codePages[codePage].Xmlns = attribute.LocalName;
            }
        }
    }
}
