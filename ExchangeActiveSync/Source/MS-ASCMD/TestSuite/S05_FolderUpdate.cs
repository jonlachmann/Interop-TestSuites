namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using Common;
using Common.Response;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// This scenario is designed to test the FolderUpdate command.
/// </summary>
[TestClass]
public class S05_FolderUpdate : TestSuiteBase
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
    /// This test case is used to verify if the FolderUpdate command request was successful, the status should be equal to 1.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC01_FolderUpdate_Success()
    {
        #region Call method FolderCreate command to create a new folder as a child folder of the specified parent folder.
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Call method FolderUpdate to rename a folder.
        var updateFolderName = Common.GenerateResourceName(Site, "FolderUpdate");
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(folderCreateResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId, updateFolderName, "0");
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);
        #endregion

        #region Call method FolderSync to synchronize the collection hierarchy.
        var folderSyncResponse = FolderSync();
        var isFolderRenamed = false;
        foreach (var add in folderSyncResponse.ResponseData.Changes.Add)
        {
            if ((add.DisplayName == updateFolderName) && (add.ServerId == folderCreateResponse.ResponseData.ServerId))
            {
                isFolderRenamed = true;
                break;
            }
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R140");

        // Verify MS-ASCMD requirement: MS-ASCMD_R140
        Site.CaptureRequirementIfIsTrue(
            isFolderRenamed,
            140,
            @"[In FolderUpdate] The [FolderUpdate] command is also used to rename a folder.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4100");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4100
        Site.CaptureRequirementIfAreEqual<string>(
            "1",
            folderUpdateResponse.ResponseData.Status,
            4100,
            @"[In Status(FolderUpdate)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_ R5784");

        // IF folder has been updated successfully, the server must send a synchronization key to the client in a response.
        Site.CaptureRequirementIfIsNotNull(
            folderUpdateResponse.ResponseData.SyncKey,
            5784,
            @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] After a successful [FolderCreate command (section 2.2.2.2), FolderDelete command (section 2.2.2.3), or] FolderUpdate command (section 2.2.2.5), the server MUST send a synchronization key to the client in a response.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if specified folder is a special folder, the status in return value is 2.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC02_FolderUpdate_Status2()
    {
        // Call method FolderUpdate to rename the Calendar folder.
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(LastFolderSyncKey, User1Information.CalendarCollectionId, "Notes", "0");
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4101");

        // When the special folder, such as the Inbox, Outbox, Contacts, or Drafts folders, be updated, server will return status 2.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4101
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4101,
            @"[In Status(FolderUpdate)] [When the scope is] Item, [the meaning of the status value] 2 [is] A folder with that name already exists or the specified folder is a special folder.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4102");

        // When the special folder, such as the Inbox, Outbox, Contacts, or Drafts folders, be updated, server will return status 2.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4102
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4102,
            @"[In Status(FolderUpdate)] [When the scope is Item], [the cause of the status value 2 is] [A folder with that name already exists or] the specified folder is a special folder, such as the Inbox, Outbox, Contacts, or Drafts folders. Special folders cannot be updated.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4096");

        // When the special folder, such as the Inbox, Outbox, Contacts, or Drafts folders, be updated, server will return status 2.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4096
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4096,
            @"[In Status(FolderUpdate)] If the command fails, the Status element contains a code that indicates the type of failure.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_ R5786");

        // If the FolderUpdate command is not successful, the server must not return a SyncKey element.
        Site.CaptureRequirementIfIsNull(
            folderUpdateResponse.ResponseData.SyncKey,
            5786,
            @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] If the [FolderCreate command, FolderDelete command, or] FolderUpdate command is not successful, the server MUST NOT return a SyncKey element.");
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if specified folder does not exist, the status in return value is 4.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC03_FolderUpdate_Status4()
    {
        // Call method FolderUpdate to rename a non existent folder.
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(LastFolderSyncKey, "InvalidServerId", Common.GenerateResourceName(Site, "FolderUpdate"), "0");
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4105");

        // If the specified folder is a non existent folder when call FolderUpdate command, server will return status 4.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4105
        Site.CaptureRequirementIfAreEqual<int>(
            4,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4105,
            @"[In Status(FolderUpdate)] [When the scope is] Item, [the meaning of the status value] 4 [is] The specified folder does not exist.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4106");

        // If the specified folder is a non existent folder when call FolderUpdate command, server will return status 4.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4106
        Site.CaptureRequirementIfAreEqual<int>(
            4,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4106,
            @"[In Status(FolderUpdate)] [When the scope is Item], [the cause of the status value 4 is] Client specified a nonexistent folder in a FolderUpdate command request.");
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if specified parent folder does not exist, the status in return value is 5.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC04_FolderUpdate_Status5()
    {
        // Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);

        // Call method FolderUpdate to move the folder to a nonexistent parent folder.
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(folderCreateResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId, Common.GenerateResourceName(Site, "FolderUpdate"), "InvalidParentId");
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4108");

        // If calls method FolderUpdate to move the folder to a nonexistent parent folder, server will return status 5.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4108
        Site.CaptureRequirementIfAreEqual<int>(
            5,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4108,
            @"[In Status(FolderUpdate)] [When the scope is] Item, [the meaning of the status value] 5 [is] The specified parent folder was not found.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4109");

        // If calls method FolderUpdate to move the folder to a nonexistent parent folder, server will return status 5.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4109
        Site.CaptureRequirementIfAreEqual<int>(
            5,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4109,
            @"[In Status(FolderUpdate)] [When the scope is Item], [the cause of the status value 5 is] Client specified a nonexistent folder in a FolderUpdate command request.");
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if the request contains invalid synchronization key, the status in return value is 9.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC05_FolderUpdate_Status9()
    {
        #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Call method FolderUpdate to rename a folder with invalid synchronization key.
        var folderUpdateRequest = Common.CreateFolderUpdateRequest("InvalidSyncKey", folderCreateResponse.ResponseData.ServerId, Common.GenerateResourceName(Site, "FolderUpdate"), "0");
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4120");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4120
        Site.CaptureRequirementIfAreEqual<int>(
            9,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4120,
            @"[In Status(FolderUpdate)] [When the scope is Global], [the cause of the status value 9 is] The client sent a malformed or mismatched synchronization key, or the synchronization state is corrupted on the server.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if the request contains a semantic error, the status in return value is 10.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC06_FolderUpdate_Status10()
    {
        #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Call method FolderUpdate to rename the created folder without SyncKey element.
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(null, folderCreateResponse.ResponseData.ServerId, Common.GenerateResourceName(Site, "FolderUpdate"), "0");
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3102");

        // The server should return a status code 10 in the FolderUpdate command response to indicate the client sent FolderUpdate request contains a semantic error.
        // Verify MS-ASCMD requirement: MS-ASCMD_R3102
        Site.CaptureRequirementIfAreEqual<int>(
            10,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            3102,
            @"[In FolderUpdate] Including the Status element in a FolderUpdate request results in a Status element value of 10 being returned in the response.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4123");

        // The server should return a status code 10 in the FolderUpdate command response to indicate the client sent FolderUpdate request contains a semantic error.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4123
        Site.CaptureRequirementIfAreEqual<int>(
            10,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4123,
            @"[In Status(FolderUpdate)] [When the scope is Global], [the cause of the status value 10 is] The client sent a FolderUpdate command request that contains a semantic error.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4568");

        // The server should return a status code 10 in the FolderUpdate command response to indicate the client sent FolderUpdate request does not contain SyncKey element.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4568
        Site.CaptureRequirementIfAreEqual<int>(
            10,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4568,
            @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] The server returns a Status element (section 2.2.3.162.5) value of 10 if the SyncKey element is not included in a FolderUpdate command request.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if moves the folder failed, the status in return value is 2.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC07_FolderUpdate_Moves()
    {
        #region Call method FolderCreate to create a new folder as a child folder of the mailbox Root folder.
        var folderName = Common.GenerateResourceName(Site, "FolderCreate");
        var folderCreateRequest = Common.CreateFolderCreateRequest(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
        var folderCreateResponse = CMDAdapter.FolderCreate(folderCreateRequest);
        var folderServerId1 = folderCreateResponse.ResponseData.ServerId;
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Call method FolderUpdate to move the new created folder from the mailbox Root folder to SentItems folder on the server.
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(folderCreateResponse.ResponseData.SyncKey, folderServerId1, folderName, ((byte)FolderType.SentItems).ToString());
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);
        Site.Assert.AreEqual<int>(1, int.Parse(folderUpdateResponse.ResponseData.Status), "Server should return status 1 to indicate FolderUpdate command success.");
        #endregion

        #region Call method FolderSync to synchronize the collection hierarchy.
        var folderSyncResponse = FolderSync();
        var isFolderMoved = false;
        foreach (var add in folderSyncResponse.ResponseData.Changes.Add)
        {
            if ((add.ServerId == folderServerId1) && (add.ParentId == ((byte)FolderType.SentItems).ToString()))
            {
                isFolderMoved = true;
                break;
            }
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R139");

        // Verify MS-ASCMD requirement: MS-ASCMD_R139
        Site.CaptureRequirementIfIsTrue(
            isFolderMoved,
            139,
            @"[In FolderUpdate] The FolderUpdate command moves a folder from one location to another on the server.");

        // Call method FolderCreate to create another new folder which its name is same with above step as a child folder of the mailbox Root folder.
        folderCreateRequest = Common.CreateFolderCreateRequest(folderSyncResponse.ResponseData.SyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
        folderCreateResponse = CMDAdapter.FolderCreate(folderCreateRequest);
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);

        // Call method FolderUpdate to move the newest created folder in mailbox Root folder from mailbox Root folder to SentItems folder on the server.
        folderUpdateRequest = Common.CreateFolderUpdateRequest(folderCreateResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId, folderName, ((byte)FolderType.SentItems).ToString());
        folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5438");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5438
        Site.CaptureRequirementIfAreEqual<int>(
            2,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            5438,
            @"[In Status(FolderUpdate)] [When the scope is Item], [the cause of the status value 2 is] A folder with that name already exists [or the specified folder is a special folder, such as the Inbox, Outbox, Contacts, or Drafts folders. Special folders cannot be updated].");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderUpdate command, if specified folder is the recipient information cache, the status in return value is 3.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S05_TC08_FolderUpdate_Status3()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The recipient information cache is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        // Call method FolderUpdate attempting to update a recipient information cache
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(LastFolderSyncKey, "RI", Common.GenerateResourceName(Site, "FolderUpdate"), User1Information.RecipientInformationCacheCollectionId);
        var folderUpdateResponse = CMDAdapter.FolderUpdate(folderUpdateRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R143");

        // Verify MS-ASCMD requirement: MS-ASCMD_R143
        Site.CaptureRequirementIfAreEqual<int>(
            3,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            143,
            @"[In FolderUpdate] Attempting to update a recipient information cache using this [FolderUpdate] command results in a Status element (section 2.2.3.162.5) value of 3.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R4103");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4103
        Site.CaptureRequirementIfAreEqual<int>(
            3,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4103,
            @"[In Status(FolderUpdate)] [When the scope is] Item, [the meaning of the status value] 3 [is] The specified folder is the Recipient information folder, which cannot be updated by the client.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R4104");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4104
        Site.CaptureRequirementIfAreEqual<int>(
            3,
            int.Parse(folderUpdateResponse.ResponseData.Status),
            4104,
            @"[In Status(FolderUpdate)] [When the scope is Item], [the cause of the status value 3 is] The client specified the Recipient information folder, which is a special folder. Special folders cannot be updated.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R141");

        // Since server return status 3, the FolderUpdate command cannot be used to update a recipient information cache.
        // Verify MS-ASCMD requirement: MS-ASCMD_R141
        Site.CaptureRequirement(
            141,
            @"[In FolderUpdate] The FolderUpdate command cannot be used to update a recipient information cache.");
    }
    #endregion
}