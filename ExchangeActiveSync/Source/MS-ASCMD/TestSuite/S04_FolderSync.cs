namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using Common;
using Common.Response;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Response = Common.Response;

/// <summary>
/// This scenario is designed to test the FolderSync command.
/// </summary>
[TestClass]
public class S04_FolderSync : TestSuiteBase
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
    /// This test case is used to verify FolderSync command, all folders MUST be returned to the client when initial folder synchronization is done with a synchronization key of 0(zero).
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC01_FolderSync_SyncKey0()
    {
        // Call method FolderSync to synchronize the collection hierarchy.
        var folderSyncResponse = FolderSync();
        Site.Assert.AreEqual<int>(
            1,
            int.Parse(folderSyncResponse.ResponseData.Status),
            "The server should return a status code 1 in the FolderSync command response to indicate success.");

        Site.Assert.IsNotNull(
            folderSyncResponse.ResponseData.SyncKey,
            "The server should return a non-null SyncKey in the FolderSync command response to indicate success.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R113");

        // The Count element is great than 0 in the FolderSync command response to indicate folders are returned to the client.
        // Verify MS-ASCMD requirement: MS-ASCMD_R113
        Site.CaptureRequirementIfIsTrue(
            folderSyncResponse.ResponseData.Changes.Count > 0,
            113,
            @"[In FolderSync] The FolderSync command synchronizes the collection hierarchy.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R119");

        // The Count element is great than 0 in the FolderSync command response to indicate folders are returned to the client.
        // Verify MS-ASCMD requirement: MS-ASCMD_R119
        Site.CaptureRequirementIfIsTrue(
            folderSyncResponse.ResponseData.Changes.Count > 0,
            119,
            @"[In FolderSync] All folders MUST be returned to the client when initial folder synchronization is done with a synchronization key of 0 (zero).");

        var isVerifyR5416 = false;
        foreach (var add in folderSyncResponse.ResponseData.Changes.Add)
        {
            var name = add.DisplayName;
            var serverId = add.ServerId;
            foreach (var addNew in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (serverId != addNew.ServerId)
                {
                    if (name != addNew.DisplayName)
                    {
                        isVerifyR5416 = true;
                    }
                    else
                    {
                        isVerifyR5416 = false;
                        break;
                    }
                }
            }
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5416");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5416
        Site.CaptureRequirementIfIsTrue(
            isVerifyR5416,
            5416,
            @"[In DisplayName(FolderSync)] Subfolder display names MUST be unique for a sample of N (default N=10) within a folder.");

        // Folder has been synced successfully, server returns a non-null SyncKey.
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R116");

        // Verify MS-ASCMD requirement: MS-ASCMD_R116
        Site.CaptureRequirementIfIsNotNull(
            folderSyncResponse.ResponseData.SyncKey,
            116,
            @"[In FolderSync] The synchronization key is returned in the SyncKey element of the response if the FolderSync command succeeds.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4074");

        // Verify MS-ASCMD requirement: MS-ASCMD_R4074
        Site.CaptureRequirementIfAreEqual<string>(
            "1",
            folderSyncResponse.ResponseData.Status,
            4074,
            @"[In Status(FolderSync)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4578");

        // Folder has been synced successfully, server returns a non-null SyncKey.
        Site.CaptureRequirementIfIsNotNull(
            folderSyncResponse.ResponseData.SyncKey,
            4578,
            @"[In SyncKey(FolderSync)] After successful folder synchronization, the server MUST send a synchronization key to the client.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5008");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5008
        // Folder has been synced successfully, server returns a non-null SyncKey.
        Site.CaptureRequirementIfIsNotNull(
            folderSyncResponse.ResponseData.SyncKey,
            5008,
            @"[In Synchronizing a Folder Hierarchy] The server responds with a new folderhierarchy:SyncKey element value and provides a list of all the folders in the user's mailbox.");
    }

    /// <summary>
    /// This test case is used to verify FolderSync command, if there are no changes since the last folders synchronization, a Count element value of 0 (zero) is returned.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC02_FolderSync_NoChanged()
    {
        // The client calls FolderSync command to synchronize the collection hierarchy if no changes occurred for folder.
        var folderSyncRequest = Common.CreateFolderSyncRequest(LastFolderSyncKey);
        var folderSyncResponse = CMDAdapter.FolderSync(folderSyncRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2120");

        // Verify MS-ASCMD requirement: MS-ASCMD_R2120
        Site.CaptureRequirementIfAreEqual<uint>(
            0,
            folderSyncResponse.ResponseData.Changes.Count,
            2120,
            @"[In Count] If there are no changes since the last folder synchronization, a Count element value of 0 (zero) is returned.");
    }

    /// <summary>
    /// This test case is used to verify FolderSync command, if any changes have occurred on the server, the count is not equal to 0.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC03_FolderSync_Changed()
    {
        #region Change a new DeviceID and call FolderSync command.
        CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", Site));
        var folderSyncResponseForDefaultDeviceID = FolderSync(); 
        #endregion

        #region Change a new DeviceID and call FolderSync command.
        CMDAdapter.ChangeDeviceID("NewDeviceID");
        RecordDeviceInfoChanged();
        var folderName = Common.GenerateResourceName(Site, "FolderSync");
        var folderSyncResponseForNewDeviceID = FolderSync();
        #endregion

        #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
        var folderCreateResponse = GetFolderCreateResponse(folderSyncResponseForNewDeviceID.ResponseData.SyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Change the DeviceId back and call method FolderSync to synchronize the collection hierarchy.
        CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", Site));
        var folderSyncRequest = Common.CreateFolderSyncRequest(folderSyncResponseForNewDeviceID.ResponseData.SyncKey);
        var folderSyncResponse = CMDAdapter.FolderSync(folderSyncRequest);
        foreach (var add in folderSyncResponse.ResponseData.Changes.Add)
        {
            if (add.DisplayName == folderName)
            {
                User1Information.UserCreatedFolders.Clear();
                RecordCaseRelativeFolders(User1Information, add.ServerId);
                break;
            }
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5024");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5024
        Site.CaptureRequirementIfAreEqual<uint>(
            (uint)1,
            folderSyncResponse.ResponseData.Changes.Count,
            5024,
            @"[In Synchronizing a Folder Hierarchy] [FolderSync sequence for folder hierarchy synchronization, order 2:] If any changes have occurred on the server, the new, deleted, or changed folders are returned to the client.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderSync command, if client resynchronizes the existing folder hierarchy, ServerId values do not change. 
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC04_FolderSync_Resynchronizes()
    {
        // The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
        var folderName = Common.GenerateResourceName(Site, "FolderSync");
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);

        // The client calls FolderSync method to synchronize the collection hierarchy, then server returns latest folder SyncKey.
        var folderSyncResponse = FolderSync();

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5012");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5012
        Site.CaptureRequirementIfAreEqual(
            folderCreateResponse.ResponseData.ServerId, 
            GetCollectionId(folderSyncResponse, folderName),
            5012, 
            @"[In Synchronizing a Folder Hierarchy] Existing folderhierarchy:ServerId values do not change when the client resynchronizes.");
    }

    /// <summary>
    /// This test case is used to verify FolderSync command, if the SyncKey is an empty string, the status is equal to 9.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC05_FolderSync_Status9()
    {
        // Call method FolderSync with an empty SyncKey to synchronize the collection hierarchy.
        var folderSyncRequest = new FolderSyncRequest { RequestData = { SyncKey = string.Empty } };
        var folderSyncResponse = CMDAdapter.FolderSync(folderSyncRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4068");

        // If the client sent a malformed or mismatched synchronization key, the server should return a status code 9 in the FolderSync command response.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4068
        Site.CaptureRequirementIfAreEqual<int>(
            9,
            int.Parse(folderSyncResponse.ResponseData.Status),
            4068,
            @"[In Status(FolderSync)] If the command fails, the Status element contains a code that indicates the type of failure.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4070");

        // If the client sent a malformed or mismatched synchronization key, the server should return a status code 9 in the FolderSync command response.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4070
        Site.CaptureRequirementIfAreEqual<int>(
            9,
            int.Parse(folderSyncResponse.ResponseData.Status),
            4070,
            @"[In Status(FolderSync)] If one collection fails, a failure status MUST be returned for all collections.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4083");

        // The server should return a status code 9 in the FolderSync command response to indicate the client sent a malformed or mismatched synchronization key.
        // If the SyncKey is an empty string, the status is equal to 9.
        Site.CaptureRequirementIfAreEqual<int>(
            9,
            int.Parse(folderSyncResponse.ResponseData.Status),
            4083,
            @"[In Status(FolderSync)] [When the scope is Global], [the cause of the status value 9 is] The client sent a malformed or mismatched synchronization key [, or the synchronization state is corrupted on the server].");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4581");

        // The server should return a status code 9 in the FolderSync command response to indicate the client sent a malformed or mismatched synchronization key.
        // Verify MS-ASCMD requirement: MS-ASCMD_R4581
        Site.CaptureRequirementIfAreEqual<int>(
            9,
            int.Parse(folderSyncResponse.ResponseData.Status),
            4581,
            @"[In SyncKey(FolderSync)] The server MUST return a Status element (section 2.2.3.162.4) value of 9 if the value of the SyncKey element does not match the value of the synchronization key on the server.");
    }

    /// <summary>
    /// This test case is used to verify FolderSync command, if the SyncKey is invalid, the status is equal to 10.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC06_FolderSync_Status10()
    {
        // Call method FolderSync to synchronize the collection hierarchy with a null SyncKey.
        var folderSyncRequest = new FolderSyncRequest { RequestData = { SyncKey = null } };
        var folderSyncResponse = CMDAdapter.FolderSync(folderSyncRequest);

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4086");

        // If the SyncKey is invalid, the status is equal to 10.
        Site.CaptureRequirementIfAreEqual<int>(
            10,
            int.Parse(folderSyncResponse.ResponseData.Status),
            4086,
            @"[In Status(FolderSync)] [When the scope is Global], [the cause of the status value 10 is] The client sent a FolderSync command request that contains a semantic or syntactic error.");
    }

    /// <summary>
    /// This test case is used to verify FolderSync command synchronizes the folder hierarchy successfully after adding a folder.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC07_FolderSync_AddFolder()
    {
        #region The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderSync"), "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Call method FolderSync to synchronize the collection hierarchy.
        var folderSyncResponse = FolderSync();
        var isVerifyR5860 = false;
        foreach (var add in folderSyncResponse.ResponseData.Changes.Add)
        {
            if (add.ServerId == folderCreateResponse.ResponseData.ServerId)
            {
                isVerifyR5860 = true;
                break;
            }
        }

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5860");

        // Verify MS-ASCMD requirement: MS-ASCMD_R5860
        Site.CaptureRequirementIfIsTrue(
            isVerifyR5860,
            5860,
            @"[In Add(FolderSync)] [The Add element] creates a new folder on the client.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderSync command synchronizes the updated folder successfully.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC08_FolderSync_UpdateFolder()
    {
        #region Call method FolderCreate command to create a new folder as a child folder of the specified parent folder.
        var folderName = Common.GenerateResourceName(Site, "FolderSync");
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        RecordCaseRelativeFolders(User1Information, folderCreateResponse.ResponseData.ServerId);
        #endregion

        #region Change DeviceID
        CMDAdapter.ChangeDeviceID("NewDeviceID");
        RecordDeviceInfoChanged();
        var folderSyncKey = folderCreateResponse.ResponseData.SyncKey;
        #endregion

        #region Call method FolderSync to synchronize the collection hierarchy.
        var foldersyncResponseForNewDeviceID = FolderSync();
        var changeDeviceIDFolderId = GetCollectionId(foldersyncResponseForNewDeviceID, folderName);
        Site.Assert.IsFalse(string.IsNullOrEmpty(changeDeviceIDFolderId), "If the new folder created by FolderCreate command, server should return a server ID for the new created folder.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5020");

        // If client sends the FolderSync request successfully, the server must send a synchronization key to the client in a response.
        Site.CaptureRequirementIfIsTrue(
            (foldersyncResponseForNewDeviceID.ResponseData.SyncKey != null) && (foldersyncResponseForNewDeviceID.ResponseData.SyncKey != folderSyncKey),
            5020,
            @"[In Synchronizing a Folder Hierarchy] [FolderSync sequence for folder hierarchy synchronization, order 1:] The server responds with [the folder hierarchy and] a new folderhierarchy:SyncKey value.");

        #endregion

        #region Call method FolderUpdate to rename a folder.
        var folderUpdateName = Common.GenerateResourceName(Site, "FolderUpdate");
        var folderUpdateRequest = Common.CreateFolderUpdateRequest(foldersyncResponseForNewDeviceID.ResponseData.SyncKey, changeDeviceIDFolderId, folderUpdateName, "0");
        CMDAdapter.FolderUpdate(folderUpdateRequest);
        #endregion

        #region Restore DeviceID and call FolderSync command.
        CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", Site));

        // The client calls FolderSync command to synchronize the collection hierarchy with original device id.
        var folderSyncRequest = Common.CreateFolderSyncRequest(folderSyncKey);
        var folderSyncResponse = CMDAdapter.FolderSync(folderSyncRequest);
        var isUpdated = false;
        foreach (var update in folderSyncResponse.ResponseData.Changes.Update)
        {
            if (update.DisplayName == folderUpdateName)
            {
                isUpdated = true;
                break;
            }
        }

        Site.Assert.IsTrue(isUpdated, "Rename successfully");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderSync command synchronizes the folder hierarchy successfully after deleting a folder.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC09_FolderSync_DeleteFolder()
    {
        #region The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
        var folderName = Common.GenerateResourceName(Site, "FolderSync");
        var folderCreateResponse = GetFolderCreateResponse(LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
        Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
        #endregion

        #region Changes DeviceID.
        CMDAdapter.ChangeDeviceID("NewDeviceID");
        RecordDeviceInfoChanged();
        #endregion

        #region Calls method FolderSync to synchronize the collection hierarchy.
        var folderSyncRequestForNewDeviceID = Common.CreateFolderSyncRequest("0");
        var folderSyncResponseForNewDeviceID = CMDAdapter.FolderSync(folderSyncRequestForNewDeviceID);

        // Gets the server ID for new folder after change DeviceID.
        var serverId = GetCollectionId(folderSyncResponseForNewDeviceID, folderName);
        Site.Assert.IsNotNull(serverId, "Call method GetServerId to get a non-null ServerId to indicate success.");
        #endregion

        #region The client calls FolderDelete command to delete the created folder in step 2 from the server.
        var folderDeleteRequest = Common.CreateFolderDeleteRequest(folderSyncResponseForNewDeviceID.ResponseData.SyncKey, serverId);
        var folderDeleteResponse = CMDAdapter.FolderDelete(folderDeleteRequest);
        Site.Assert.AreEqual<int>(1, int.Parse(folderDeleteResponse.ResponseData.Status), "The server should return a status code 1 in the FolderDelete command response to indicate success.");
        #endregion

        #region Restore DeviceID and call FolderSync command
        CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", Site));

        // The client calls FolderSync command to synchronize the collection hierarchy with original device id.
        var folderSyncRequest = Common.CreateFolderSyncRequest(folderCreateResponse.ResponseData.SyncKey);
        var folderSyncResponse = CMDAdapter.FolderSync(folderSyncRequest);
        Site.Assert.AreEqual<int>(1, int.Parse(folderSyncResponse.ResponseData.Status), "Server should return status 1 in the FolderSync response to indicate success.");
        Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes, "Server should return Changes element in the FolderSync response after the collection hierarchy changed by call FolderDelete command.");
        Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes.Delete, "Server should return Changes element in the FolderSync response after the specified folder deleted.");
            
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5863");

        // The folderDeleteSuccess is true indicates the folder which deleted by FolderDelete command is deleted successfully.
        // Verify MS-ASCMD requirement: MS-ASCMD_R5863
        Site.CaptureRequirementIfIsNotNull(
            folderSyncResponse.ResponseData.Changes.Delete[0].ServerId,
            5863,
            @"[In Delete(FolderSync)] [The Delete element] specifies that a folder on the server was deleted since the last folder synchronization.");
        #endregion
    }

    /// <summary>
    /// This test case is used to verify FolderSync command synchronizes the folder hierarchy successfully, but does not synchronize the items in the collections themselves.
    /// </summary>
    [TestCategory("MSASCMD"), TestMethod]
    public void MSASCMD_S04_TC10_FolderSync_NoSynchronizeItems()
    {
        #region User2 calls method SendMail to send MIME-formatted e-mail messages to user1.
        SwitchUser(User2Information);
        var subject = Common.GenerateResourceName(Site, "subject");
        var responseSendMail = SendPlainTextEmail(null, subject, User2Information.UserName, User1Information.UserName, null);
        Site.Assert.AreEqual(string.Empty, responseSendMail.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
        #endregion

        #region Switch to user1 mailbox and call FolderSync command
        SwitchUser(User1Information);
        GetMailItem(User1Information.InboxCollectionId, subject);
        RecordCaseRelativeItems(User1Information, User1Information.InboxCollectionId, subject);

        // Call method FolderSync to synchronize the collection hierarchy.
        var folderSyncResponse = FolderSync();

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5775");
        var isVerifyR5775 = true;

        // FolderSync command does not synchronize the new email items, so FolderChangesAdd does not contain new email items.
        foreach (var add in folderSyncResponse.ResponseData.Changes.Add)
        {
            if (add.DisplayName == subject && add.ParentId == User1Information.InboxCollectionId)
            {
                isVerifyR5775 = false;
                break;
            }
        }

        // Verify MS-ASCMD requirement: MS-ASCMD_R5775
        Site.CaptureRequirementIfIsTrue(
            isVerifyR5775,
            5775,
            @"[In FolderSync] But [FolderSync command] does not synchronize the items in the collections themselves.");
        #endregion
    }
    #endregion
}