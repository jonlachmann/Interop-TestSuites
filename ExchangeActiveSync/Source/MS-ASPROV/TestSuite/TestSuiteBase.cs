namespace Microsoft.Protocols.TestSuites.MS_ASPROV;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using Common;
using Common.DataStructures;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Request = Common.Request;

/// <summary>
/// The base class of scenario class.
/// </summary>
[TestClass]
public class TestSuiteBase : TestClassBase
{
    #region Variables

    /// <summary>
    /// Gets or sets the DeviceType of ActiveSync client.
    /// </summary>
    public string DeviceType { get; protected set; }

    /// <summary>
    /// Gets or sets the related information of User1.
    /// </summary>
    protected UserInformation User1Information { get; set; }

    /// <summary>
    /// Gets or sets the related information of User2.
    /// </summary>
    protected UserInformation User2Information { get; set; }

    /// <summary>
    /// Gets or sets the related information of User3.
    /// </summary>
    protected UserInformation User3Information { get; set; }

    /// <summary>
    /// Gets or sets the related information of current user.
    /// </summary>
    protected UserInformation CurrentUserInformation { get; set; }

    /// <summary>
    /// Gets MS-ASPROV protocol adapter.
    /// </summary>
    protected IMS_ASPROVAdapter PROVAdapter { get; private set; }

    /// <summary>
    /// Gets MS-ASPROV SUT control adapter.
    /// </summary>
    protected IMS_ASPROVSUTControlAdapter PROVSUTControlAdapter { get; private set; }

    /// <summary>
    /// Gets the value of 'SutComputerName' specified in ptfconfig.
    /// </summary>
    protected string SutComputerName { get; private set; }

    #endregion

    #region Test case initialize and cleanup
    /// <summary>
    /// Initialize the Test suite.
    /// </summary>
    protected override void TestInitialize()
    {
        base.TestInitialize();
        PROVAdapter = Site.GetAdapter<IMS_ASPROVAdapter>();
        PROVSUTControlAdapter = Site.GetAdapter<IMS_ASPROVSUTControlAdapter>();

        // Set the information of User1.
        User1Information = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User1Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User1Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        // Set the information of User2.
        User2Information = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User2Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User2Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        // Set the information of User3.
        User3Information = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User3Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User3Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        // Initialize the UserInformation of CurrentUser
        CurrentUserInformation = new UserInformation();

        SutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", Site);

        // Switch the current user to the user with default policy.
        SwitchUser(User1Information, false);
    }

    /// <summary>
    /// Clean up the environment.
    /// </summary>
    protected override void TestCleanup()
    {
        if (!string.IsNullOrEmpty(DeviceType) && DeviceType != Common.GetConfigurationPropertyValue("DeviceType", Site) && !string.IsNullOrEmpty(CurrentUserInformation.UserName))
        {
            // Remove the device from the mobile list after wipe operation is successes.
            var deviceRemoved = false;
            var userEmail = Common.GetMailAddress(CurrentUserInformation.UserName, CurrentUserInformation.UserDomain);
            if (CurrentUserInformation.UserName == User1Information.UserName)
            {
                deviceRemoved = PROVSUTControlAdapter.RemoveDevice(SutComputerName, userEmail, User1Information.UserPassword, DeviceType);
            }
            else if (CurrentUserInformation.UserName == User2Information.UserName)
            {
                deviceRemoved = PROVSUTControlAdapter.RemoveDevice(SutComputerName, userEmail, User2Information.UserPassword, DeviceType);
            }
            else if (CurrentUserInformation.UserName == User3Information.UserName)
            {
                deviceRemoved = PROVSUTControlAdapter.RemoveDevice(SutComputerName, userEmail, User3Information.UserPassword, DeviceType);
            }

            Site.Assert.IsTrue(deviceRemoved, "The device with DeviceType {0} should be removed successfully.", DeviceType);

            // Restore the DeviceType.
            PROVAdapter.ApplyDeviceType(Common.GetConfigurationPropertyValue("DeviceType", Site));
        }

        // Clean the created items of User1.
        if (User1Information.UserCreatedItems.Count != 0)
        {
            // Switch the user credential to User1.
            SwitchUser(User1Information, true);
            DeleteCreatedItems(User1Information.UserCreatedItems);
        }

        // Clean the created items of User2.
        if (User2Information.UserCreatedItems.Count != 0)
        {
            // Switch the user credential to User2.
            SwitchUser(User2Information, true);
            DeleteCreatedItems(User2Information.UserCreatedItems);
        }

        // Clean the created items of User3.
        if (User3Information.UserCreatedItems.Count != 0)
        {
            // Switch the user credential to User3.
            SwitchUser(User3Information, true);
            DeleteCreatedItems(User3Information.UserCreatedItems);
        }

        // Reset the user credential.
        SwitchUser(User1Information, false);

        // Restore the PolicyKey.
        PROVAdapter.ApplyPolicyKey(string.Empty);

        base.TestCleanup();
    }
    #endregion

    #region Test case base methods
    /// <summary>
    /// Change user to call active sync operations and resynchronize the collection hierarchy.
    /// </summary>
    /// <param name="userInformation">The information of the user.</param>
    /// <param name="isFolderSyncNeeded">A boolean value that indicates whether needs to synchronize the folder hierarchy.</param>
    protected void SwitchUser(UserInformation userInformation, bool isFolderSyncNeeded)
    {
        PROVAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

        if (isFolderSyncNeeded)
        {
            AcknowledgeSecurityPolicySettings();

            // Call FolderSync command to synchronize the collection hierarchy.
            var folderSyncRequest = Common.CreateFolderSyncRequest("0");
            var folderSyncReponse = PROVAdapter.FolderSync(folderSyncRequest);

            // Verify FolderSync command response.
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderSyncReponse.ResponseData.Status),
                "If the FolderSync command executes successfully, the Status in response should be 1.");

            // Get the folder collectionId of User1
            if (userInformation.UserName == User1Information.UserName)
            {
                if (string.IsNullOrEmpty(User1Information.InboxCollectionId))
                {
                    User1Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncReponse, FolderType.Inbox, Site);
                }

                if (string.IsNullOrEmpty(User1Information.SentItemsCollectionId))
                {
                    User1Information.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncReponse, FolderType.SentItems, Site);
                }
            }

            // Get the folder collectionId of User3
            if (userInformation.UserName == User3Information.UserName)
            {
                if (string.IsNullOrEmpty(User3Information.InboxCollectionId))
                {
                    User3Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncReponse, FolderType.Inbox, Site);
                }
            }
        }
    }

    /// <summary>
    /// Acknowledge the current device policy setting.
    /// </summary>
    protected void AcknowledgeSecurityPolicySettings()
    {
        // Download the policy setting.
        var provisionResponse = CallProvisionCommand(string.Empty, "MS-EAS-Provisioning-WBXML", "1");
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        var temporaryPolicyKey = provisionResponse.ResponseData.Policies.Policy.PolicyKey;

        // Acknowledge the policy setting.
        provisionResponse = CallProvisionCommand(temporaryPolicyKey, "MS-EAS-Provisioning-WBXML", "1");
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        PROVAdapter.ApplyPolicyKey(provisionResponse.ResponseData.Policies.Policy.PolicyKey);
    }

    /// <summary>
    /// Call Provision command.
    /// </summary>
    /// <param name="policyKey">The policy key.</param>
    /// <param name="policyType">The format in which the policy settings are to be provided to the client device.</param>
    /// <param name="status">The status of the initial Provision command.</param>
    /// <returns>The response of Provision command.</returns>
    protected ProvisionResponse CallProvisionCommand(string policyKey, string policyType, string status)
    {
        // Create Provision command request.
        var provisionRequest = Common.CreateProvisionRequest(null, new Request.ProvisionPolicies(), null);
        var policy = new Request.ProvisionPoliciesPolicy { PolicyType = policyType };

        // The format in which the policy settings are to be provided to the client device.
        if (!string.IsNullOrEmpty(policyKey))
        {
            policy.PolicyKey = policyKey;
            policy.Status = status;
        }
        else if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "14.1" ||
                 Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "16.0" || 
                 Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) == "16.1")
        {
            // Configure the DeviceInformation.
            var deviceInfomation = new Request.DeviceInformation();
            var deviceInformationSet = new Request.DeviceInformationSet
            {
                Model = "ASPROVTest"
            };
            deviceInfomation.Set = deviceInformationSet;
            provisionRequest.RequestData.DeviceInformation = deviceInfomation;
        }

        provisionRequest.RequestData.Policies.Policy = policy;

        // Call Provision command.
        var provisionResponse = PROVAdapter.Provision(provisionRequest);

        return provisionResponse;
    }

    /// <summary>
    /// Get the initial syncKey of the specified folder.
    /// </summary>
    /// <param name="collectionId">The collection id of the specified folder.</param>
    /// <returns>The initial syncKey of the specified folder.</returns>
    protected string GetInitialSyncKey(string collectionId)
    {
        // Obtains the key by sending an initial Sync request with a SyncKey element value of zero and the CollectionId element
        var syncCollection = new Request.SyncCollection
        {
            CollectionId = collectionId,
            SyncKey = "0"
        };

        var syncRequest = Common.CreateSyncRequest([syncCollection]);
        var syncResult = PROVAdapter.Sync(syncRequest);

        Site.Assert.IsNotNull(
            syncResult,
            "The result for an initial synchronize should not null.");

        // Verify sync result.
        Site.Assert.AreEqual<byte>(
            1,
            syncResult.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        return syncResult.SyncKey;
    }

    /// <summary>
    /// Loop to get the recorded emails.
    /// </summary>
    /// <param name="collectionId">The collection id of the specified folder.</param>
    /// <param name="emailSubject">The email subject.</param>
    /// <param name="isRetryNeeded">A boolean indicating whether need retry.</param>
    /// <returns>The ServerId of recorded item.</returns>
    protected Collection<string> SyncCaseRelativeItems(string collectionId, string emailSubject, bool isRetryNeeded)
    {
        // Acknowledge the security policy settings.
        AcknowledgeSecurityPolicySettings();

        // Call FolderSync command to synchronize the collection hierarchy.
        var folderSynReponse = PROVAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

        // Verify FolderSync command response.
        Site.Assert.AreEqual<int>(
            1,
            int.Parse(folderSynReponse.ResponseData.Status),
            "If the FolderSync command executes successfully, the Status in response should be 1.");

        // Loop to get the specified email.
        var request = CreateSyncRequest(GetInitialSyncKey(collectionId), collectionId);
        var syncResult = PROVAdapter.SyncEmail(request, emailSubject, isRetryNeeded);

        var itemServerIds = new Collection<string> { syncResult.ServerId };

        return itemServerIds;
    }
    #endregion

    #region Private methods

    /// <summary>
    /// Build a generic Sync request without command references by using the specified sync key, folder collection ID.
    /// </summary>
    /// <param name="syncKey">The current sync key.</param>
    /// <param name="collectionId">The collection id which to sync with.</param>
    /// <returns>A Sync command request.</returns>
    private static SyncRequest CreateSyncRequest(string syncKey, string collectionId)
    {
        var syncCollection = new Request.SyncCollection
        {
            SyncKey = syncKey,
            CollectionId = collectionId
        };

        var items = new List<object>();
        var itemsElementName = new List<Request.ItemsChoiceType1>();

        var bodyPreference = new Request.BodyPreference()
        {
            Type = 1,
            TruncationSize = 2000,
            TruncationSizeSpecified = true,
            AllOrNone = true,
            AllOrNoneSpecified = true
        };

        items.Add(bodyPreference);
        itemsElementName.Add(Request.ItemsChoiceType1.BodyPreference);

        syncCollection.Options =
        [
            new Request.Options()
            {
                ItemsElementName = itemsElementName.ToArray(),
                Items = items.ToArray()
            }
        ];

        return Common.CreateSyncRequest([syncCollection]);
    }

    /// <summary>
    /// Delete the specified item.
    /// </summary>
    /// <param name="itemsToDelete">The collection of the items to delete.</param>
    private void DeleteCreatedItems(Collection<CreatedItems> itemsToDelete)
    {
        foreach (var itemToDelete in itemsToDelete)
        {
            var itemsToDeleteServerIds = SyncCaseRelativeItems(itemToDelete.CollectionId, itemToDelete.ItemSubject[0], true);

            var syncRequest = CreateSyncRequest(GetInitialSyncKey(itemToDelete.CollectionId), itemToDelete.CollectionId);
            var syncResponse = PROVAdapter.Sync(syncRequest);
            Site.Assert.AreNotEqual<int>(0, syncResponse.AddElements.Count, "There is not items added in {0} folder.", itemToDelete.CollectionId);

            var deleteData = new List<Request.SyncCollectionDelete>();
            var syncCollection = new Request.SyncCollection();

            foreach (var serverId in itemsToDeleteServerIds)
            {
                deleteData.Add(new Request.SyncCollectionDelete() { ServerId = serverId });
            }

            syncCollection.Commands = deleteData.ToArray();
            syncCollection.SyncKey = syncResponse.SyncKey;
            syncCollection.CollectionId = itemToDelete.CollectionId;
            syncCollection.DeletesAsMoves = false;
            syncCollection.DeletesAsMovesSpecified = true;

            syncRequest = Common.CreateSyncRequest([syncCollection]);
            var deleteResult = PROVAdapter.Sync(syncRequest);

            Site.Assert.AreEqual<byte>(
                1,
                deleteResult.CollectionStatus,
                "The value of Status should be 1 to indicate the Sync command executes successfully.");
        }
    }

    #endregion
}