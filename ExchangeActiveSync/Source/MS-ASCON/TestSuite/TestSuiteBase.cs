namespace Microsoft.Protocols.TestSuites.MS_ASCON;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Common;
using Common.DataStructures;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Request = Common.Request;
using Response = Common.Response;

/// <summary>
/// The base class of scenario class.
/// </summary>
[TestClass]
public class TestSuiteBase : TestClassBase
{
    #region Variables
    /// <summary>
    /// Gets or sets the information of User1.
    /// </summary>
    protected UserInformation User1Information { get; set; }

    /// <summary>
    /// Gets or sets the information of User2.
    /// </summary>
    protected UserInformation User2Information { get; set; }

    /// <summary>
    /// Gets or sets the information of User3.
    /// </summary>
    protected UserInformation User3Information { get; set; }

    /// <summary>
    /// Gets MS-ASCON protocol adapter.
    /// </summary>
    protected IMS_ASCONAdapter CONAdapter { get; private set; }

    /// <summary>
    /// Gets the latest SyncKey.
    /// </summary>
    protected string LatestSyncKey { get; private set; }
    #endregion

    /// <summary>
    /// Record the user name, folder collectionId and subjects the current test case impacts.
    /// </summary>
    /// <param name="userInformation">The information of the user.</param>
    /// <param name="folderCollectionId">The collectionId of folders that the current test case impacts.</param>
    /// <param name="itemSubject">The subject of items that the current test case impacts.</param>
    /// <param name="isDeleted">Whether the item has been deleted and should be removed from the record.</param>
    protected static void RecordCaseRelativeItems(UserInformation userInformation, string folderCollectionId, string itemSubject, bool isDeleted)
    {
        // Record the item in the specified folder.
        var items = new CreatedItems { CollectionId = folderCollectionId };
        items.ItemSubject.Add(itemSubject);
        var isSame = false;

        if (!isDeleted)
        {
            if (userInformation.UserCreatedItems.Count > 0)
            {
                foreach (var createdItems in userInformation.UserCreatedItems)
                {
                    if (createdItems.CollectionId == folderCollectionId && createdItems.ItemSubject[0] == itemSubject)
                    {
                        isSame = true;
                    }
                }

                if (!isSame)
                {
                    userInformation.UserCreatedItems.Add(items);
                }
            }
            else
            {
                userInformation.UserCreatedItems.Add(items);
            }
        }
        else
        {
            if (userInformation.UserCreatedItems.Count > 0)
            {
                foreach (var existItem in userInformation.UserCreatedItems)
                {
                    if (existItem.CollectionId == folderCollectionId && existItem.ItemSubject[0] == itemSubject)
                    {
                        userInformation.UserCreatedItems.Remove(existItem);
                        break;
                    }
                }
            }
        }
    }

    #region Test case initialize and cleanup
    /// <summary>
    /// Initialize the Test suite.
    /// </summary>
    protected override void TestInitialize()
    {
        base.TestInitialize();
        CONAdapter = Site.GetAdapter<IMS_ASCONAdapter>();

        // If implementation doesn't support this specification [MS-ASCON], the case will not start.
        if (!bool.Parse(Common.GetConfigurationPropertyValue("MS-ASCON_Supported", Site)))
        {
            Site.Assert.Inconclusive("This test suite is not supported under current SUT, because MS-ASCON_Supported value is set to false in MS-ASCON_{0}_SHOULDMAY.deployment.ptfconfig file.", Common.GetSutVersion(Site));
        }

        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The airsyncbase:BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        User1Information = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User1Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User1Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        User2Information = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User2Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User2Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        User3Information = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User3Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User3Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        if (Common.GetSutVersion(Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "12.1"))
        {
            // Switch the current user to User1 and synchronize the folder hierarchy of User1.
            SwitchUser(User1Information, true);
        }
    }

    /// <summary>
    /// Clean up the environment.
    /// </summary>
    protected override void TestCleanup()
    {
        // If implementation doesn't support this specification [MS-ASCON], the case will not start.
        if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASCON_Supported", Site)))
        {
            if (User1Information.UserCreatedItems.Count != 0)
            {
                // Switch to User1
                SwitchUser(User1Information, false);
                DeleteItemsInFolder(User1Information.UserCreatedItems);
            }

            if (User2Information.UserCreatedItems.Count != 0)
            {
                // Switch to User2
                SwitchUser(User2Information, false);
                DeleteItemsInFolder(User2Information.UserCreatedItems);
            }

            if (User3Information.UserCreatedItems.Count != 0)
            {
                // Switch to User3
                SwitchUser(User3Information, false);
                DeleteItemsInFolder(User3Information.UserCreatedItems);
            }
        }

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
        CONAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

        if (isFolderSyncNeeded)
        {
            // Call FolderSync command to synchronize the collection hierarchy.
            var folderSyncRequest = Common.CreateFolderSyncRequest("0");
            var folderSyncResponse = CONAdapter.FolderSync(folderSyncRequest);

            // Verify FolderSync command response.
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderSyncResponse.ResponseData.Status),
                "If the FolderSync command executes successfully, the Status in response should be 1.");

            // Get the folder collectionId of User1
            if (userInformation.UserName == User1Information.UserName)
            {
                if (string.IsNullOrEmpty(User1Information.InboxCollectionId))
                {
                    User1Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, Site);
                }

                if (string.IsNullOrEmpty(User1Information.DeletedItemsCollectionId))
                {
                    User1Information.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, Site);
                }

                if (string.IsNullOrEmpty(User1Information.CalendarCollectionId))
                {
                    User1Information.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, Site);
                }

                if (string.IsNullOrEmpty(User1Information.SentItemsCollectionId))
                {
                    User1Information.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, Site);
                }

                if (string.IsNullOrEmpty(User1Information.RecipientInformationCacheCollectionId))
                {
                    User1Information.RecipientInformationCacheCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.RecipientInformationCache, Site);
                }
            }

            // Get the folder collectionId of User2
            if (userInformation.UserName == User2Information.UserName)
            {
                if (string.IsNullOrEmpty(User2Information.InboxCollectionId))
                {
                    User2Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, Site);
                }
            }

            // Get the folder collectionId of User3
            if (userInformation.UserName == User3Information.UserName)
            {
                if (string.IsNullOrEmpty(User3Information.InboxCollectionId))
                {
                    User3Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, Site);
                }
            }
        }
    }

    /// <summary>
    /// Create a conversation.
    /// </summary>
    /// <param name="subject">The subject of the emails in the conversation.</param>
    /// <returns>The created conversation item.</returns>
    protected ConversationItem CreateConversation(string subject)
    {
        #region Send email from User2 to User1
        SwitchUser(User2Information, true);
        var user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
        var user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
        CallSendMailCommand(user2MailboxAddress, user1MailboxAddress, subject, null);
        RecordCaseRelativeItems(User1Information, User1Information.InboxCollectionId, subject, false);
        #endregion

        #region SmartReply the received email from User1 to User2.
        SwitchUser(User1Information, false);
        var syncResult = SyncEmail(subject, User1Information.InboxCollectionId, true, null, null);

        CallSmartReplyCommand(syncResult.ServerId, User1Information.InboxCollectionId, user1MailboxAddress, user2MailboxAddress, subject);
        RecordCaseRelativeItems(User2Information, User2Information.InboxCollectionId, subject, false);
        #endregion

        #region SmartReply the received email from User2 to User1.
        SwitchUser(User2Information, false);
        syncResult = SyncEmail(subject, User2Information.InboxCollectionId, true, null, null);
        CallSmartReplyCommand(syncResult.ServerId, User2Information.InboxCollectionId, user2MailboxAddress, user1MailboxAddress, subject);
        #endregion

        #region Switch current user to User1 and get the conversation item.
        SwitchUser(User1Information, false);

        var counter = 0;
        int itemsCount;
        var retryLimit = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        do
        {
            Thread.Sleep(waitTime);
            var syncStore = CallSyncCommand(User1Information.InboxCollectionId, false);

            // Reset the item count.
            itemsCount = 0;

            foreach (var item in syncStore.AddElements)
            {
                if (item.Email.Subject.Contains(subject))
                {
                    syncResult = item;
                    itemsCount++;
                }
            }

            counter++;
        }
        while (itemsCount < 2 && counter < retryLimit);

        Site.Assert.AreEqual<int>(2, itemsCount, "There should be 2 emails with subject {0} in the Inbox folder, actual {1}.", subject, itemsCount);

        return GetConversationItem(User1Information.InboxCollectionId, syncResult.Email.ConversationId);
        #endregion
    }

    /// <summary>
    /// Call Search command to find a specified conversation.
    /// </summary>
    /// <param name="conversationId">The ConversationId of the items to search.</param>
    /// <param name="itemsCount">The count of the items expected to be found.</param>
    /// <param name="bodyPartPreference">The BodyPartPreference element.</param>
    /// <param name="bodyPreference">The BodyPreference element.</param>
    /// <returns>The SearchStore instance that contains the search result.</returns>
    protected SearchStore CallSearchCommand(string conversationId, int itemsCount, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
    {
        // Create Search command request.
        var searchRequest = TestSuiteHelper.GetSearchRequest(conversationId, bodyPartPreference, bodyPreference);
        var searchStore = CONAdapter.Search(searchRequest, true, itemsCount);

        Site.Assert.AreEqual("1", searchStore.Status, "The Search operation should be success.");

        return searchStore;
    }

    /// <summary>
    /// Call SendMail command to send mail.
    /// </summary>
    /// <param name="from">The mailbox address of sender.</param>
    /// <param name="to">The mailbox address of recipient.</param>
    /// <param name="subject">The subject of the email.</param>
    /// <param name="body">The body content of the email.</param>
    protected void CallSendMailCommand(string from, string to, string subject, string body)
    {
        if (string.IsNullOrEmpty(body))
        {
            body = Common.GenerateResourceName(Site, "body");
        }

        // Create the SendMail command request.
        var template =
            @"From: {0}
To: {1}
Subject: {2}
Content-Type: text/html; charset=""us-ascii""
MIME-Version: 1.0

<html>
<body>
<font color=""blue"">{3}</font>
</body>
</html>
";

        var mime = Common.FormatString(template, from, to, subject, body);
        var sendMailRequest = Common.CreateSendMailRequest(null, Guid.NewGuid().ToString(), mime);

        // Call SendMail command.
        var sendMailResponse = CONAdapter.SendMail(sendMailRequest);

        Site.Assert.AreEqual(string.Empty, sendMailResponse.ResponseDataXML, "The SendMail command should be executed successfully.");
    }

    /// <summary>
    /// Call SmartReply command to reply an email.
    /// </summary>
    /// <param name="itemServerId">The ServerId of the email to reply.</param>
    /// <param name="collectionId">The folder collectionId of the source email.</param>
    /// <param name="from">The mailbox address of sender.</param>
    /// <param name="replyTo">The mailbox address of recipient.</param>
    /// <param name="subject">The subject of the email to reply.</param>
    protected void CallSmartReplyCommand(string itemServerId, string collectionId, string from, string replyTo, string subject)
    {
        // Create SmartReply command request.
        var source = new Request.Source();
        var mime = Common.CreatePlainTextMime(from, replyTo, string.Empty, string.Empty, subject, "SmartReply content");
        var smartReplyRequest = Common.CreateSmartReplyRequest(null, Guid.NewGuid().ToString(), mime, source);

        // Set the command parameters.
        smartReplyRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>());

        source.FolderId = collectionId;
        source.ItemId = itemServerId;
        smartReplyRequest.CommandParameters.Add(CmdParameterName.CollectionId, collectionId);
        smartReplyRequest.CommandParameters.Add(CmdParameterName.ItemId, itemServerId);
        smartReplyRequest.RequestData.ReplaceMime = string.Empty;

        // Call SmartReply command.
        var smartReplyResponse = CONAdapter.SmartReply(smartReplyRequest);

        Site.Assert.AreEqual(string.Empty, smartReplyResponse.ResponseDataXML, "The SmartReply command should be executed successfully.");
    }

    /// <summary>
    /// Call SmartForward command to forward an email.
    /// </summary>
    /// <param name="itemServerId">The ServerId of the email to reply.</param>
    /// <param name="collectionId">The folder collectionId of the source email.</param>
    /// <param name="from">The mailbox address of sender.</param>
    /// <param name="forwardTo">The mailbox address of recipient.</param>
    /// <param name="subject">The subject of the email to reply.</param>
    protected void CallSmartForwardCommand(string itemServerId, string collectionId, string from, string forwardTo, string subject)
    {
        // Create SmartForward command request.
        var source = new Request.Source();
        var mime = Common.CreatePlainTextMime(from, forwardTo, string.Empty, string.Empty, subject, "SmartForward content");
        var smartForwardRequest = Common.CreateSmartForwardRequest(null, Guid.NewGuid().ToString(), mime, source);

        // Set the command parameters.
        smartForwardRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>());

        source.FolderId = collectionId;
        source.ItemId = itemServerId;
        smartForwardRequest.CommandParameters.Add(CmdParameterName.CollectionId, collectionId);
        smartForwardRequest.CommandParameters.Add(CmdParameterName.ItemId, itemServerId);
        smartForwardRequest.RequestData.ReplaceMime = string.Empty;

        // Call SmartForward command.
        var smartForwardResponse = CONAdapter.SmartForward(smartForwardRequest);

        Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "The SmartForward command should be executed successfully.");
    }

    /// <summary>
    /// Call ItemOperations command to fetch an email in the specific folder.
    /// </summary>
    /// <param name="collectionId">The folder collection id to be fetched.</param>
    /// <param name="serverId">The ServerId of the item</param>
    /// <param name="bodyPartPreference">The BodyPartPreference element.</param>
    /// <param name="bodyPreference">The bodyPreference element.</param>
    /// <returns>An Email instance that includes the fetch result.</returns>
    protected Email ItemOperationsFetch(string collectionId, string serverId, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
    {
        var itemOperationsRequest = TestSuiteHelper.GetItemOperationsRequest(collectionId, serverId, bodyPartPreference, bodyPreference);
        var itemOperationsResponse = CONAdapter.ItemOperations(itemOperationsRequest);
        Site.Assert.AreEqual("1", itemOperationsResponse.ResponseData.Status, "The ItemOperations operation should be success.");

        var itemOperationsStore = Common.LoadItemOperationsResponse(itemOperationsResponse);
        Site.Assert.AreEqual(1, itemOperationsStore.Items.Count, "Only one email is supposed to be fetched.");
        Site.Assert.AreEqual("1", itemOperationsStore.Items[0].Status, "The fetch result should be success.");
        Site.Assert.IsNotNull(itemOperationsStore.Items[0].Email, "The fetched email should not be null.");

        return itemOperationsStore.Items[0].Email;
    }

    /// <summary>
    /// Call ItemOperations command to move a conversation to a folder.
    /// </summary>
    /// <param name="conversationId">The Id of conversation to be moved.</param>
    /// <param name="destinationFolder">The destination folder id.</param>
    /// <param name="moveAlways">Should future messages always be moved.</param>
    /// <returns>An instance of the ItemOperationsResponse.</returns>
    protected ItemOperationsResponse ItemOperationsMove(string conversationId, string destinationFolder, bool moveAlways)
    {
        var move = new Request.ItemOperationsMove
        {
            DstFldId = destinationFolder,
            ConversationId = conversationId
        };

        if (moveAlways)
        {
            move.Options = new Request.ItemOperationsMoveOptions { MoveAlways = string.Empty };
        }

        var itemOperationRequest = Common.CreateItemOperationsRequest([move]);
        var itemOperationResponse = CONAdapter.ItemOperations(itemOperationRequest);

        Site.Assert.AreEqual("1", itemOperationResponse.ResponseData.Status, "The ItemOperations operation should be success.");
        Site.Assert.AreEqual(1, itemOperationResponse.ResponseData.Response.Move.Length, "The server should return one Move element in ItemOperationsResponse.");

        return itemOperationResponse;
    }

    /// <summary>
    /// Call Sync command to add items to the specified folder.
    /// </summary>
    /// <param name="collectionId">The CollectionId of the specified folder.</param>
    /// <param name="subject">Subject of the item to add.</param>
    /// <param name="syncKey">The latest SyncKey.</param>
    protected void SyncAdd(string collectionId, string subject, string syncKey)
    {
        // Create Sync request.
        var add = new Request.SyncCollectionAdd
        {
            ClientId = Guid.NewGuid().ToString(),
            ApplicationData = new Request.SyncCollectionAddApplicationData
            {
                Items = [subject],
                ItemsElementName = [Request.ItemsChoiceType8.Subject2]
            }
        };

        var collection = new Request.SyncCollection
        {
            Commands = [add],
            CollectionId = collectionId,
            SyncKey = syncKey
        };
        var syncRequest = Common.CreateSyncRequest([collection]);

        // Call Sync command to add the item.
        var syncStore = CONAdapter.Sync(syncRequest);

        // Verify Sync command response.
        Site.Assert.AreEqual<byte>(
            1,
            syncStore.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        LatestSyncKey = syncStore.SyncKey;
    }

    /// <summary>
    /// Call Sync command to change the status of the emails in Inbox folder.
    /// </summary>
    /// <param name="syncKey">The latest SyncKey.</param>
    /// <param name="serverIds">The collection of ServerIds.</param>
    /// <param name="collectionId">The folder collectionId which needs to be sychronized.</param>
    /// <param name="read">Read element of the item.</param> 
    /// <param name="status">Flag status of the item.</param> 
    /// <returns>The SyncStore instance returned from Sync command.</returns>
    protected SyncStore SyncChange(string syncKey, Collection<string> serverIds, string collectionId, bool? read, string status)
    {
        var changes = new List<Request.SyncCollectionChange>();

        foreach (var serverId in serverIds)
        {
            var change = new Request.SyncCollectionChange
            {
                ServerId = serverId,
                ApplicationData = new Request.SyncCollectionChangeApplicationData()
            };

            var changeItems = new List<object>();
            var changeItemsElementName = new List<Request.ItemsChoiceType7>();

            if (read != null)
            {
                changeItems.Add(read);
                changeItemsElementName.Add(Request.ItemsChoiceType7.Read);
            }

            if (!string.IsNullOrEmpty(status))
            {
                var flag = new Request.Flag();
                if (status == "1")
                {
                    // The Complete Time format is yyyy-MM-ddThh:mm:ss.fffZ.
                    flag.CompleteTime = DateTime.Now.ToUniversalTime();
                    flag.CompleteTimeSpecified = true;
                    flag.DateCompleted = DateTime.Now.ToUniversalTime();
                    flag.DateCompletedSpecified = true;
                }

                flag.Status = status;
                flag.FlagType = "Flag for follow up";

                changeItems.Add(flag);
                changeItemsElementName.Add(Request.ItemsChoiceType7.Flag);
            }

            change.ApplicationData.Items = changeItems.ToArray();
            change.ApplicationData.ItemsElementName = changeItemsElementName.ToArray();

            changes.Add(change);
        }

        var collection = new Request.SyncCollection
        {
            CollectionId = collectionId,
            SyncKey = syncKey,
            Commands = changes.ToArray()
        };

        var syncRequest = Common.CreateSyncRequest([collection]);
        var syncStore = CONAdapter.Sync(syncRequest);

        // Verify Sync command response.
        Site.Assert.AreEqual<byte>(
            1,
            syncStore.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        LatestSyncKey = syncStore.SyncKey;

        return syncStore;
    }

    /// <summary>
    /// Call Sync command to delete items.
    /// </summary>
    /// <param name="collectionId">The CollectionId of the folder.</param>
    /// <param name="syncKey">The latest SyncKey.</param> 
    /// <param name="serverIds">The ServerId of the items to delete.</param> 
    /// <returns>The SyncStore instance returned from Sync command.</returns>
    protected SyncStore SyncDelete(string collectionId, string syncKey, string[] serverIds)
    {
        var deleteCollection = new List<Request.SyncCollectionDelete>();
        foreach (var itemId in serverIds)
        {
            var delete = new Request.SyncCollectionDelete { ServerId = itemId };
            deleteCollection.Add(delete);
        }

        var collection = new Request.SyncCollection
        {
            Commands = deleteCollection.ToArray(),
            DeletesAsMoves = true,
            DeletesAsMovesSpecified = true,
            CollectionId = collectionId,
            SyncKey = syncKey
        };

        var syncRequest = Common.CreateSyncRequest([collection]);

        var syncStore = CONAdapter.Sync(syncRequest);

        // Verify Sync command response.
        Site.Assert.AreEqual<byte>(
            1,
            syncStore.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        return syncStore;
    }

    /// <summary>
    /// Find the specified email.
    /// </summary>
    /// <param name="subject">The subject of the email to find.</param>
    /// <param name="collectionId">The folder collectionId which needs to be synchronized.</param>
    /// <param name="isRetryNeeded">A Boolean value indicates whether need retry.</param>
    /// <param name="bodyPartPreference">The bodyPartPreference in the options element.</param>
    /// <param name="bodyPreference">The bodyPreference in the options element.</param>
    /// <returns>The found email object.</returns>
    protected Sync SyncEmail(string subject, string collectionId, bool isRetryNeeded, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
    {
        // Call initial Sync command.
        var syncRequest = Common.CreateInitialSyncRequest(collectionId);
        var syncStore = CONAdapter.Sync(syncRequest);

        // Find the specific email.
        syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, syncStore.SyncKey, bodyPartPreference, bodyPreference, false);
        var syncResult = CONAdapter.SyncEmail(syncRequest, subject, isRetryNeeded);

        LatestSyncKey = syncStore.SyncKey;

        return syncResult;
    }

    /// <summary>
    /// Sync items in the specified folder.
    /// </summary>
    /// <param name="collectionId">The CollectionId of the folder.</param>
    /// <param name="conversationMode">The value of ConversationMode element.</param>
    /// <returns>A SyncStore instance that contains the result.</returns>
    protected SyncStore CallSyncCommand(string collectionId, bool conversationMode)
    {
        // Call initial Sync command.
        var syncRequest = Common.CreateInitialSyncRequest(collectionId);

        var syncStore = CONAdapter.Sync(syncRequest);

        // Verify Sync command response.
        Site.Assert.AreEqual<byte>(
            1,
            syncStore.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        if (conversationMode && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) != "12.1")
        {
            syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, syncStore.SyncKey, null, null, true);
        }
        else
        {
            syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, syncStore.SyncKey, null, null, false);
        }

        syncStore = CONAdapter.Sync(syncRequest);

        // Verify Sync command response.
        Site.Assert.AreEqual<byte>(
            1,
            syncStore.CollectionStatus,
            "If the Sync command executes successfully, the Status in response should be 1.");

        var checkSyncStore = syncStore.AddElements != null && syncStore.AddElements.Count != 0;
        Site.Assert.IsTrue(checkSyncStore, "The items should be gotten from the Sync command response.");

        LatestSyncKey = syncStore.SyncKey;

        return syncStore;
    }

    /// <summary>
    /// Gets an estimate of the number of items in the specific folder.
    /// </summary>
    /// <param name="syncKey">The latest SyncKey.</param> 
    /// <param name="collectionId">The CollectionId of the folder.</param> 
    /// <returns>The response of GetItemEstimate command.</returns>
    protected GetItemEstimateResponse CallGetItemEstimateCommand(string syncKey, string collectionId)
    {
        var itemsElementName = new List<Request.ItemsChoiceType10>()
        {
            Request.ItemsChoiceType10.SyncKey,
            Request.ItemsChoiceType10.CollectionId,
        };
            
        var items = new List<object>()
        {
            syncKey,
            collectionId,
        };
           
        if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site) != "12.1")
        {
            itemsElementName.Add(Request.ItemsChoiceType10.ConversationMode);
            items.Add(true);
        }
        // Create GetItemEstimate command request.
        var collection = new Request.GetItemEstimateCollection
        {
            ItemsElementName = itemsElementName.ToArray(),
            Items = items.ToArray()
        };

        var getItemEstimateRequest = Common.CreateGetItemEstimateRequest([collection]);

        var getItemEstimateResponse = CONAdapter.GetItemEstimate(getItemEstimateRequest);

        return getItemEstimateResponse;
    }

    /// <summary>
    /// Move items to the specific folder.
    /// </summary>
    /// <param name="serverIds">The ServerId of the items to move.</param> 
    /// <param name="sourceFolder">The CollectionId of the source folder.</param>
    /// <param name="destinationFolder">The CollectionId of the destination folder.</param> 
    /// <returns>The response of MoveItems command.</returns>
    protected MoveItemsResponse CallMoveItemsCommand(Collection<string> serverIds, string sourceFolder, string destinationFolder)
    {
        // Move the items from sourceFolder to destinationFolder.
        var moveItems = new List<Request.MoveItemsMove>();
        foreach (var serverId in serverIds)
        {
            var move = new Request.MoveItemsMove
            {
                SrcFldId = sourceFolder,
                DstFldId = destinationFolder,
                SrcMsgId = serverId
            };

            moveItems.Add(move);
        }

        var moveItemsRequest = Common.CreateMoveItemsRequest(moveItems.ToArray());

        // Call MoveItems command to move the items.
        var moveItemsResponse = CONAdapter.MoveItems(moveItemsRequest);

        Site.Assert.AreEqual<int>(serverIds.Count, moveItemsResponse.ResponseData.Response.Length, "The count of Response element should be {0}, actual {1}.", serverIds.Count, moveItemsResponse.ResponseData.Response.Length);
        foreach (var response in moveItemsResponse.ResponseData.Response)
        {
            Site.Assert.AreEqual<int>(3, int.Parse(response.Status), "If the MoveItems command executes successfully, the Status should be 3, actual {0}.", response.Status);
        }

        return moveItemsResponse;
    }

    /// <summary>
    /// Gets the created ConversationItem.
    /// </summary>
    /// <param name="collectionId">The CollectionId of the parent folder which has the conversation.</param>
    /// <param name="conversationId">The ConversationId of the conversation.</param>
    /// <returns>A ConversationItem object.</returns>
    protected ConversationItem GetConversationItem(string collectionId, string conversationId)
    {
        // Call Sync command to get the emails in Inbox folder.
        var syncStore = CallSyncCommand(collectionId, false);

        // Get the emails from Sync response according to the ConversationId.
        var conversationItem = new ConversationItem { ConversationId = conversationId };

        foreach (var addElement in syncStore.AddElements)
        {
            if (addElement.Email.ConversationId == conversationId)
            {
                conversationItem.ServerId.Add(addElement.ServerId);
            }
        }

        Site.Assert.AreNotEqual<int>(0, conversationItem.ServerId.Count, "The conversation should have at least one email.");

        return conversationItem;
    }

    /// <summary>
    /// Gets the conversation with the expected emails count.
    /// </summary>
    /// <param name="collectionId">The CollectionId of the parent folder which has the conversation.</param>
    /// <param name="conversationId">The ConversationId of the conversation.</param>
    /// <param name="expectEmailCount">The expect count of the conversation emails.</param>
    /// <returns>A ConversationItem object.</returns>
    protected ConversationItem GetConversationItem(string collectionId, string conversationId, int expectEmailCount)
    {
        ConversationItem coversationItem;
        var counter = 0;
        var retryLimit = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        do
        {
            Thread.Sleep(waitTime);
            coversationItem = GetConversationItem(collectionId, conversationId);
            counter++;
        }
        while (coversationItem.ServerId.Count != expectEmailCount && counter < retryLimit);

        return coversationItem;
    }

    /// <summary>
    /// Checks if ActiveSync Protocol Version is not "14.0".
    /// </summary>
    protected void CheckActiveSyncVersionIsNot140()
    {
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The airsyncbase:BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
    }
    #endregion

    #region Capture code
    /// <summary>
    /// Verify the message part when the request contains neither BodyPreference nor BodyPartPreference elements.
    /// </summary>
    /// <param name="email">The email item server returned.</param>
    protected void VerifyMessagePartWithoutPreference(Email email)
    {
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R245");

        // Verify MS-ASCON requirement: MS-ASCON_R245
        var isVerifiedR245 = email.Body != null && email.BodyPart == null;

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR245,
            245,
            @"[In Sending a Message Part] If request contains neither airsyncbase:BodyPreference nor airsyncbase:BodyPartPreference elements, then the response contains only airsyncbase:Body element.");
    }

    /// <summary>
    /// Verify the message part when the request contains only BodyPreference element.
    /// </summary>
    /// <param name="email">The email item server returned.</param>
    protected void VerifyMessagePartWithBodyPreference(Email email)
    {
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R246");

        // Verify MS-ASCON requirement: MS-ASCON_R246
        var isVerifiedR246 = email.Body != null && email.BodyPart == null;

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR246,
            246,
            @"[In Sending a Message Part] If request contains only airsyncbase:BodyPreference element, then the response contains only airsyncbase:Body element.");
    }

    /// <summary>
    /// Verify the message part when the request contains only BodyPartPreference element.
    /// </summary>
    /// <param name="email">The email item server returned.</param>
    /// <param name="truncatedData">The truncated email data returned in BodyPart.</param>
    /// <param name="allData">All email data without being truncated.</param>
    /// <param name="truncationSize">The TruncationSize element specified in BodyPartPreference.</param>
    protected void VerifyMessagePartWithBodyPartPreference(Email email, string truncatedData, string allData, int truncationSize)
    {
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R239");

        // Verify MS-ASCON requirement: MS-ASCON_R239
        var isVerifiedR239 = email.BodyPart.TruncatedSpecified && email.BodyPart.Truncated
                                                               && truncatedData == TestSuiteHelper.TruncateData(allData, truncationSize);

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR239,
            239,
            @"[In Sending a Message Part] The client's preferences affect the server response as follows: If the size of the message part exceeds the value specified in the airsyncbase:TruncationSize element ([MS-ASAIRS] section 2.2.2.40.1) of the request, then the server truncates the message part.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R240");

        // Verify MS-ASCON requirement: MS-ASCON_R240
        var isVerifiedR240 = email.BodyPart.TruncatedSpecified && email.BodyPart.Truncated && email.BodyPart.EstimatedDataSize > 0;

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR240,
            240,
            @"[In Sending a Message Part] The server includes the airsyncbase:Truncated element ([MS-ASAIRS] section 2.2.2.39.1) and the airsyncbase:EstimatedDataSize element ([MS-ASAIRS] section 2.2.2.23.2) in the response when it truncates the message part.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R247");

        // Verify MS-ASCON requirement: MS-ASCON_R247
        var isVerifiedR247 = email.Body == null && email.BodyPart != null;

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR247,
            247,
            @"[In Sending a Message Part] If request contains only airsyncbase:BodyPartPreference element, then the response contains only airsyncbase:BodyPart element.");
    }

    /// <summary>
    /// Verify the message part when the request contains both BodyPreference and BodyPartPreference elements.
    /// </summary>
    /// <param name="email">The email item server returned.</param>
    protected void VerifyMessagePartWithBothPreference(Email email)
    {
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R248");

        // Verify MS-ASCON requirement: MS-ASCON_R248
        var isVerifiedR248 = email.Body != null && email.BodyPart != null;

        Site.CaptureRequirementIfIsTrue(
            isVerifiedR248,
            248,
            @"[In Sending a Message Part] If request contains both airsyncbase:BodyPreference and airsyncbase:BodyPartPreference element, then the response contains both airsyncbase:Body and airsyncbase:BodyPart element.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R243");

        // Verify MS-ASCON requirement: MS-ASCON_R243
        // If R248 is captured, then BodyPart element and Body element do co-exist in the server response.
        Site.CaptureRequirement(
            243,
            @"[In Sending a Message Part] The airsyncbase:BodyPart element and the airsyncbase:Body element ([MS-ASAIRS] section 2.2.2.9) can co-exist in the server response.");
    }

    /// <summary>
    /// Verify status 164 is returned when the Type element in the BodyPartPreference is other than 2.
    /// </summary>
    /// <param name="status">The status that server returned.</param>
    protected void VerifyMessagePartStatus164(int status)
    {
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R241");

        // Verify MS-ASCON requirement: MS-ASCON_R241
        Site.CaptureRequirementIfAreEqual<int>(
            164,
            status,
            241,
            @"[In Sending a Message Part] [The client's preferences affect the server response as follows:] If a value other than 2 is specified in the airsyncbase:Type element ([MS-ASAIRS] section 2.2.2.41.3) of the request, then the server returns a status value of 164.");
    }
    #endregion

    #region Private methods

    /// <summary>
    /// Delete all the items in a folder.
    /// </summary>
    /// <param name="createdItems">The created items which should be deleted.</param>
    private void DeleteItemsInFolder(Collection<CreatedItems> createdItems)
    {
        foreach (var createdItem in createdItems)
        {
            var syncResult = CallSyncCommand(createdItem.CollectionId, false);
            var deleteData = new List<Request.SyncCollectionDelete>();
            var serverIds = new List<string>();

            foreach (var subject in createdItem.ItemSubject)
            {
                if (syncResult != null)
                {
                    foreach (var item in syncResult.AddElements)
                    {
                        if (item.Email.Subject != null && item.Email.Subject.Equals(subject, StringComparison.CurrentCulture))
                        {
                            serverIds.Add(item.ServerId);
                        }

                        if (item.Calendar.Subject != null && item.Calendar.Subject.Equals(subject, StringComparison.CurrentCulture))
                        {
                            serverIds.Add(item.ServerId);
                        }
                    }
                }

                Site.Assert.AreNotEqual<int>(0, serverIds.Count, "The items with subject '{0}' should be found!", subject);

                foreach (var serverId in serverIds)
                {
                    deleteData.Add(new Request.SyncCollectionDelete() { ServerId = serverId });
                }

                var syncCollection = new Request.SyncCollection
                {
                    Commands = deleteData.ToArray(),
                    DeletesAsMoves = false,
                    DeletesAsMovesSpecified = true,
                    CollectionId = createdItem.CollectionId,
                    SyncKey = syncResult.SyncKey
                };

                var syncRequest = Common.CreateSyncRequest([syncCollection]);
                var deleteResult = CONAdapter.Sync(syncRequest);

                Site.Assert.AreEqual<byte>(
                    1,
                    deleteResult.CollectionStatus,
                    "The value of Status should be 1 to indicate the Sync command executed successfully.");
            }
        }
    }

    #endregion
}