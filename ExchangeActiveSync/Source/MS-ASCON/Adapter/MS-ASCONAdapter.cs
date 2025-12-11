namespace Microsoft.Protocols.TestSuites.MS_ASCON;

using System.Threading;
using System.Xml.XPath;
using Common;
using Common.DataStructures;
using TestTools;

/// <summary>
/// Adapter class of MS-ASCON.
/// </summary>
public partial class MS_ASCONAdapter : ManagedAdapterBase, IMS_ASCONAdapter
{
    /// <summary>
    /// The instance of ActiveSync client.
    /// </summary>
    private ActiveSyncClient activeSyncClient;

    /// <summary>
    /// Gets the raw XML request sent to protocol SUT.
    /// </summary>
    public IXPathNavigable LastRawRequestXml
    {
        get { return activeSyncClient.LastRawRequestXml; }
    }

    /// <summary>
    /// Gets the raw XML response received from protocol SUT.
    /// </summary>
    public IXPathNavigable LastRawResponseXml
    {
        get { return activeSyncClient.LastRawResponseXml; }
    }

    /// <summary>
    /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
    /// </summary>
    /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
    public override void Initialize(ITestSite testSite)
    {
        base.Initialize(testSite);
        testSite.DefaultProtocolDocShortName = "MS-ASCON";

        // Merge the common configuration
        Common.MergeConfiguration(testSite);

        activeSyncClient = new ActiveSyncClient(testSite)
        {
            UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
            Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
        };
    }

    /// <summary>
    /// Change the user authentication.
    /// </summary>
    /// <param name="userName">The name of a user.</param>
    /// <param name="userPassword">The password of a user.</param>
    /// <param name="userDomain">The domain which the user belongs to.</param>
    public void SwitchUser(string userName, string userPassword, string userDomain)
    {
        activeSyncClient.UserName = userName;
        activeSyncClient.Password = userPassword;
        activeSyncClient.Domain = userDomain;
    }

    /// <summary>
    /// Synchronizes changes in a collection between the client and the server.
    /// </summary>
    /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
    /// <returns>The SyncStore result which is returned from server.</returns>
    public SyncStore Sync(SyncRequest syncRequest)
    {
        var syncResponse = activeSyncClient.Sync(syncRequest, true);
        Site.Assert.IsNotNull(syncResponse, "The Sync response returned from server should not be null.");

        var syncStore = Common.LoadSyncResponse(syncResponse);

        if (1 == syncStore.CollectionStatus && syncStore.AddElements.Count != 0)
        {
            foreach (var addElement in syncStore.AddElements)
            {
                VerifySyncCommandResponse(addElement);
            }
        }

        // Verify related requirements.
        VerifyCommonRequirements();
        VerifyWBXMLCapture();

        return syncStore;
    }

    /// <summary>
    /// Find an email with specific subject.
    /// </summary>
    /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
    /// <param name="subject">The subject of the email to find.</param>
    /// <param name="isRetryNeeded">A boolean whether need retry.</param>
    /// <returns>The email with specific subject.</returns>
    public Sync SyncEmail(SyncRequest syncRequest, string subject, bool isRetryNeeded)
    {
        var syncResult = activeSyncClient.SyncEmail(syncRequest, subject, isRetryNeeded);

        // Verify related requirements.
        VerifyCommonRequirements();
        VerifyWBXMLCapture();

        // If the email should be gotten, then verify the related requirements.
        if (isRetryNeeded)
        {
            VerifySyncCommandResponse(syncResult);
        }

        return syncResult;
    }

    /// <summary>
    /// Find entries address book, mailbox, or document library.
    /// </summary>
    /// <param name="searchRequest">A SearchRequest object that contains the request information.</param>
    /// <param name="expectSuccess">Whether the Search command is expected to be successful.</param>
    /// <param name="itemsCount">The count of the items expected to be found.</param>
    /// <returns>The SearchStore result which is returned from server.</returns>
    public SearchStore Search(SearchRequest searchRequest, bool expectSuccess, int itemsCount)
    {
        SearchResponse searchResponse;

        if (expectSuccess)
        {
            searchResponse = activeSyncClient.Search(searchRequest, true, itemsCount);
        }
        else
        {
            searchResponse = activeSyncClient.Search(searchRequest);
        }

        var retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        var counter = 1;

        while (counter < retryCount && searchResponse.ResponseData.Status.Equals("10"))
        {
            Thread.Sleep(waitTime);

            if (expectSuccess)
            {
                searchResponse = activeSyncClient.Search(searchRequest, true, itemsCount);
            }
            else
            {
                searchResponse = activeSyncClient.Search(searchRequest);
            }

            counter++;
        }

        Site.Assert.IsNotNull(searchResponse, "The Search response returned from server should not be null.");

        // Verify related requirements.
        VerifyCommonRequirements();
        VerifyWBXMLCapture();

        var searchStore = Common.LoadSearchResponse(searchResponse, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site));

        return searchStore;
    }

    /// <summary>
    /// Acts as a container for the Fetch element, the EmptyFolderContents element, and the Move element to provide batched online handling of these operations against the server.
    /// </summary>
    /// <param name="itemOperationsRequest">An ItemOperationsRequest object that contains the request information.</param>
    /// <returns>ItemOperations command response.</returns>
    public ItemOperationsResponse ItemOperations(ItemOperationsRequest itemOperationsRequest)
    {
        var itemOperationsResponse = activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
        Site.Assert.IsNotNull(itemOperationsResponse, "The ItemOperations response returned from server should not be null.");

        // Verify related requirements.
        VerifyCommonRequirements();
        VerifyWBXMLCapture();
        VerifyItemOperationsCommandResponse(itemOperationsResponse);

        return itemOperationsResponse;
    }

    /// <summary>
    /// Synchronizes the collection hierarchy 
    /// </summary>
    /// <param name="folderSyncRequest">A FolderSyncRequest object that contains the request information.</param>
    /// <returns>FolderSync command response.</returns>
    public FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest)
    {
        var folderSyncResponse = activeSyncClient.FolderSync(folderSyncRequest);
        Site.Assert.IsNotNull(folderSyncResponse, "The FolderSync response returned from server should not be null.");

        return folderSyncResponse;
    }

    /// <summary>
    /// Sends MIME-formatted e-mail messages to the server.
    /// </summary>
    /// <param name="sendMailRequest">A SendMailRequest object that contains the request information.</param>
    /// <returns>SendMail command response.</returns>
    public SendMailResponse SendMail(SendMailRequest sendMailRequest)
    {
        var sendMailResponse = activeSyncClient.SendMail(sendMailRequest);
        Site.Assert.IsNotNull(sendMailResponse, "The SendMail response returned from server should not be null.");

        return sendMailResponse;
    }

    /// <summary>
    /// Replies to messages without retrieving the full, original message from the server.
    /// </summary>
    /// <param name="smartReplyRequest">A SmartReplyRequest object that contains the request information.</param>
    /// <returns>SmartReply command response.</returns>
    public SmartReplyResponse SmartReply(SmartReplyRequest smartReplyRequest)
    {
        var smartReplyResponse = activeSyncClient.SmartReply(smartReplyRequest);
        Site.Assert.IsNotNull(smartReplyResponse, "The SmartReply response returned from server should not be null.");

        return smartReplyResponse;
    }

    /// <summary>
    /// Forwards messages without retrieving the full, original message from the server.
    /// </summary>
    /// <param name="smartForwardRequest">A SmartForwardRequest object that contains the request information.</param>
    /// <returns>SmartForward command response.</returns>
    public SmartForwardResponse SmartForward(SmartForwardRequest smartForwardRequest)
    {
        var smartForwardResponse = activeSyncClient.SmartForward(smartForwardRequest);
        Site.Assert.IsNotNull(smartForwardResponse, "The SmartForward response returned from server should not be null.");

        return smartForwardResponse;
    }

    /// <summary>
    /// Moves an item or items from one folder on the server to another.
    /// </summary>
    /// <param name="moveItemsRequest">A MoveItemsRequest object that contains the request information.</param>
    /// <returns>MoveItems command response.</returns>
    public MoveItemsResponse MoveItems(MoveItemsRequest moveItemsRequest)
    {
        var moveItemsResponse = activeSyncClient.MoveItems(moveItemsRequest);
        Site.Assert.IsNotNull(moveItemsResponse, "The MoveItems response returned from server should not be null.");

        // Verify related requirements.
        VerifyCommonRequirements();
        VerifyWBXMLCapture();

        return moveItemsResponse;
    }

    /// <summary>
    /// Gets an estimate of the number of items in a collection or folder on the server that have to be synchronized.
    /// </summary>
    /// <param name="getItemEstimateRequest">A GetItemEstimateRequest object that contains the request information.</param>
    /// <returns>GetItemEstimate command response.</returns>
    public GetItemEstimateResponse GetItemEstimate(GetItemEstimateRequest getItemEstimateRequest)
    {
        var getItemEstimateResponse = activeSyncClient.GetItemEstimate(getItemEstimateRequest);
        Site.Assert.IsNotNull(getItemEstimateResponse, "The GetItemEstimate response returned from server should not be null.");

        // Verify related requirements.
        VerifyCommonRequirements();

        return getItemEstimateResponse;
    }
}