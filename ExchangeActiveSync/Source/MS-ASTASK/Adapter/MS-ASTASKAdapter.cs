namespace Microsoft.Protocols.TestSuites.MS_ASTASK;

using System.Xml.XPath;
using Common;
using Common.DataStructures;
using TestTools;
using Request = Common.Request;

/// <summary>
/// Adapter class of MS-ASTASK.
/// </summary>
public partial class MS_ASTASKAdapter : ManagedAdapterBase, IMS_ASTASKAdapter
{
    #region Private field

    /// <summary>
    /// The instance of ActiveSync client.
    /// </summary>
    private ActiveSyncClient activeSyncClient;

    #endregion

    #region IMS_ASTASKAdapter Properties

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

    #endregion

    #region IMS_ASTASKAdapter Initialize method

    /// <summary>
    /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
    /// </summary>
    /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
    public override void Initialize(ITestSite testSite)
    {
        base.Initialize(testSite);
        Site.DefaultProtocolDocShortName = "MS-ASTASK";

        // Merge configuration.
        Common.MergeConfiguration(testSite);

        activeSyncClient = new ActiveSyncClient(testSite)
        {
            UserName = Common.GetConfigurationPropertyValue("UserName", Site),
            Password = Common.GetConfigurationPropertyValue("Password", Site)
        };
    }

    #endregion

    #region IMS_ASTASKAdapter Commands

    /// <summary>
    /// Sync data from the server.
    /// </summary>
    /// <param name="syncRequest">A Sync command request.</param>
    /// <returns>A Sync command response returned from the server.</returns>
    public SyncStore Sync(SyncRequest syncRequest)
    {
        var response = activeSyncClient.Sync(syncRequest, true);
        Site.Assert.IsNotNull(response, "The Sync response should be returned.");
        VerifyTransport();
        VerifyWBXMLRequirements();

        var syncResponse = Common.LoadSyncResponse(response);

        foreach (var collection in syncRequest.RequestData.Collections)
        {
            if (collection.SyncKey != "0")
            {
                VerifyMessageSyntax();
                VerifySyncCommandResponse(syncResponse);
            }
        }

        return syncResponse;
    }

    /// <summary>
    /// Synchronize the collection hierarchy.
    /// </summary>
    /// <returns>A FolderSync command response returned from the server.</returns>
    public FolderSyncResponse FolderSync()
    {
        var request = Common.CreateFolderSyncRequest("0");
        var folderSyncResponse = activeSyncClient.FolderSync(request);
        Site.Assert.IsNotNull(folderSyncResponse, "The FolderSync response should be returned.");

        return folderSyncResponse;
    }

    /// <summary>
    /// Search data using the given keyword text.
    /// </summary>
    /// <param name="searchRequest">A Search command request.</param>
    /// <returns>A Search command response returned from the server.</returns>
    public SearchStore Search(SearchRequest searchRequest)
    {
        var response = activeSyncClient.Search(searchRequest, true);
        Site.Assert.IsNotNull(response, "The Search response should be returned.");
        var searchResponse = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site));
        VerifySearchCommandResponse(searchResponse);

        return searchResponse;
    }

    /// <summary>
    /// Fetch all the data about tasks.
    /// </summary>
    /// <param name="itemOperationsRequest">An ItemOperations command request.</param>
    /// <returns>An ItemOperations command response returned from the server.</returns>
    public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
    {
        var response = activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
        Site.Assert.IsNotNull(response, "The ItemOperations response should be returned.");
        var itemOperationResponse = Common.LoadItemOperationsResponse(response);
        VerifyItemOperationsResponse(itemOperationResponse);

        return itemOperationResponse;
    }

    /// <summary>
    /// Send a string request and get a response from server.
    /// </summary>
    /// <param name="stringRequest">A string request for a certain command.</param>
    /// <param name="commandName">Commands choices.</param>
    /// <returns>A string response returned from the server.</returns>
    public SendStringResponse SendStringRequest(string stringRequest, CommandName commandName)
    {
        var response = activeSyncClient.SendStringRequest(commandName, null, stringRequest);
        Site.Assert.IsNotNull(response, "The string response should be returned.");

        return response;
    }

    #endregion
}