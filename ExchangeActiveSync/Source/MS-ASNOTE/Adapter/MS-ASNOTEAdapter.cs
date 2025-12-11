namespace Microsoft.Protocols.TestSuites.MS_ASNOTE;

using System.Xml.XPath;
using Common;
using Common.DataStructures;
using TestTools;

/// <summary>
/// Adapter class of MS-ASNOTE.
/// </summary>
public partial class MS_ASNOTEAdapter : ManagedAdapterBase, IMS_ASNOTEAdapter
{
    #region private field
    /// <summary>
    /// The instance of ActiveSync client.
    /// </summary>
    private ActiveSyncClient activeSyncClient;

    #endregion

    #region IMS_ASNOTEAdapter Properties
    /// <summary>
    /// Gets the XML request sent to protocol SUT
    /// </summary>
    public IXPathNavigable LastRawRequestXml
    {
        get { return activeSyncClient.LastRawRequestXml; }
    }

    /// <summary>
    /// Gets the XML response received from protocol SUT
    /// </summary>
    public IXPathNavigable LastRawResponseXml
    {
        get { return activeSyncClient.LastRawResponseXml; }
    }
    #endregion

    /// <summary>
    /// Sync data from the server
    /// </summary>
    /// <param name="syncRequest">Sync command request.</param>
    /// <param name="isResyncNeeded">A bool value indicates whether need to re-sync when the response contains MoreAvailable element.</param>
    /// <returns>The sync result which is returned from server</returns>
    public SyncStore Sync(SyncRequest syncRequest, bool isResyncNeeded)
    {
        var response = activeSyncClient.Sync(syncRequest, isResyncNeeded);
        VerifySyncResponse(response);
        var result = Common.LoadSyncResponse(response);
        VerifyTransport();
        VerifySyncResult(result);
        VerifyWBXMLCapture();
        return result;
    }

    /// <summary>
    /// Synchronizes the collection hierarchy
    /// </summary>
    /// <param name="folderSyncRequest">FolderSync command request.</param>
    /// <returns>The FolderSync response which is returned from the server</returns>
    public FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest)
    {
        var response = activeSyncClient.FolderSync(folderSyncRequest);
        Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
        return response;
    }

    /// <summary>
    /// Loop to get the results of the specific query request by Search command.
    /// </summary>
    /// <param name="collectionId">The CollectionId of the folder to search.</param>
    /// <param name="subject">The subject of the note to get.</param>
    /// <param name="isLoopNeeded">A boolean value specify whether need the loop</param>
    /// <param name="expectedCount">The expected number of the note to be found.</param>
    /// <returns>The results in response of Search command</returns>
    public SearchStore Search(string collectionId, string subject, bool isLoopNeeded, int expectedCount)
    {
        var searchRequest = Common.CreateSearchRequest(subject, collectionId);
        var response = activeSyncClient.Search(searchRequest, isLoopNeeded, expectedCount);
        VerifySearchResponse(response);
        var result = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site));
        VerifyTransport();
        VerifySearchResult(result);
        VerifyWBXMLCapture();
        return result;
    }

    /// <summary>
    /// Fetch all information about exchange object
    /// </summary>
    /// <param name="itemOperationsRequest">ItemOperations command request.</param>
    /// <returns>The ItemOperations result which is returned from server</returns>
    public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
    {
        var response = activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
        VerifyItemOperationsResponse(response);
        var result = Common.LoadItemOperationsResponse(response);
        var hasSchemaElement = AdapterHelper.ContainsSchemaElement(itemOperationsRequest);
        VerifyTransport();
        VerifyItemOperationResult(result, hasSchemaElement);
        VerifyWBXMLCapture();
        return result;
    }

    /// <summary>
    /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
    /// </summary>
    /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
    public override void Initialize(ITestSite testSite)
    {
        base.Initialize(testSite);
        testSite.DefaultProtocolDocShortName = "MS-ASNOTE";

        // Merge the common configuration
        Common.MergeConfiguration(testSite);
        activeSyncClient = new ActiveSyncClient(testSite)
        {
            UserName = Common.GetConfigurationPropertyValue("UserName", testSite),
            Password = Common.GetConfigurationPropertyValue("UserPassword", testSite)
        };
    }
}
