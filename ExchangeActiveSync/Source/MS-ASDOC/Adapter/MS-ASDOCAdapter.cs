namespace Microsoft.Protocols.TestSuites.MS_ASDOC;

using System.Net;
using System.Xml.XPath;
using Common;
using TestTools;

/// <summary>
/// Adapter class of MS-ASDOC. 
/// </summary>
public partial class MS_ASDOCAdapter : ManagedAdapterBase, IMS_ASDOCAdapter
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
        Site.DefaultProtocolDocShortName = "MS-ASDOC";

        // Get the name of common configuration file.
        var commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSite);

        // Merge the common configuration
        Common.MergeGlobalConfig(commonConfigFileName, testSite);
        activeSyncClient = new ActiveSyncClient(testSite)
        {
            UserName = Common.GetConfigurationPropertyValue("UserName", testSite),
            Password = Common.GetConfigurationPropertyValue("UserPassword", testSite)
        };
    }

    /// <summary>
    /// Retrieves data from the server for one or more individual documents.
    /// </summary>
    /// <param name="itemOperationsRequest">ItemOperations command request.</param>
    /// <param name="deliverMethod">Deliver method parameter.</param>
    /// <returns>ItemOperations command response.</returns>
    public ItemOperationsResponse ItemOperations(ItemOperationsRequest itemOperationsRequest, DeliveryMethodForFetch deliverMethod)
    {
        var itemOperationsResponse = activeSyncClient.ItemOperations(itemOperationsRequest, deliverMethod);
        Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, itemOperationsResponse.StatusCode, "The call should be successful.");
        VerifyTransport();
        VerifyItemOperations(itemOperationsResponse);
        VerifyWBXMLCapture();
        return itemOperationsResponse;
    }

    /// <summary>
    /// Finds entries in document library (using Universal Naming Convention paths).
    /// </summary>
    /// <param name="searchRequest">Search command request.</param>
    /// <returns>Search command response.</returns>
    public SearchResponse Search(SearchRequest searchRequest)
    {
        var searchResponse = activeSyncClient.Search(searchRequest);
        Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, searchResponse.StatusCode, "The call should be successful.");
        VerifyTransport();
        VerifySearch(searchResponse);
        VerifyWBXMLCapture();
        return searchResponse;
    }
}