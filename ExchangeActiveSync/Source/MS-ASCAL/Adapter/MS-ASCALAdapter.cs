namespace Microsoft.Protocols.TestSuites.MS_ASCAL;

using System.Xml.XPath;
using Common;
using TestTools;
using ItemOperationsStore = Common.DataStructures.ItemOperationsStore;
using SearchStore = Common.DataStructures.SearchStore;
using SyncStore = Common.DataStructures.SyncStore;

/// <summary>
/// Adapter class of MS-ASCAL.
/// </summary>
public partial class MS_ASCALAdapter : ManagedAdapterBase, IMS_ASCALAdapter
{
    #region Private field

    /// <summary>
    /// The instance of ActiveSync client.
    /// </summary>
    private ActiveSyncClient activeSyncClient;

    /// <summary>
    ///  The calendar item server ID
    /// </summary>
    private string calendarId;

    #endregion

    #region IMS_ASCALAdapter Properties
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

    #region IMS_ASCALAdapter initialize method

    /// <summary>
    /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
    /// </summary>
    /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
    public override void Initialize(ITestSite testSite)
    {
        base.Initialize(testSite);
        testSite.DefaultProtocolDocShortName = "MS-ASCAL";

        // Merge the common configuration
        Common.MergeConfiguration(testSite);

        activeSyncClient = new ActiveSyncClient(testSite)
        {
            UserName = Common.GetConfigurationPropertyValue("OrganizerUserName", Site),
            Password = Common.GetConfigurationPropertyValue("OrganizerUserPassword", Site),
        };
    }

    #endregion

    #region IMS_ASCALAdapter Commands

    /// <summary>
    /// Sync calendars from the server
    /// </summary>
    /// <param name="syncRequest">The request for Sync command</param>
    /// <returns>The Sync response which is returned from server</returns>
    public SyncStore Sync(SyncRequest syncRequest)
    {
        var response = activeSyncClient.Sync(syncRequest, true);

        VerifyTransport();
        VerifyWBXMLRequirements();

        var syncResponse = Common.LoadSyncResponse(response);

        for (var i = syncRequest.RequestData.Collections.Length - 1; i >= 0; i--)
        {
            // Only verify the Sync response related calendar element
            if (syncRequest.RequestData.Collections[i].CollectionId == calendarId && syncResponse != null)
            {
                VerifyMessageSyntax();
                VerifySyncCommandResponse(syncResponse);
            }
        }

        return syncResponse;
    }

    /// <summary>
    /// FolderSync command to synchronize the collection hierarchy
    /// </summary>
    /// <returns>The FolderSync response</returns>
    public FolderSyncResponse FolderSync()
    {
        var request = Common.CreateFolderSyncRequest("0");
        var response = activeSyncClient.FolderSync(request);

        calendarId = Common.GetDefaultFolderServerId(response, FolderType.Calendar, Site);

        return response;
    }

    /// <summary>
    /// Search calendars using the given keyword text
    /// </summary>
    /// <param name="searchRequest">The request for Search command</param>
    /// <returns>The search data returned from the server</returns>
    public SearchStore Search(SearchRequest searchRequest)
    {
        var response = activeSyncClient.Search(searchRequest);
        var searchResponse = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site));

        VerifyTransport();
        VerifySearchCommandResponse(searchResponse);

        return searchResponse;
    }

    /// <summary>
    /// Fetch all the information about calendars using longIds or ServerIds
    /// </summary>
    /// <param name="itemOperationsRequest">The request for ItemOperations</param>
    /// <returns>The fetch items information</returns>
    public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
    {
        var response = activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
        var itemOperationResponse = Common.LoadItemOperationsResponse(response);

        Site.Assert.IsNotNull(itemOperationResponse, "ItemOperation command response should be returned successfully.");
        VerifyTransport();
        VerifyItemOperationsResponse(itemOperationResponse);

        return itemOperationResponse;
    }

    /// <summary>
    /// MeetingResponse for accepting or declining a meeting request
    /// </summary>
    /// <param name="meetingResponseRequest">The request for MeetingResponse</param>
    /// <returns>The MeetingResponse response which is returned from server</returns>
    public MeetingResponseResponse MeetingResponse(MeetingResponseRequest meetingResponseRequest)
    {
        var response = activeSyncClient.MeetingResponse(meetingResponseRequest);
        Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
        return response;
    }

    /// <summary>
    /// Sends MIME-formatted e-mail messages to the server.
    /// </summary>
    /// <param name="sendMailRequest">The request for SendMail command</param>
    /// <returns>The SendMail response which is returned from the server</returns>
    public SendMailResponse SendMail(SendMailRequest sendMailRequest)
    {
        var response = activeSyncClient.SendMail(sendMailRequest);
        Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

        return response;
    }

    /// <summary>
    /// Send a Sync command string request and get Sync response from server.
    /// </summary>
    /// <param name="stringRequest">The request for Sync command</param>
    /// <returns>The Sync response which is returned from server</returns>
    public SendStringResponse SendStringRequest(string stringRequest)
    {
        var response = activeSyncClient.SendStringRequest(CommandName.Sync, null, stringRequest);
        return response;
    }

    /// <summary>
    /// Change user to call ActiveSync command.
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

    #endregion
}
