namespace Microsoft.Protocols.TestSuites.MS_ASCMD;

using System;
using System.Collections.Generic;
using System.Net;
using System.Threading;
using System.Xml.XPath;
using Common;
using TestTools;

/// <summary>
/// Adapter class of MS-ASCMD.
/// </summary>
public partial class MS_ASCMDAdapter : ManagedAdapterBase, IMS_ASCMDAdapter
{
    #region Variables
    /// <summary>
    /// The instance of ActiveSync client.
    /// </summary>
    private ActiveSyncClient activeSyncClient;
    #endregion

    #region IMS_ASCMDAdapter Properties
    /// <summary>
    /// Gets the raw XML request sent to protocol SUT
    /// </summary>
    public IXPathNavigable LastRawRequestXml
    {
        get { return activeSyncClient.LastRawRequestXml; }
    }

    /// <summary>
    /// Gets the raw XML response received from protocol SUT
    /// </summary>
    public IXPathNavigable LastRawResponseXml
    {
        get { return activeSyncClient.LastRawResponseXml; }
    }
    #endregion

    #region Initialize TestSuite
    /// <summary>
    /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
    /// </summary>
    /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
    public override void Initialize(ITestSite testSite)
    {
        base.Initialize(testSite);
        Site.DefaultProtocolDocShortName = "MS-ASCMD";

        // Merge the common configuration
        Common.MergeConfiguration(testSite);

        activeSyncClient = new ActiveSyncClient(testSite)
        {
            AcceptLanguage = "en-us",
            UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
            Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
        };
    }
    #endregion

    #region IMS-ASCMDAdapter Members
    /// <summary>
    /// Facilitates the discovery of core account configuration information by using the user's Simple Mail Transfer Protocol (SMTP) address as the primary input
    /// </summary>
    /// <param name="request">An AutodiscoverRequest object that contains the request information.</param>
    /// <param name="contentType">Content Type that indicates the body's format</param>
    /// <returns>Autodiscover command response</returns>
    public AutodiscoverResponse Autodiscover(AutodiscoverRequest request, ContentTypeEnum contentType)
    {
        var response = activeSyncClient.Autodiscover(request, contentType);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.Autodiscover, response);
        VerifyAutodiscoverCommand(response);
        return response;
    }

    /// <summary>
    /// Synchronizes changes in a collection between the client and the server.
    /// </summary>
    /// <param name="request">A SyncRequest object that contains the request information.</param>
    /// <param name="isResyncNeeded">A bool value indicate whether need to re-sync when the response contains MoreAvailable.</param>
    /// <returns>Sync command response</returns>
    public SyncResponse Sync(SyncRequest request, bool isResyncNeeded = true)
    {
        var response = activeSyncClient.Sync(request, isResyncNeeded);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.Sync, response);
        VerifySyncCommand(response);
        return response;
    }

    /// <summary>
    /// Sends MIME-formatted e-mail messages to the server.
    /// </summary>
    /// <param name="request">A SendMailRequest object that contains the request information.</param>
    /// <returns>SendMail command response</returns>
    public SendMailResponse SendMail(SendMailRequest request)
    {
        var response = activeSyncClient.SendMail(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.SendMail, response);
        VerifySendMailCommand(response);
        return response;
    }

    /// <summary>
    /// Retrieves an e-mail attachment from the server.
    /// </summary>
    /// <param name="request">A GetAttachmentRequest object that contains the request information.</param>
    /// <returns>GetAttachment command response</returns>
    public GetAttachmentResponse GetAttachment(GetAttachmentRequest request)
    {
        var response = activeSyncClient.GetAttachment(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.GetAttachment, response);
        return response;
    }

    /// <summary>
    /// Synchronizes the collection hierarchy
    /// </summary>
    /// <param name="request">A FolderSyncRequest object that contains the request information.</param>
    /// <returns>FolderSync command response</returns>
    public FolderSyncResponse FolderSync(FolderSyncRequest request)
    {
        var response = activeSyncClient.FolderSync(request);
        VerifyTransportRequirements();
        if (response.StatusCode == HttpStatusCode.OK)
        {
            VerifyWBXMLCapture(CommandName.FolderSync, response);
            VerifyFolderSyncCommand(response);
        }

        return response;
    }

    /// <summary>
    /// Creates a new folder as a child folder of the specified parent folder.
    /// </summary>
    /// <param name="request">A FolderCreateRequest object that contains the request information.</param>
    /// <returns>FolderCreate command response</returns>
    public FolderCreateResponse FolderCreate(FolderCreateRequest request)
    {
        var response = activeSyncClient.FolderCreate(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.FolderCreate, response);
        VerifyFolderCreateCommand(response);
        return response;
    }

    /// <summary>
    /// Deletes a folder from the server.
    /// </summary>
    /// <param name="request">A FolderDeleteRequest object that contains the request information.</param>
    /// <returns>FolderDelete command response</returns>
    public FolderDeleteResponse FolderDelete(FolderDeleteRequest request)
    {
        var response = activeSyncClient.FolderDelete(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.FolderDelete, response);
        VerifyFolderDeleteCommand(response);
        return response;
    }

    /// <summary>
    /// Moves a folder from one location to another on the server or renames a folder.
    /// </summary>
    /// <param name="request">A FolderUpdateRequest object that contains the request information.</param>
    /// <returns>FolderUpdate command response</returns>
    public FolderUpdateResponse FolderUpdate(FolderUpdateRequest request)
    {
        var response = activeSyncClient.FolderUpdate(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.FolderUpdate, response);
        VerifyFolderUpdateCommand(response);
        return response;
    }

    /// <summary>
    /// Moves an item or items from one folder to another on the server..
    /// </summary>
    /// <param name="request">A MoveItemsRequest object that contains the request information.</param>
    /// <returns>MoveItems command response</returns>
    public MoveItemsResponse MoveItems(MoveItemsRequest request)
    {
        var response = activeSyncClient.MoveItems(request);
        VerifyTransportRequirements();
        if (response.ResponseData.Response != null)
        {
            VerifyWBXMLCapture(CommandName.MoveItems, response);
            VerifyMoveItemsCommand(response);
        }

        return response;
    }

    /// <summary>
    /// Gets the list of email folders from the server
    /// </summary>
    /// <returns>GetHierarchy command response.</returns>
    public GetHierarchyResponse GetHierarchy()
    {
        var response = activeSyncClient.GetHierarchy();
        VerifyGetHierarchyCommand(response);
        return response;
    }

    /// <summary>
    /// Gets an estimated number of items in a collection or folder on the server that has to be synchronized.
    /// </summary>
    /// <param name="request">A GetItemEstimateRequest object that contains the request information.</param>
    /// <returns>GetItemEstimate command response</returns>
    public GetItemEstimateResponse GetItemEstimate(GetItemEstimateRequest request)
    {
        var response = activeSyncClient.GetItemEstimate(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.GetItemEstimate, response);
        VerifyGetItemEstimateCommand(response);
        return response;
    }

    /// <summary>
    /// Accepts, tentatively accepts, or declines a meeting request in the user's Inbox folder or Calendar folder.
    /// </summary>
    /// <param name="request">A MeetingResponseRequest object that contains the request information.</param>
    /// <returns>MeetingResponse command response</returns>
    public MeetingResponseResponse MeetingResponse(MeetingResponseRequest request)
    {
        var response = activeSyncClient.MeetingResponse(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.MeetingResponse, response);
        VerifyMeetingResponseCommand(response);
        return response;
    }

    /// <summary>
    /// Finds entries in an address book, mailbox, or document library.
    /// </summary>
    /// <param name="request">A SearchRequest object that contains the request information.</param>
    /// <returns>Search command response.</returns>
    public SearchResponse Search(SearchRequest request)
    {
        var counter = 0;
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        var retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var response = activeSyncClient.Search(request);

        while (counter < retryCount && response.ResponseData.Status.Equals("10"))
        {
            Thread.Sleep(waitTime);
            response = activeSyncClient.Search(request);
            counter++;
        }

        Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.Search, response);
        VerifySearchCommand(response);
        return response;
    }

    /// <summary>
    /// Finds entries in an address book, mailbox, or document library.
    /// </summary>
    /// <param name="request">A SearchRequest object that contains the request information.</param>
    /// <returns>Search command response.</returns>
    public FindResponse Find(FindRequest request)
    {
        var counter = 0;
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
        var retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
        var response = activeSyncClient.Find(request);

        while (counter < retryCount && response.ResponseData.Status.Equals("10"))
        {
            Thread.Sleep(waitTime);
            response = activeSyncClient.Find(request);
            counter++;
        }

        Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.Find, response);
        VerifyFindCommand(response);
        return response;
    }

    /// <summary>
    /// Supports get and set operations on global properties and Out of Office (OOF) settings for the user, sends device information to the server, implements the device password/personal identification number (PIN) recovery, and retrieves a list of the user's e-mail addresses.
    /// </summary>
    /// <param name="request">A SettingsRequest object that contains the request information.</param>
    /// <returns>Settings command response</returns>
    public SettingsResponse Settings(SettingsRequest request)
    {
        var response = activeSyncClient.Settings(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.Settings, response);
        VerifySettingsCommand(response);
        return response;
    }

    /// <summary>
    /// Forwards messages without retrieving the full, original message from the server.
    /// </summary>
    /// <param name="request">A SmartForwardRequest object that contains the request information.</param>
    /// <returns>SmartForward command response</returns>
    public SmartForwardResponse SmartForward(SmartForwardRequest request)
    {
        var response = activeSyncClient.SmartForward(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.SmartForward, response);
        VerifySmartForwardCommand(response);
        return response;
    }

    /// <summary>
    /// Reply to messages without retrieving the full, original message from the server.
    /// </summary>
    /// <param name="request">A SmartReplyRequest object that contains the request information.</param>
    /// <returns>SmartReply command response</returns>
    public SmartReplyResponse SmartReply(SmartReplyRequest request)
    {
        var response = activeSyncClient.SmartReply(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.SmartReply, response);
        VerifySmartReplyCommand(response);
        return response;
    }

    /// <summary>
    /// Requests that the server monitor specified folders for changes that would require the client to resynchronize.
    /// </summary>
    /// <param name="request">A PingRequest object that contains the request information.</param>
    /// <returns>Ping command response</returns>
    public PingResponse Ping(PingRequest request)
    {
        var response = activeSyncClient.Ping(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.Ping, response);
        VerifyPingCommand(response);
        return response;
    }

    /// <summary>
    /// Acts as a container for the Fetch element, the EmptyFolderContents element, and the Move element to provide batched online handling of these operations against the server.
    /// </summary>
    /// <param name="request">An ItemOperationsRequest object that contains the request information.</param>
    /// <param name="deliveryMethod">Delivery method specifies what kind of response is accepted.</param>
    /// <returns>ItemOperations command response</returns>
    public ItemOperationsResponse ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod)
    {
        var response = activeSyncClient.ItemOperations(request, deliveryMethod);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.ItemOperations, response);
        VerifyItemOperationsCommand(response);
        return response;
    }

    /// <summary>
    /// Enables client devices to request the administrator's security policy settings on the server..
    /// </summary>
    /// <param name="request">A ProvisionRequest object that contains the request information.</param>
    /// <returns>Provision command response</returns>
    public ProvisionResponse Provision(ProvisionRequest request)
    {
        // When the value of the MS-ASProtocolVersion header is 14.0 or 12.1, the client MUST NOT send the setting:DeviceInformation element in any Provision command request.
        if (activeSyncClient.ActiveSyncProtocolVersion.Equals("140", StringComparison.OrdinalIgnoreCase) || activeSyncClient.ActiveSyncProtocolVersion.Equals("121", StringComparison.OrdinalIgnoreCase))
        {
            request.RequestData.DeviceInformation = null;
        }

        var response = activeSyncClient.Provision(request);
        VerifyTransportRequirements();
        VerifyProvisionCommand(response);
        VerifyWBXMLCapture(CommandName.Provision, response);
        return response;
    }

    /// <summary>
    /// Resolves a list of supplied recipients, retrieves their free/busy information, or retrieves their S/MIME certificates so that clients can send encrypted S/MIME e-mail messages.
    /// </summary>
    /// <param name="request">A ResolveRecipientsRequest object that contains the request information.</param>
    /// <returns>ResolveRecipients command response</returns>
    public ResolveRecipientsResponse ResolveRecipients(ResolveRecipientsRequest request)
    {
        var response = activeSyncClient.ResolveRecipients(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.ResolveRecipients, response);
        VerifyResolveRecipientsCommand(response);
        return response;
    }

    /// <summary>
    /// Validates a certificate that has been received via an S/MIME mail.
    /// </summary>
    /// <param name="request">A ValidateCertRequest object that contains the request information.</param>
    /// <returns>ValidateCert command response</returns>
    public ValidateCertResponse ValidateCert(ValidateCertRequest request)
    {
        var response = activeSyncClient.ValidateCert(request);
        VerifyTransportRequirements();
        VerifyWBXMLCapture(CommandName.ValidateCert, response);
        VerifyValidateCertCommand(response);
        return response;
    }

    /// <summary>
    /// Sends a plain text request.
    /// </summary>
    /// <param name="cmdName">The name of the command to send</param>
    /// <param name="parameters">The command parameters</param>
    /// <param name="request">The plain text request</param>
    /// <returns>The plain text response.</returns>
    public SendStringResponse SendStringRequest(CommandName cmdName, IDictionary<CmdParameterName, object> parameters, string request)
    {
        var response = activeSyncClient.SendStringRequest(cmdName, parameters, request);
        return response;
    }

    /// <summary>
    /// Changes device id.
    /// </summary>
    /// <param name="newDeviceId">The new device id.</param>
    public void ChangeDeviceID(string newDeviceId)
    {
        activeSyncClient.DeviceID = newDeviceId;
    }

    /// <summary>
    /// Changes the specified PolicyKey.
    /// </summary>
    /// <param name="appliedPolicyKey">The Policy Key to apply.</param>
    public void ChangePolicyKey(string appliedPolicyKey)
    {
        activeSyncClient.PolicyKey = appliedPolicyKey;
    }

    /// <summary>
    /// Changes http request header encoding type
    /// </summary>
    /// <param name="headerEncodingType">The header encoding type</param>
    public void ChangeHeaderEncodingType(QueryValueType headerEncodingType)
    {
        activeSyncClient.QueryValueType = headerEncodingType;
    }

    /// <summary>
    /// Changes device type.
    /// </summary>
    /// <param name="newDeviceType">The value of the new device type.</param>
    public void ChangeDeviceType(string newDeviceType)
    {
        activeSyncClient.DeviceType = newDeviceType;
    }

    /// <summary>
    /// Changes user to call ActiveSync operation.
    /// </summary>
    /// <param name="userName">The user's name.</param>
    /// <param name="userPassword">The user's password.</param>
    /// <param name="userDomain">The domain which the user belongs to.</param>
    public void SwitchUser(string userName, string userPassword, string userDomain)
    {
        activeSyncClient.UserName = userName;
        activeSyncClient.Password = userPassword;
        activeSyncClient.Domain = userDomain;
    }
    #endregion
}
