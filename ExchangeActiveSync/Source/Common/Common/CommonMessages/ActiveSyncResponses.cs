namespace Microsoft.Protocols.TestSuites.Common;

#region AutodiscoverResponse

/// <summary>
/// The class of Autodiscover response.
/// </summary>
public class AutodiscoverResponse : ActiveSyncResponseBase<Response.Autodiscover>
{
    /// <summary>
    /// Initializes a new instance of the AutodiscoverResponse class.
    /// </summary>
    public AutodiscoverResponse()
    {
        ResponseData = new Response.Autodiscover();
    }
}
#endregion

#region SyncResponse
/// <summary>
/// The class of Sync response.
/// </summary>
public class SyncResponse : ActiveSyncResponseBase<Response.Sync>
{
    /// <summary>
    /// Initializes a new instance of the SyncResponse class.
    /// </summary>
    public SyncResponse()
    {
        ResponseData = new Response.Sync();
    }
}
#endregion

#region Find
/// <summary>
/// The class of Find response.
/// </summary>
public class FindResponse : ActiveSyncResponseBase<Response.Find>
{
    /// <summary>
    /// Initializes a new instance of the FindResponse class.
    /// </summary>
    public FindResponse()
    {
        ResponseData = new Response.Find();
    }
}
#endregion

#region FolderSyncResponse
/// <summary>
/// The class of FolderSync response.
/// </summary>
public class FolderSyncResponse : ActiveSyncResponseBase<Response.FolderSync>
{
    /// <summary>
    /// Initializes a new instance of the FolderSyncResponse class.
    /// </summary>
    public FolderSyncResponse()
    {
        ResponseData = new Response.FolderSync();
    }
}
#endregion

#region FolderCreateResponse
/// <summary>
/// The class of FolderCreate response.
/// </summary>
public class FolderCreateResponse : ActiveSyncResponseBase<Response.FolderCreate>
{
    /// <summary>
    /// Initializes a new instance of the FolderCreateResponse class.
    /// </summary>
    public FolderCreateResponse()
    {
        ResponseData = new Response.FolderCreate();
    }
}
#endregion

#region FolderDeleteResponse
/// <summary>
/// The class of FolderDelete response.
/// </summary>
public class FolderDeleteResponse : ActiveSyncResponseBase<Response.FolderDelete>
{
    /// <summary>
    /// Initializes a new instance of the FolderDeleteResponse class.
    /// </summary>
    public FolderDeleteResponse()
    {
        ResponseData = new Response.FolderDelete();
    }
}
#endregion

#region FolderUpdateResponse
/// <summary>
/// The class of FolderUpdate response.
/// </summary>
public class FolderUpdateResponse : ActiveSyncResponseBase<Response.FolderUpdate>
{
    /// <summary>
    /// Initializes a new instance of the FolderUpdateResponse class.
    /// </summary>
    public FolderUpdateResponse()
    {
        ResponseData = new Response.FolderUpdate();
    }
}
#endregion

#region GetAttachmentResponse
/// <summary>
/// The class of GetAttachment response.
/// </summary>
public class GetAttachmentResponse : ActiveSyncResponseBase<object>
{
    /// <summary>
    /// Initializes a new instance of the GetAttachmentResponse class.
    /// </summary>
    public GetAttachmentResponse()
    {
        ResponseData = null;
    }
}
#endregion

#region GetHierarchyResponse
/// <summary>
/// The class of  GetHierarchy response.
/// </summary>
public class GetHierarchyResponse : ActiveSyncResponseBase<Response.Folders>
{
    /// <summary>
    ///  Initializes a new instance of the GetHierarchyResponse class.
    /// </summary>
    public GetHierarchyResponse()
    {
        ResponseData = new Response.Folders();
    }
}
#endregion

#region GetItemEstimateResponse
/// <summary>
/// The class of GetItemEstimate response.
/// </summary>
public class GetItemEstimateResponse : ActiveSyncResponseBase<Response.GetItemEstimate>
{
    /// <summary>
    /// Initializes a new instance of the GetItemEstimateResponse class.
    /// </summary>
    public GetItemEstimateResponse()
    {
        ResponseData = new Response.GetItemEstimate();
    }
}
#endregion

#region ItemOperationsResponse
/// <summary>
/// The class of ItemOperations response.
/// </summary>
public class ItemOperationsResponse : ActiveSyncResponseBase<Response.ItemOperations>
{
    /// <summary>
    /// The MultiPart response.
    /// </summary>
    private MultipartMetadata metadata;

    /// <summary>
    /// Initializes a new instance of the ItemOperationsResponse class.
    /// </summary>
    public ItemOperationsResponse()
    {
        ResponseData = new Response.ItemOperations();
        metadata = null;
    }

    /// <summary>
    /// Gets the MultiPart response.
    /// </summary>
    public MultipartMetadata MultipartMetadata
    {
        get
        {
            return metadata;
        }
    }
}
#endregion

#region MeetingResponseResponse
/// <summary>
/// The class of MeetingResponse response.
/// </summary>
public class MeetingResponseResponse : ActiveSyncResponseBase<Response.MeetingResponse>
{
    /// <summary>
    /// Initializes a new instance of the MeetingResponseResponse class.
    /// </summary>
    public MeetingResponseResponse()
    {
        ResponseData = new Response.MeetingResponse();
    }
}
#endregion

#region MoveItemsResponse
/// <summary>
/// The class of MoveItems response.
/// </summary>
public class MoveItemsResponse : ActiveSyncResponseBase<Response.MoveItems>
{
    /// <summary>
    /// Initializes a new instance of the MoveItemsResponse class.
    /// </summary>
    public MoveItemsResponse()
    {
        ResponseData = new Response.MoveItems();
    }
}
#endregion

#region PingResponse
/// <summary>
/// The class of Ping response.
/// </summary>
public class PingResponse : ActiveSyncResponseBase<Response.Ping>
{
    /// <summary>
    /// Initializes a new instance of the PingResponse class.
    /// </summary>
    public PingResponse()
    {
        ResponseData = new Response.Ping();
    }
}
#endregion

#region ProvisionResponse
/// <summary>
/// The class of Provision response.
/// </summary>
public class ProvisionResponse : ActiveSyncResponseBase<Response.Provision>
{
    /// <summary>
    /// Initializes a new instance of the ProvisionResponse class.
    /// </summary>
    public ProvisionResponse()
    {
        ResponseData = new Response.Provision();
    }
}
#endregion

#region ResolveRecipientsResponse
/// <summary>
/// The class of ResolveRecipient response.
/// </summary>
public class ResolveRecipientsResponse : ActiveSyncResponseBase<Response.ResolveRecipients>
{
    /// <summary>
    /// Initializes a new instance of the ResolveRecipientsResponse class.
    /// </summary>
    public ResolveRecipientsResponse()
    {
        ResponseData = new Response.ResolveRecipients();
    }
}
#endregion

#region SearchResponse
/// <summary>
/// The class of Search response.
/// </summary>
public class SearchResponse : ActiveSyncResponseBase<Response.Search>
{
    /// <summary>
    /// Initializes a new instance of the SearchResponse class.
    /// </summary>
    public SearchResponse()
    {
        ResponseData = new Response.Search();
    }
}
#endregion

#region SendMailResponse
/// <summary>
/// The class of SendMail response.
/// </summary>
public class SendMailResponse : ActiveSyncResponseBase<Response.SendMail>
{
    /// <summary>
    /// Initializes a new instance of the SendMailResponse class.
    /// </summary>
    public SendMailResponse()
    {
        ResponseData = new Response.SendMail();
    }
}
#endregion

#region SettingsResponse
/// <summary>
/// The class of Settings response.
/// </summary>
public class SettingsResponse : ActiveSyncResponseBase<Response.Settings>
{
    /// <summary>
    /// Initializes a new instance of the SettingsResponse class.
    /// </summary>
    public SettingsResponse()
    {
        ResponseData = new Response.Settings();
    }
}
#endregion

#region SmartForwardResponse
/// <summary>
/// The class of SmartForward response.
/// </summary>
public class SmartForwardResponse : ActiveSyncResponseBase<Response.SmartForward>
{
    /// <summary>
    /// Initializes a new instance of the SmartForwardResponse class.
    /// </summary>
    public SmartForwardResponse()
    {
        ResponseData = new Response.SmartForward();
    }
}
#endregion

#region SmartReplyResponse
/// <summary>
/// The class of SmartReply response.
/// </summary>
public class SmartReplyResponse : ActiveSyncResponseBase<Response.SmartReply>
{
    /// <summary>
    /// Initializes a new instance of the SmartReplyResponse class.
    /// </summary>
    public SmartReplyResponse()
    {
        ResponseData = new Response.SmartReply();
    }
}
#endregion

#region ValidateCertResponse
/// <summary>
/// The class of ValidateCert response.
/// </summary>
public class ValidateCertResponse : ActiveSyncResponseBase<Response.ValidateCert>
{
    /// <summary>
    /// Initializes a new instance of the ValidateCertResponse class.
    /// </summary>
    public ValidateCertResponse()
    {
        ResponseData = new Response.ValidateCert();
    }
}
#endregion

#region OptionsResponse
/// <summary>
/// The class of Options response.
/// </summary>
public class OptionsResponse : ActiveSyncResponseBase<object>
{
    /// <summary>
    /// Initializes a new instance of the OptionsResponse class.
    /// </summary>
    public OptionsResponse()
    {
        ResponseData = null;
    }
}
#endregion

#region SendStringResponse
/// <summary>
/// The class of SendString response.
/// </summary>
public class SendStringResponse : ActiveSyncResponseBase<object>
{
    /// <summary>
    /// Initializes a new instance of the SendStringResponse class.
    /// </summary>
    public SendStringResponse()
    {
        ResponseData = null;
    }
}
#endregion
