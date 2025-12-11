namespace Microsoft.Protocols.TestSuites.Common;

#region AutodiscoverRequest

/// <summary>
/// The class of Autodiscover request.
/// </summary>
public class AutodiscoverRequest : ActiveSyncRequestBase<Request.Autodiscover>
{
    /// <summary>
    /// Initializes a new instance of the AutodiscoverRequest class.
    /// </summary>
    public AutodiscoverRequest()
    {
        RequestData = new Request.Autodiscover { Request = new Request.RequestType() };
    }
}
#endregion

#region SynRequest

/// <summary>
/// The class of Sync request.
/// </summary>
public class SyncRequest : ActiveSyncRequestBase<Request.Sync>
{
    /// <summary>
    /// Initializes a new instance of the SyncRequest class.
    /// </summary>
    public SyncRequest()
    {
        RequestData = new Request.Sync();
    }
}
#endregion

#region FindRequest
/// <summary>
/// The class of Find request.
/// </summary>
public class FindRequest : ActiveSyncRequestBase<Request.Find>
{
    /// <summary>
    /// Initializes a new instance of the FindRequest class.
    /// </summary>
    public FindRequest()
    {
        RequestData = new Request.Find();
    }
}
#endregion

#region FolderSyncRequest
/// <summary>
/// The class of FolderSync request.
/// </summary>
public class FolderSyncRequest : ActiveSyncRequestBase<Request.FolderSync>
{
    /// <summary>
    /// Initializes a new instance of the FolderSyncRequest class.
    /// </summary>
    public FolderSyncRequest()
    {
        RequestData = new Request.FolderSync();
    }
}
#endregion

#region FolderCreateRequest
/// <summary>
/// The class of FolderCreate request.
/// </summary>
public class FolderCreateRequest : ActiveSyncRequestBase<Request.FolderCreate>
{
    /// <summary>
    /// Initializes a new instance of the FolderCreateRequest class.
    /// </summary>
    public FolderCreateRequest()
    {
        RequestData = new Request.FolderCreate();
    }
}
#endregion

#region FolderDeleteRequest
/// <summary>
/// The class of FolderDelete request.
/// </summary>
public class FolderDeleteRequest : ActiveSyncRequestBase<Request.FolderDelete>
{
    /// <summary>
    /// Initializes a new instance of the FolderDeleteRequest class.
    /// </summary>
    public FolderDeleteRequest()
    {
        RequestData = new Request.FolderDelete();
    }
}
#endregion

#region FolderUpdateRequest
/// <summary>
/// The class of FolderUpdate request.
/// </summary>
public class FolderUpdateRequest : ActiveSyncRequestBase<Request.FolderUpdate>
{
    /// <summary>
    /// Initializes a new instance of the FolderUpdateRequest class.
    /// </summary>
    public FolderUpdateRequest()
    {
        RequestData = new Request.FolderUpdate();
    }
}
#endregion

#region GetAttachmentRequest
/// <summary>
/// The class of GetAttachment request.
/// </summary>
public class GetAttachmentRequest : ActiveSyncRequestBase<object>
{
    /// <summary>
    /// Initializes a new instance of the GetAttachmentRequest class.
    /// </summary>
    public GetAttachmentRequest()
    {
        RequestData = null;
    }
}
#endregion

#region GetItemEstimateRequest
/// <summary>
/// The class of GetItemEstimate request.
/// </summary>
public class GetItemEstimateRequest : ActiveSyncRequestBase<Request.GetItemEstimate>
{
    /// <summary>
    /// Initializes a new instance of the GetItemEstimateRequest class.
    /// </summary>
    public GetItemEstimateRequest()
    {
        RequestData = new Request.GetItemEstimate();
    }
}
#endregion

#region ItemOperationsRequest
/// <summary>
/// The class of ItemOperations request.
/// </summary>
public class ItemOperationsRequest : ActiveSyncRequestBase<Request.ItemOperations>
{
    /// <summary>
    /// Initializes a new instance of the ItemOperationsRequest class.
    /// </summary>
    public ItemOperationsRequest()
    {
        RequestData = new Request.ItemOperations();
    }
}
#endregion

#region MeetingResponseRequest
/// <summary>
/// The class of MeetingResponse request.
/// </summary>
public class MeetingResponseRequest : ActiveSyncRequestBase<Request.MeetingResponse>
{
    /// <summary>
    /// Initializes a new instance of the MeetingResponseRequest class.
    /// </summary>
    public MeetingResponseRequest()
    {
        RequestData = new Request.MeetingResponse();
    }
}
#endregion

#region MoveItemsRequest
/// <summary>
/// The class of MoveItems request.
/// </summary>
public class MoveItemsRequest : ActiveSyncRequestBase<Request.MoveItems>
{
    /// <summary>
    /// Initializes a new instance of the MoveItemsRequest class.
    /// </summary>
    public MoveItemsRequest()
    {
        RequestData = new Request.MoveItems();
    }
}
#endregion

#region PingRequest
/// <summary>
/// The class of Ping request.
/// </summary>
public class PingRequest : ActiveSyncRequestBase<Request.Ping>
{
    /// <summary>
    /// Initializes a new instance of the PingRequest class.
    /// </summary>
    public PingRequest()
    {
        RequestData = new Request.Ping();
    }
}
#endregion

#region ProvisionRequest
/// <summary>
/// The class of Provision request.
/// </summary>
public class ProvisionRequest : ActiveSyncRequestBase<Request.Provision>
{
    /// <summary>
    /// Initializes a new instance of the ProvisionRequest class.
    /// </summary>
    public ProvisionRequest()
    {
        RequestData = new Request.Provision();
    }
}
#endregion

#region ResolveRecipientsRequest
/// <summary>
/// The class of ResolveRecipients request.
/// </summary>
public class ResolveRecipientsRequest : ActiveSyncRequestBase<Request.ResolveRecipients>
{
    /// <summary>
    /// Initializes a new instance of the ResolveRecipientsRequest class.
    /// </summary>
    public ResolveRecipientsRequest()
    {
        RequestData = new Request.ResolveRecipients();
    }
}
#endregion

#region SearchRequest
/// <summary>
/// The class of Search request.
/// </summary>
public class SearchRequest : ActiveSyncRequestBase<Request.Search>
{
    /// <summary>
    /// Initializes a new instance of the SearchRequest class.
    /// </summary>
    public SearchRequest()
    {
        RequestData = new Request.Search();
    }
}
#endregion

#region SendMailRequest
/// <summary>
/// The class of SendMail request.
/// </summary>
public class SendMailRequest : ActiveSyncRequestBase<Request.SendMail>
{
    /// <summary>
    /// Initializes a new instance of the SendMailRequest class.
    /// </summary>
    public SendMailRequest()
    {
        RequestData = new Request.SendMail();
    }
}
#endregion

#region SettingsRequest
/// <summary>
/// The class of Settings request.
/// </summary>
public class SettingsRequest : ActiveSyncRequestBase<Request.Settings>
{
    /// <summary>
    /// Initializes a new instance of the SettingsRequest class.
    /// </summary>
    public SettingsRequest()
    {
        RequestData = new Request.Settings();
    }
}
#endregion

#region SmartForwardRequest
/// <summary>
/// The class of SmartForward request.
/// </summary>
public class SmartForwardRequest : ActiveSyncRequestBase<Request.SmartForward>
{
    /// <summary>
    /// Initializes a new instance of the SmartForwardRequest class.
    /// </summary>
    public SmartForwardRequest()
    {
        RequestData = new Request.SmartForward();
    }
}
#endregion

#region SmartReplyRequest
/// <summary>
/// The class of SmartReply request.
/// </summary>
public class SmartReplyRequest : ActiveSyncRequestBase<Request.SmartReply>
{
    /// <summary>
    /// Initializes a new instance of the SmartReplyRequest class.
    /// </summary>
    public SmartReplyRequest()
    {
        RequestData = new Request.SmartReply();
    }
}
#endregion

#region ValidateCertRequest
/// <summary>
/// The class of ValidateCert request.
/// </summary>
public class ValidateCertRequest : ActiveSyncRequestBase<Request.ValidateCert>
{
    /// <summary>
    /// Initializes a new instance of the ValidateCertRequest class.
    /// </summary>
    public ValidateCertRequest()
    {
        RequestData = new Request.ValidateCert();
    }
}
#endregion