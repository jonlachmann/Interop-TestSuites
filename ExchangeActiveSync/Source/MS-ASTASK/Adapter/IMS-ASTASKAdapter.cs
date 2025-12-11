namespace Microsoft.Protocols.TestSuites.MS_ASTASK;

using System.Xml.XPath;
using Common;
using TestTools;
using ItemOperationsStore = Common.DataStructures.ItemOperationsStore;
using SearchStore = Common.DataStructures.SearchStore;
using SyncStore = Common.DataStructures.SyncStore;

/// <summary>
/// The adapter interface which provides methods defined in MS-ASTASK.
/// </summary>
public interface IMS_ASTASKAdapter : IAdapter
{
    /// <summary>
    /// Gets the raw XML request sent to protocol SUT.
    /// </summary>
    IXPathNavigable LastRawRequestXml { get; }

    /// <summary>
    /// Gets the raw XML response received from protocol SUT.
    /// </summary>
    IXPathNavigable LastRawResponseXml { get; }

    /// <summary>
    /// Sync data from the server.
    /// </summary>
    /// <param name="syncRequest">A Sync command request.</param>
    /// <returns>A Sync command response returned from the server.</returns>
    SyncStore Sync(SyncRequest syncRequest);

    /// <summary>
    /// Search data using the given keyword text.
    /// </summary>
    /// <param name="searchRequest">A Search command request.</param>
    /// <returns>A Search command response returned from the server.</returns>
    SearchStore Search(SearchRequest searchRequest);

    /// <summary>
    /// Fetch all the data about tasks.
    /// </summary>
    /// <param name="itemOperationsRequest">An ItemOperations command request.</param>
    /// <returns>An ItemOperations command response returned from the server.</returns>
    ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest);

    /// <summary>
    /// Synchronize the collection hierarchy.
    /// </summary>
    /// <returns>A FolderSync command response returned from the server.</returns>
    FolderSyncResponse FolderSync();

    /// <summary>
    /// Send a string request and get a response from server.
    /// </summary>
    /// <param name="stringRequest">A string request for a certain command.</param>
    /// <param name="commandName">Commands choices.</param>
    /// <returns>A string response returned from the server.</returns>
    SendStringResponse SendStringRequest(string stringRequest, CommandName commandName);
}