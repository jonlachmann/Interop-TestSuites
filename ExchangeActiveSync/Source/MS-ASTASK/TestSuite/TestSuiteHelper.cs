namespace Microsoft.Protocols.TestSuites.MS_ASTASK;

using System;
using System.Collections.Generic;
using Common;
using Request = Common.Request;

/// <summary>
/// A static class contains all helper methods used in test cases.
/// </summary>
internal static class TestSuiteHelper
{
    /// <summary>
    /// Creates a Sync change request by using the specified sync key, folder collection ID and change application data.
    /// </summary>
    /// <param name="syncKey">Specify the sync key obtained from the last sync response.</param>
    /// <param name="collectionId">Specify the server ID of the folder to be synchronized.</param>
    /// <param name="data">Contains the data used to specify the Change element for Sync command.</param>
    /// <returns>Returns the SyncRequest instance.</returns>
    internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, List<object> data)
    {
        var syncCollection = new Request.SyncCollection
        {
            SyncKey = syncKey,
            GetChanges = true,
            GetChangesSpecified = true,
            DeletesAsMoves = false,
            DeletesAsMovesSpecified = true,
            CollectionId = collectionId
        };

        var option = new Request.Options();
        var preference = new Request.BodyPreference
        {
            Type = 2,
            Preview = 0,
            PreviewSpecified = false
        };

        option.Items = new object[] { preference };
        option.ItemsElementName = new Request.ItemsChoiceType1[]
        {
            Request.ItemsChoiceType1.BodyPreference,
        };

        syncCollection.Options = new Request.Options[1];
        syncCollection.Options[0] = option;

        syncCollection.WindowSize = "512";

        if (data != null)
        {
            syncCollection.Commands = data.ToArray();
        }

        return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
    }

    /// <summary>
    /// Builds an ItemOperations request to fetch the whole content of the tasks.
    /// </summary>
    /// <param name="collectionId">Specify the tasks folder.</param>
    /// <param name="serverIds">Specify a unique identifier that is assigned by the server for the tasks.</param>
    /// <param name="longIds">Specifies a unique identifier that was assigned by the server to each result returned by a previous Search response.</param>
    /// <param name="bodyPreference">Sets preference information related to the type and size of the body.</param>
    /// <param name="schema">Specifies the schema of the item to be fetched.</param>
    /// <returns>Returns the ItemOperationsRequest instance</returns>
    internal static ItemOperationsRequest CreateItemOperationsFetchRequest(
        string collectionId,
        List<string> serverIds,
        List<string> longIds,
        Request.BodyPreference bodyPreference,
        Request.Schema schema)
    {
        var fetchOptions = new Request.ItemOperationsFetchOptions();
        var fetchOptionItems = new List<object>();
        var fetchOptionItemsName = new List<Request.ItemsChoiceType5>();

        if (null != bodyPreference)
        {
            fetchOptionItemsName.Add(Request.ItemsChoiceType5.BodyPreference);
            fetchOptionItems.Add(bodyPreference);
        }

        if (null != schema)
        {
            fetchOptionItemsName.Add(Request.ItemsChoiceType5.Schema);
            fetchOptionItems.Add(schema);
        }

        fetchOptions.Items = fetchOptionItems.ToArray();
        fetchOptions.ItemsElementName = fetchOptionItemsName.ToArray();

        var fetchElements = new List<Request.ItemOperationsFetch>();

        if (serverIds != null)
        {
            foreach (var serverId in serverIds)
            {
                var fetchElement = new Request.ItemOperationsFetch()
                {
                    CollectionId = collectionId,
                    ServerId = serverId,
                    Store = SearchName.Mailbox.ToString(),
                    Options = fetchOptions
                };
                fetchElements.Add(fetchElement);
            }
        }

        if (longIds != null)
        {
            foreach (var longId in longIds)
            {
                var fetchElement = new Request.ItemOperationsFetch()
                {
                    LongId = longId,
                    Store = SearchName.Mailbox.ToString(),
                    Options = fetchOptions
                };
                fetchElements.Add(fetchElement);
            }
        }

        return Common.CreateItemOperationsRequest(fetchElements.ToArray());
    }

    /// <summary>
    /// Builds a Search request on the Mailbox store by using the specified keyword and folder collection ID
    /// In general, returns the XML formatted search request as follows:
    /// <!--
    /// <?xml version="1.0" encoding="utf-8"?>
    /// <Search xmlns="Search" xmlns:airsync="AirSync">
    /// <Store>
    ///   <Name>Mailbox</Name>
    ///     <Query>
    ///       <And>
    ///         <airsync:CollectionId>5</airsync:CollectionId>
    ///         <FreeText>Presentation</FreeText>
    ///       </And>
    ///     </Query>
    ///     <Options>
    ///       <RebuildResults />
    ///       <Range>0-9</Range>
    ///       <DeepTraversal/>
    ///     </Options>
    ///   </Store>
    /// </Search>
    /// -->
    /// </summary>
    /// <param name="storeName">Specify the store for which to search. Refer to [MS-ASCMD] section 2.2.3.110.2.</param>
    /// <param name="option">Specify a string value for which to search. Refer to [MS-ASCMD] section 2.2.3.73.</param>
    /// <param name="queryType">Specify the folder in which to search. Refer to [MS-ASCMD] section 2.2.3.30.4.</param>
    /// <returns>Returns a SearchRequest instance.</returns>
    internal static SearchRequest CreateSearchRequest(string storeName, Request.Options1 option, Request.queryType queryType)
    {
        var searchStore = new Request.SearchStore
        {
            Name = storeName,
            Options = option,
            Query = queryType
        };

        return Common.CreateSearchRequest(new Request.SearchStore[] { searchStore });
    }

    /// <summary>
    /// Create the elements of a task.
    /// </summary>
    /// <returns>The dictionary of value and name for task's elements to be created.</returns>
    internal static Dictionary<Request.ItemsChoiceType8, object> CreateTaskElements()
    {
        var addElements = new Dictionary<Request.ItemsChoiceType8, object>();

        var body = new Request.Body { Type = 1, Data = "Content of the body." + Guid.NewGuid().ToString() };
        addElements.Add(Request.ItemsChoiceType8.Body, body);

        var startTime = DateTime.Now.AddHours(1).AddDays(1);
        var utcStartTime = startTime.ToUniversalTime();

        addElements.Add(Request.ItemsChoiceType8.UtcStartDate, utcStartTime);
        addElements.Add(Request.ItemsChoiceType8.StartDate, startTime);

        addElements.Add(Request.ItemsChoiceType8.UtcDueDate, utcStartTime.AddHours(5));
        addElements.Add(Request.ItemsChoiceType8.DueDate, startTime.AddHours(5));

        addElements.Add(Request.ItemsChoiceType8.ReminderSet, byte.Parse("1"));
        addElements.Add(Request.ItemsChoiceType8.ReminderTime, startTime.AddDays(-1));
        return addElements;
    }
}