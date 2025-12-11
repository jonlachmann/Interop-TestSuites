namespace Microsoft.Protocols.TestSuites.MS_ASNOTE;

using Common;
using Request = Common.Request;

/// <summary>
/// The class provides the methods to assist MS_ASNOTEAdapter.
/// </summary>
public static class AdapterHelper
{
    #region Adapter Help Methods
    /// <summary>
    /// Check if the request of ItemOperations command contains the schema element
    /// </summary>
    /// <param name="itemOperationsRequest">ItemOperations command request.</param>
    /// <returns>The boolean value represents whether the request of ItemOperations command contains the schema element</returns>
    public static bool ContainsSchemaElement(ActiveSyncRequestBase<Request.ItemOperations> itemOperationsRequest)
    {
        var fetch = (Request.ItemOperationsFetch)itemOperationsRequest.RequestData.Items[0];

        var hasSchemaElement = false;

        // Check if the request contains schema
        foreach (var item in fetch.Options.Items)
        {
            if (item.GetType().Equals(typeof(Request.Schema)))
            {
                hasSchemaElement = true;
            }
        }

        return hasSchemaElement;
    }
    #endregion
}