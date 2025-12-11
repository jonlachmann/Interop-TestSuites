namespace Microsoft.Protocols.TestSuites.MS_ASPROV;

using System.Collections.Generic;
using System.Xml;
using Common;
using Response = Common.Response;

/// <summary>
/// The class provides the methods to assist MS_ASPROVAdapter.
/// </summary>
public static class AdapterHelper
{
    /// <summary>
    /// Get policies from Provision command response.
    /// </summary>
    /// <param name="provisionResponse">The response of Provision command.</param>
    /// <returns>The dictionary of policies gotten from Provision command response.</returns>
    public static Dictionary<string, string> GetPoliciesFromProvisionResponse(ActiveSyncResponseBase<Response.Provision> provisionResponse)
    {
        var policiesSetting = new Dictionary<string, string>();
        if (null == provisionResponse || string.IsNullOrEmpty(provisionResponse.ResponseDataXML))
        {
            return policiesSetting;
        }

        var xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(provisionResponse.ResponseDataXML);
        var namespaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
        namespaceManager.AddNamespace("prov", "Provision");
        var provisionDocNode = xmlDoc.SelectSingleNode(@"//prov:EASProvisionDoc", namespaceManager);

        if (provisionDocNode != null && provisionDocNode.HasChildNodes)
        {
            foreach (XmlNode policySetting in provisionDocNode.ChildNodes)
            {
                var policyValue = string.IsNullOrEmpty(policySetting.InnerText) ? string.Empty : policySetting.InnerText;
                var policyName = string.IsNullOrEmpty(policySetting.LocalName) ? string.Empty : policySetting.LocalName;
                policiesSetting.Add(policyName, policyValue);
            }
        }

        return policiesSetting;
    }
}