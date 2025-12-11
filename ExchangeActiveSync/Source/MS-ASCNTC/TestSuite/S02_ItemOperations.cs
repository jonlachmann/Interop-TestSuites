namespace Microsoft.Protocols.TestSuites.MS_ASCNTC;

using System.Collections.Generic;
using System.Xml;
using Common.DataStructures;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Request = Common.Request;

/// <summary>
/// This scenario is designed to use ItemOperations command to retrieve Contact class data from the server.
/// </summary>
[TestClass]
public class S02_ItemOperations : TestSuiteBase
{
    #region Class initialize and clean up
    /// <summary>
    /// Initialize the class.
    /// </summary>
    /// <param name="testContext">VSTS test context.</param>
    [ClassInitialize]
    public static void ClassInitialize(TestContext testContext)
    {
        Initialize(testContext);
    }

    /// <summary>
    /// Clear the class.
    /// </summary>
    [ClassCleanup]
    public static void ClassCleanup()
    {
        Cleanup();
    }
    #endregion

    #region MSASCNTC_S02_TC01_ItemOperations_RetrieveContact
    /// <summary>
    /// This case is designed to retrieve contact item using ItemOperations command.
    /// </summary>
    [TestCategory("MSASCNTC"), TestMethod]
    public void MSASCNTC_S02_TC01_ItemOperations_RetrieveContact()
    {
        #region Call Sync command with Add element to add a contact with all Contact class elements to the server
        var picture = "SmallPhoto.jpg";

        var contactProperties = SetContactProperties(picture);
        AddContact(User1Information.ContactsCollectionId, contactProperties);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());

        // Get the new added contact
        var newAddedItem = GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), User1Information.ContactsCollectionId, null, null);
        #endregion

        #region Call ItemOperations command to retrieve the contact item that added in previous step
        var bodyPreference = new Request.BodyPreference { Type = 1 };

        var itemOperationsItem = GetItemOperationsResult(User1Information.ContactsCollectionId, newAddedItem.ServerId, bodyPreference, null);
        #endregion

        #region Verify requirements
        VerifyContactClassElements(contactProperties, itemOperationsItem.Contact);
        #endregion
    }
    #endregion

    #region MSASCNTC_S02_TC02_ItemOperations_TruncateBody
    /// <summary>
    /// This case is designed to test server truncates the contents of the airsyncbase:Body element in the ItemOperations command response if the client requests truncation.
    /// </summary>
    [TestCategory("MSASCNTC"), TestMethod]
    public void MSASCNTC_S02_TC02_ItemOperations_TruncateBody()
    {
        #region Call Sync command with Add element to add a contact with all Contact class elements to the server
        var picture = "SmallPhoto.jpg";

        var contactProperties = SetContactProperties(picture);
        AddContact(User1Information.ContactsCollectionId, contactProperties);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());

        // Get the new added contact
        var newAddedItem = GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), User1Information.ContactsCollectionId, null, null);
        #endregion

        #region Call ItemOperations command with TruncationSize element smaller than the available data of the body
        var bodyPreference = new Request.BodyPreference
        {
            Type = 1,
            TruncationSize = 8,
            TruncationSizeSpecified = true
        };

        var item = GetItemOperationsResult(User1Information.ContactsCollectionId, newAddedItem.ServerId, bodyPreference, null);

        // Assert the body data is truncated.
        Site.Assert.AreEqual<string>(
            ((Request.Body)contactProperties[Request.ItemsChoiceType8.Body]).Data.Substring(0, (int)bodyPreference.TruncationSize),
            item.Contact.Body.Data,
            "The body data should be truncated when the value of TruncationSize element is smaller than the available data size.");
        #endregion
    }
    #endregion

    #region MSASCNTC_S02_TC03_ItemOperations_SchemaViewFetch
    /// <summary>
    /// This case is designed to test if an airsync:Schema element is included in the ItemOperations command request; server response must be restricted to the elements that were included as child elements of the airsync:Schema element in the command request.
    /// </summary>
    [TestCategory("MSASCNTC"), TestMethod]
    public void MSASCNTC_S02_TC03_ItemOperations_SchemaViewFetch()
    {
        #region Call Sync command with Add element to add a contact to the server
        var contactProperties = SetContactProperties(null);
        AddContact(User1Information.ContactsCollectionId, contactProperties);

        RecordCaseRelativeItems(User1Information.UserName, User1Information.ContactsCollectionId, contactProperties[Request.ItemsChoiceType8.FileAs].ToString());

        // Get the new added contact
        var newAddedItem = GetSyncAddResult(contactProperties[Request.ItemsChoiceType8.FileAs].ToString(), User1Information.ContactsCollectionId, null, null);
        #endregion

        #region Call ItemOperations command with Schema element to retrieve the contact item added in previous step
        // Just including FileAs element in Schema
        var schema = new Request.Schema
        {
            Items = [string.Empty],
            ItemsElementName = [Request.ItemsChoiceType4.FileAs]
        };

        GetItemOperationsResult(User1Information.ContactsCollectionId, newAddedItem.ServerId, null, schema);
        #endregion

        #region Verify requirement
        // If only FilsAs element is returned in server response, then capture MS-ASCNTC_R485.
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R485.");

        // Verify MS-ASCNTC requirement: MS-ASCNTC_R485
        Site.CaptureRequirementIfIsTrue(
            TestSuiteHelper.IsOnlySpecifiedElementExist((XmlElement)ASCNTCAdapter.LastRawResponseXml, "Properties", "FileAs"),
            485,
            @"[In ItemOperations Command Response] If an ItemOperations:Schema element ([MS-ASCMD] section 2.2.3.145) is included in the ItemOperations command request, the elements returned in the ItemOperations command response MUST be restricted to the elements that were included as child elements of the itemoperations:Schema element in the command request.");
        #endregion
    }
    #endregion
}