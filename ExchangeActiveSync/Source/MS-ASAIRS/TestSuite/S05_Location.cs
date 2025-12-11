namespace Microsoft.Protocols.TestSuites.MS_ASAIRS;

using System.Collections.Generic;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataStructures = Common.DataStructures;
using Request = Common.Request;

/// <summary>
/// This scenario is designed to test the Location element and its sub elements, which is used by the Sync command, Search command and ItemOperations command.
/// </summary>
[TestClass]
public class S05_Location : TestSuiteBase
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
    public static void ClassCleanUp()
    {
        Cleanup();
    }
    #endregion

    #region MSASAIRS_S05_TC01_Location
    /// <summary>
    /// This case is designed to test element Location and its sub elements.
    /// </summary>
    [TestCategory("MSASAIRS"), TestMethod]
    public void MSASAIRS_S05_TC01_Location()
    {
        Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The Location element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The Location element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The Location element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

        #region Call Sync command with Add element to add an appointment to the server
        var applicationData = new Request.SyncCollectionAddApplicationData();

        var items = new List<object>();
        var itemsElementName = new List<Request.ItemsChoiceType8>();

        var subject = Common.GenerateResourceName(Site, "Subject");
        items.Add(subject);
        itemsElementName.Add(Request.ItemsChoiceType8.Subject);

        // MeetingStauts is set to 0, which means it is an appointment with no attendees.
        byte meetingStatus = 0;
        items.Add(meetingStatus);
        itemsElementName.Add(Request.ItemsChoiceType8.MeetingStatus);

        var location = new Request.Location();
        location.Accuracy = (double)1;
        location.AccuracySpecified = true;
        location.Altitude = (double)55.46;
        location.AltitudeAccuracy = (double)1;
        location.AltitudeAccuracySpecified = true;
        location.AltitudeSpecified = true;
        location.Annotation = "Location sample annotation";
        location.City = "Location sample city";
        location.Country = "Location sample country";
        location.DisplayName = "Location sample dislay name";
        location.Latitude = (double)11.56;
        location.LatitudeSpecified = true;
        location.LocationUri = "Location Uri";
        location.Longitude = (double)1.9;
        location.LongitudeSpecified = true;
        location.PostalCode = "Location sample postal code";
        location.State = "Location sample state";
        location.Street = "Location sample street";
        items.Add(location);
        itemsElementName.Add(Request.ItemsChoiceType8.Location);

        applicationData.Items = items.ToArray();
        applicationData.ItemsElementName = itemsElementName.ToArray();
        var syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(GetInitialSyncKey(User1Information.CalendarCollectionId), User1Information.CalendarCollectionId, applicationData);

        var syncAddResponse = ASAIRSAdapter.Sync(syncAddRequest);
        Site.Assert.IsTrue(syncAddResponse.AddResponses[0].Status.Equals("1"), "The sync add operation should be success; It is:{0} actually", syncAddResponse.AddResponses[0].Status);

        // Add the appointment to clean up list.
        RecordCaseRelativeItems(User1Information.UserName, User1Information.CalendarCollectionId, subject);
        #endregion

        #region Call Sync command to get the new added calendar item.
        var syncItem = GetSyncResult(subject, User1Information.CalendarCollectionId, null, null, null);
        #endregion

        #region Call ItemOperations command to reterive the added calendar item.
        GetItemOperationsResult(User1Information.CalendarCollectionId, syncItem.ServerId, null, null, null, null);
        #endregion

        #region Call Sync command to remove the location of the added calender item.
        // Create empty change items list.
        var changeItems = new List<object>();
        var changeItemsElementName = new List<Request.ItemsChoiceType7>();

        // Create an empty location.
        location = new Request.Location();

        // Add the location field name into the change items element name list.
        changeItemsElementName.Add(Request.ItemsChoiceType7.Location);
        // Add the empty location value to the change items value list.
        changeItems.Add(location);

        // Create sync collection change.
        var collectionChange = new Request.SyncCollectionChange
        {
            ServerId = syncItem.ServerId,
            ApplicationData = new Request.SyncCollectionChangeApplicationData
            {
                ItemsElementName = changeItemsElementName.ToArray(),
                Items = changeItems.ToArray()
            }
        };

        // Create change sync collection
        var collection = new Request.SyncCollection
        {
            SyncKey = SyncKey,
            CollectionId = User1Information.CalendarCollectionId,
            Commands = new object[] { collectionChange }
        };

        // Create change sync request.
        var syncChangeRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { collection });

        // Change the location of the added calender by Sync request.
        var syncChangeResponse = ASAIRSAdapter.Sync(syncChangeRequest);
        Site.Assert.IsTrue(syncChangeResponse.CollectionStatus.Equals(1), "The sync change operation should be success; It is:{0} actually", syncChangeResponse.CollectionStatus);

        #region Call Sync command to get the new changed calendar item that removed the location.
        syncItem = GetSyncResult(subject, User1Information.CalendarCollectionId, null, null, null);
        #endregion

        #region Call ItemOperations command to reterive the changed calendar item that removed the location.
        var itemOperations = GetItemOperationsResult(User1Information.CalendarCollectionId, syncItem.ServerId, null, null, null, null);
        #endregion
        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R1001013");

        // Verify MS-ASAIRS requirement: MS-ASAIRS_R1001013
        Site.CaptureRequirementIfIsNull(
            itemOperations.Calendar.Location1.DisplayName,
            1001013,
            @"[In Location] The client's request can include an empty Location element to remove the location from an item.");
        #endregion

        if (Common.IsRequirementEnabled(53, Site))
        {
            #region Call Search command to search the added calendar item.
            GetSearchResult(subject, User1Information.CalendarCollectionId, null, null, null);
            #endregion
        }
    }
    #endregion
}