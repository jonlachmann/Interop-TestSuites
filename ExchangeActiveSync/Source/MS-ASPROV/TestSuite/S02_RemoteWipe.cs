namespace Microsoft.Protocols.TestSuites.MS_ASPROV;

using System;
using Common;
using TestTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// This scenario is designed to test the remote wipe directive.
/// </summary>
[TestClass]
public class S02_RemoteWipe : TestSuiteBase
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

    /// <summary>
    /// This test case is intended to validate a successful remote wipe directive of Provision.
    /// </summary>
    [TestCategory("MSASPROV"), TestMethod]
    public void MSASPROV_S02_TC01_RemoteWipe()
    {
        #region Apply a unique DeviceType.
        // Switch the user credential to User1 to get user information.
        SwitchUser(User1Information, true);

        // Apply the unique DeviceType.
        DeviceType = $"{"ASPROV"}{DateTime.Now.ToString("mmssfff")}";
        PROVAdapter.ApplyDeviceType(DeviceType);
        CurrentUserInformation.UserName = User1Information.UserName;
        CurrentUserInformation.UserDomain = User1Information.UserDomain;

        #endregion

        #region Acknowledge the policy setting and set the device status on server to be wipe pending
        AcknowledgeSecurityPolicySettings();

        // Set the device status on server to be wipe pending.
        var userEmail = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);

        var dataWiped = PROVSUTControlAdapter.WipeData(SutComputerName, userEmail, User1Information.UserPassword, DeviceType);
        Site.Assert.IsTrue(dataWiped, "The data on the device with DeviceType {0} should be wiped successfully.", DeviceType);
        #endregion

        #region Perform an initial remote wipe
        // Send an empty Provision request to indicate a remote wipe operation on client.
        var emptyRequest = new ProvisionRequest();
        var provisionResponse = PROVAdapter.Provision(emptyRequest);

        Site.Assert.IsNotNull(provisionResponse, "If the Provision command executes successfully, the response from server should not be null.");
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R653");

        // Verify MS-ASPROV requirement: MS-ASPROV_R653
        // The RemoteWipe element is not null, so this requirement can be captured.
        Site.CaptureRequirementIfIsNotNull(
            provisionResponse.ResponseData.RemoteWipe,
            653,
            @"[In Responding to an Initial Request] The RemoteWipe or AccountOnlyRemoteWipe MUST only be included if a remote wipe or an account only remote wipe has been requested for the client.");
        #endregion

        #region Perform a failure remote wipe acknowledgment
        // Set the remote wipe status to 2 to indicate a remote wipe failure on client.
        var wipeRequest = new ProvisionRequest
        {
            RequestData =
            {
                RemoteWipe = new Microsoft.Protocols.TestSuites.Common.Request.ProvisionRemoteWipe
                {
                    Status = 2
                }
            }
        };

        provisionResponse = PROVAdapter.Provision(wipeRequest);

        if (Common.IsRequirementEnabled(1042, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                provisionResponse.ResponseData.Status,
                1042,
                @"[In Appendix B: Product Behavior]  If the client reports failure, the implementation does return a value of 2 in the Status element [and a remote wipe directive]. (<4> Section 3.2.5.1.2.2:  In Exchange 2007 and Exchange 2010, if the client reports failure, the server returns a value of 1 in the Status element.)");
        }

        if (Common.IsRequirementEnabled(1048, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionResponse.ResponseData.Status,
                1048,
                @"[In Appendix B: Product Behavior] If the client reports failure, the implementation does return a value of 2 in the Status element [and a remote wipe directive]. (Exchange 2013 and above follow this behavior.)");
        }

        // Send an empty Provision request to indicate a remote wipe operation on client.
        provisionResponse = PROVAdapter.Provision(emptyRequest);
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        if (Common.IsRequirementEnabled(702, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R702");

            // Verify MS-ASPROV requirement: MS-ASPROV_R702
            // The RemoteWipe element is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                provisionResponse.ResponseData.RemoteWipe,
                702,
                @"[In Appendix B: Product Behavior] If the client reports failure, the implementation does return [a value of 2 in the Status element and] a remote wipe directive. (Exchange 2007 and above follow this behavior.)");
        }
        #endregion

        #region Perform a successful remote wipe acknowledgment
        // Set the remote wipe status to 1 to indicate a successful wipe on client.
        wipeRequest.RequestData.RemoteWipe.Status = 1;
        var wipeResponse = PROVAdapter.Provision(wipeRequest);

        if (Common.IsRequirementEnabled(1041, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                wipeResponse.ResponseData.Status,
                1041,
                @"[In Appendix B: Product Behavior] If the client reports success, the implementation does return a value of 1 in the Status element (section 2.2.2.54.2). (<3> Section 3.2.5.1.2.2:  In Exchange 2007 and Exchange 2010, if the client reports success, the server returns a value of 1 in the Status element and a remote wipe directive.)");
        }

        if (Common.IsRequirementEnabled(1047, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                wipeResponse.ResponseData.Status,
                1047,
                @"[In Appendix B: Product Behavior] If the client reports success, the implementation does return a value of 1 in the Status element (section 2.2.2.54.2). (Exchange 2013 and above follow this behavior.)");
        }

        // Record the provision confirmation mail for user1 to the item collection of User1.
        var confirmationMailSubject = "Remote Device Wipe Confirmation";
        var inboxItemForUser1 = Common.RecordCreatedItem(User1Information.InboxCollectionId, confirmationMailSubject);
        User1Information.UserCreatedItems.Add(inboxItemForUser1);
        var sentItemForUser1 = Common.RecordCreatedItem(User1Information.SentItemsCollectionId, confirmationMailSubject);
        User1Information.UserCreatedItems.Add(sentItemForUser1);
        #endregion

        #region Remove the device from server and perform another initial remote wipe
        // Remove the device from the mobile list after wipe operation is successful.
        var deviceRemoved = PROVSUTControlAdapter.RemoveDevice(SutComputerName, userEmail, User1Information.UserPassword, DeviceType);
        Site.Assert.IsTrue(deviceRemoved, "The device with DeviceType {0} should be removed successfully.", DeviceType);

        // Send an empty Provision request when the client is not requested for a remote wipe.
        provisionResponse = PROVAdapter.Provision(emptyRequest);
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        // Add the debug information
        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R654");

        // Verify MS-ASPROV requirement: MS-ASPROV_R654
        // The RemoteWipe element is null, so this requirement can be captured.
        Site.CaptureRequirementIfIsNull(
            provisionResponse.ResponseData.RemoteWipe,
            654,
            @"[In Responding to an Initial Request] Otherwise [if a remote wipe has not been requested for the client], it [RemoteWipe] MUST be omitted");
        #endregion
    }

    /// <summary>
    /// This test case is intended to validate a successful account only remote wipe directive of Provision.
    /// </summary>
    [TestCategory("MSASPROV"), TestMethod]
    public void MSASPROV_S02_TC02_AccountOnlyRemoteWipe()
    {
        Site.Assume.AreEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "The AccountOnlyRemoteWipe element is supported when the ActiveSyncProtocolVersion is 16.1.");
            
        #region Apply a unique DeviceType.
        // Switch the user credential to User1 to get user information.
        SwitchUser(User1Information, true);

        // Apply the unique DeviceType.
        DeviceType = $"{"ASPROV"}{DateTime.Now.ToString("mmssfff")}";
        PROVAdapter.ApplyDeviceType(DeviceType);
        CurrentUserInformation.UserName = User1Information.UserName;
        CurrentUserInformation.UserDomain = User1Information.UserDomain;
        #endregion

        #region Acknowledge the policy setting and set the device status on server to be wipe pending
        AcknowledgeSecurityPolicySettings();

        // Set the device status on server to be wipe pending.
        var userEmail = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);

        var dataWiped = PROVSUTControlAdapter.AccountOnlyWipeData(SutComputerName, userEmail, User1Information.UserPassword, DeviceType);
        Site.Assert.IsTrue(dataWiped, "The data on the device with DeviceType {0} should be wiped successfully.", DeviceType);
        #endregion

        #region Perform an initial account only remote wipe
        // Send an empty Provision request to indicate an account only remote wipe operation on client.
        var emptyRequest = new ProvisionRequest();
        var provisionResponse = PROVAdapter.Provision(emptyRequest);

        Site.Assert.IsNotNull(provisionResponse, "If the Provision command executes successfully, the response from server should not be null.");
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        if (Common.IsRequirementEnabled(66613, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R66613");

            // Verify MS-ASPROV requirement: MS-ASPROV_R66613
            // The AccountOnlyRemoteWipe element is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                provisionResponse.ResponseData.AccountOnlyRemoteWipe,
                66613,
                @"[In Appendix B: Product Behavior] The [RemoteWipe or] AccountOnlyRemoteWipe MUST only be included if [a remote wipe or] an account only remote wipe has been requested for the client. (Exchange 2019 follow this behavior.)");
        }
        #endregion

        #region Perform a failure account only remote wipe acknowledgment
        // Set the account only remote wipe status to 2 to indicate an account only remote wipe failure on client.
        var wipeRequest = new ProvisionRequest
        {
            RequestData =
            {
                AccountOnlyRemoteWipe = new Microsoft.Protocols.TestSuites.Common.Request.ProvisionAccountOnlyRemoteWipe
                {
                    Status = 2
                }
            }
        };

        provisionResponse = PROVAdapter.Provision(wipeRequest);

        if (Common.IsRequirementEnabled(66610, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                provisionResponse.ResponseData.Status,
                66610,
                @"[In Appendix B: Product Behavior]  If the client reports failure, the server SHOULD return a value of 2 in the Status element[ and an account only remote wipe directive]. (Exchange 2019 follow this behavior.)");
        }

        // Send an empty Provision request to indicate an account only remote wipe operation on client.
        provisionResponse = PROVAdapter.Provision(emptyRequest);
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        if (Common.IsRequirementEnabled(66611, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R66611");

            // Verify MS-ASPROV requirement: MS-ASPROV_R66611
            // The AccountOnlyRemoteWipe element is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                provisionResponse.ResponseData.AccountOnlyRemoteWipe,
                66611,
                @"[In Appendix B: Product Behavior] If the client reports failure, the server SHOULD return [a value of 2 in the Status element and ]an account only remote wipe directive. (Exchange 2019 follow this behavior.)");
        }
        #endregion

        #region Perform a successful account only remote wipe acknowledgment
        // Set the account only remote wipe status to 1 to indicate a successful wipe on client.
        wipeRequest.RequestData.AccountOnlyRemoteWipe.Status = 1;
        var wipeResponse = PROVAdapter.Provision(wipeRequest);

        if (Common.IsRequirementEnabled(66609, Site))
        {
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                wipeResponse.ResponseData.Status,
                66609,
                @"[In Appendix B: Product Behavior] If the client reports success, the server SHOULD return a value of 1 in the Status element (section 2.2.2.54.2). (Exchange 2019 follow this behavior.)");
        }

        // Record the provision confirmation mail for user1 to the item collection of User1.
        var confirmationMailSubject = "Remote Device Wipe Confirmation";
        var inboxItemForUser1 = Common.RecordCreatedItem(User1Information.InboxCollectionId, confirmationMailSubject);
        User1Information.UserCreatedItems.Add(inboxItemForUser1);
        var sentItemForUser1 = Common.RecordCreatedItem(User1Information.SentItemsCollectionId, confirmationMailSubject);
        User1Information.UserCreatedItems.Add(sentItemForUser1);
        #endregion

        #region Remove the device from server and perform another initial account only remote wipe
        // Remove the device from the mobile list after wipe operation is successful.
        var deviceRemoved = PROVSUTControlAdapter.RemoveDevice(SutComputerName, userEmail, User1Information.UserPassword, DeviceType);
        Site.Assert.IsTrue(deviceRemoved, "The device with DeviceType {0} should be removed successfully.", DeviceType);

        // Send an empty Provision request when the client is not requested for an account only remote wipe.
        provisionResponse = PROVAdapter.Provision(emptyRequest);
        Site.Assert.AreEqual<byte>(1, provisionResponse.ResponseData.Status, "The server should return status code 1 to indicate a success.");

        if (Common.IsRequirementEnabled(66615, Site))
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R66615");

            // Verify MS-ASPROV requirement: MS-ASPROV_R66615
            // The AccountOnlyRemoteWipe element is null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNull(
                provisionResponse.ResponseData.AccountOnlyRemoteWipe,
                66615,
                @"[In Appendix B: Product Behavior] Otherwise [if an account only remote wipe has not been requested for the client], it [Account Only RemoteWipe] MUST be omitted. (Exchange 2019 follow this behavior.)");
        }
        #endregion
    }
}