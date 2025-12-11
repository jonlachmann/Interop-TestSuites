namespace Microsoft.Protocols.TestSuites.MS_ASRM;

using System;
using System.Collections.ObjectModel;
using System.Globalization;
using DataStructures=Common.DataStructures;
using Common;
using TestTools;
using Request = Common.Request;
using Response = Common.Response;

/// <summary>
/// A bass class for scenario class.
/// </summary>
public class TestSuiteBase : TestClassBase
{
    #region Variables
    /// <summary>
    /// Gets or sets the related information of User1.
    /// </summary>
    protected UserInformation UserOneInformation { get; set; }

    /// <summary>
    /// Gets or sets the related information of User2.
    /// </summary>
    protected UserInformation UserTwoInformation { get; set; }

    /// <summary>
    /// Gets or sets the related information of User3.
    /// </summary>
    protected UserInformation UserThreeInformation { get; set; }

    /// <summary>
    /// Gets or sets the related information of User4.
    /// </summary>
    protected UserInformation UserFourInformation { get; set; }

    /// <summary>
    /// Gets the MS-ASRM protocol adapter.
    /// </summary>
    protected IMS_ASRMAdapter ASRMAdapter { get; private set; }

    /// <summary>
    /// Gets MS-ASRM SUT Control adapter.
    /// </summary>
    protected IMS_ASRMSUTControlAdapter ASRMSUTControlAdapter { get; private set; }
    #endregion

    /// <summary>
    /// Add the item created in test case to the item collections needed to clean up.
    /// </summary>
    /// <param name="userInformation">The identity of the user who has the item.</param>
    /// <param name="parentFolder">The CollectionId of the folder in which the item is placed.</param>
    /// <param name="itemSubject">The subject of the item to delete.</param>
    protected static void AddCreatedItemToCollection(UserInformation userInformation, string parentFolder, string itemSubject)
    {
        var createdItems = new CreatedItems { CollectionId = parentFolder };
        createdItems.ItemSubject.Add(itemSubject);
        userInformation.UserCreatedItems.Add(createdItems);
    }

    #region Test case initialize and cleanup
    /// <summary>
    /// Override the base TestInitialize function
    /// </summary>
    protected override void TestInitialize()
    {
        base.TestInitialize();

        if (ASRMAdapter == null)
        {
            ASRMAdapter = Site.GetAdapter<IMS_ASRMAdapter>();
        }

        ASRMSUTControlAdapter = Site.GetAdapter<IMS_ASRMSUTControlAdapter>();

        // If implementation doesn't support this specification [MS-ASRM], the case will not start.
        if (!bool.Parse(Common.GetConfigurationPropertyValue("MS-ASRM_Supported", Site)))
        {
            Site.Assert.Inconclusive("This test suite is not supported under current SUT, because MS-ASRM_Supported value is set to false in MS-ASRM_{0}_SHOULDMAY.deployment.ptfconfig file.", Common.GetSutVersion(Site));
        }

        // Set the information of User1.
        UserOneInformation = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User1Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User1Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        // Set the information of User2.
        UserTwoInformation = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User2Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User2Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        // Set the information of User3.
        UserThreeInformation = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User3Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User3Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };

        // Set the information of User4.
        UserFourInformation = new UserInformation
        {
            UserName = Common.GetConfigurationPropertyValue("User4Name", Site),
            UserPassword = Common.GetConfigurationPropertyValue("User4Password", Site),
            UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
        };
        var sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", Site);

        if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASRM_Supported", Site)))
        {
            if (Common.GetConfigurationPropertyValue("TransportType", Site)
                .Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
            {
                // Use the user who is in Administrators group to enable the SSL setting.
                var isSSLUpdated = ASRMSUTControlAdapter.ConfigureSSLSetting(
                    sutComputerName,
                    UserFourInformation.UserName,
                    UserFourInformation.UserPassword,
                    UserFourInformation.UserDomain,
                    true);
                Site.Assert.IsTrue(isSSLUpdated, "The SSL setting of protocol web service should be enabled.");
            }
        }

        if (Common.GetSutVersion(Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site), "12.1"))
        {
            SwitchUser(UserOneInformation, true);
        }
    }

    /// <summary>
    /// Override the base TestCleanup function
    /// </summary>
    protected override void TestCleanup()
    {
        // If implementation doesn't support this specification [MS-ASRM], the case will not start.
        if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASRM_Supported", Site)))
        {
            var sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", Site);

            if (Common.GetConfigurationPropertyValue("TransportType", Site)
                .Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
            {
                // Use the user who is in Administrators group to disable the SSL setting.
                var isSSLUpdated = ASRMSUTControlAdapter.ConfigureSSLSetting(
                    sutComputerName,
                    UserFourInformation.UserName,
                    UserFourInformation.UserPassword,
                    UserFourInformation.UserDomain,
                    false);
                Site.Assert.IsTrue(isSSLUpdated, "The SSL setting of protocol web service should be disabled.");
            }

            // Clean the created items of User1.
            if (UserOneInformation.UserCreatedItems.Count != 0)
            {
                // Switch the user credential to User1.
                SwitchUser(UserOneInformation, false);
                DeleteCreatedItems(UserOneInformation.UserCreatedItems);
            }

            if (UserTwoInformation.UserCreatedItems.Count != 0)
            {
                // Switch the user credential to User2.
                SwitchUser(UserTwoInformation, false);
                DeleteCreatedItems(UserTwoInformation.UserCreatedItems);
            }

            if (UserThreeInformation.UserCreatedItems.Count != 0)
            {
                // Switch the user credential to User3.
                SwitchUser(UserThreeInformation, false);
                DeleteCreatedItems(UserThreeInformation.UserCreatedItems);
            }
        }

        base.TestCleanup();
    }

    #endregion

    #region protected methods
    /// <summary>
    /// Checks if ActiveSync Protocol Version is "14.1" and Transport Type is "HTTPS".
    /// </summary>
    protected void CheckPreconditions()
    {
        Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("14.1") ||
                           Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("16.0") ||
                           Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site).Equals("16.1"),
            "Implementation does consider the XML body of the command request to be invalid, if the protocol version specified by in the command request is not 14.1 and 16.0.");
        Site.Assume.AreEqual<string>("HTTPS", Common.GetConfigurationPropertyValue("TransportType", Site).ToUpper(CultureInfo.CurrentCulture), "This protocol requires that communication between the client and server occurs over an HTTP connection that uses Secure Sockets Layer (SSL).");
    }

    /// <summary>
    /// Call Settings command to get the expected template ID for template name
    /// </summary>
    /// <param name="templateName">A string that specifies the name of the rights policy template.</param>
    /// <returns>A string that identifies a particular rights policy template to be applied to the outgoing message.</returns>
    protected string GetTemplateID(string templateName)
    {
        // Get the template settings
        var settingsResponse = ASRMAdapter.Settings();

        // Choose the all rights policy template and get template ID.
        Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation, "The RightsManagementInformation element should not be null in Settings response.");
        Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation.Get, "The Get element should not be null in Settings response.");
        Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation.Get.RightsManagementTemplates, "The RightsManagementTemplates element should not be null in Settings response.");
        string templateID = null;
        foreach (var template in settingsResponse.ResponseData.RightsManagementInformation.Get.RightsManagementTemplates)
        {
            if (template.TemplateName == templateName)
            {
                templateID = template.TemplateID;
                break;
            }
        }

        Site.Assert.IsNotNull(templateID, "Template {0} is not found on the server. This may happen if MS-ASRM configuration is not performed properly.", templateName);
        return templateID;
    }

    /// <summary>
    /// Find an email with specific subject and folder.
    /// </summary>
    /// <param name="subject">The subject of the email item.</param>
    /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
    /// <param name="rightsManagementSupport">A boolean value specifies whether the server will decompress and decrypt rights-managed email messages before sending them to the client or not</param>
    /// <param name="isRetryNeeded">A boolean value specifies whether need retry.</param>
    /// <returns>Return change result</returns>
    protected DataStructures.Sync SyncEmail(string subject, string collectionId, bool? rightsManagementSupport, bool isRetryNeeded)
    {
        var syncRequest = Common.CreateInitialSyncRequest(collectionId);
        var initSyncResult = ASRMAdapter.Sync(syncRequest);

        // Verify sync change result
        Site.Assert.AreEqual<byte>(1, initSyncResult.CollectionStatus, "If the Sync command executes successfully, the Status in response should be 1.");

        syncRequest = TestSuiteHelper.CreateSyncRequest(initSyncResult.SyncKey, collectionId, rightsManagementSupport);
        var sync = ASRMAdapter.SyncEmail(syncRequest, subject, isRetryNeeded);
        return sync;
    }

    /// <summary>
    /// Sync changes between client and server
    /// </summary>
    /// <param name="syncKey">The synchronization key returned by last request.</param>
    /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
    /// <param name="rightsManagementSupport">A boolean value specifies whether the server will decompress and decrypt rights-managed email messages before sending them to the client or not</param>
    /// <returns>Return change result</returns>
    protected DataStructures.SyncStore SyncChanges(string syncKey, string collectionId, bool rightsManagementSupport)
    {
        // Get changes from server use initial syncKey
        var syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, rightsManagementSupport);
        var syncResult = ASRMAdapter.Sync(syncRequest);

        return syncResult;
    }

    /// <summary>
    /// Change user to call FolderSync command to synchronize the collection hierarchy.
    /// </summary>
    /// <param name="userInformation">The information of a user.</param>
    /// <param name="syncFolderHierarchy">Whether sync folder hierarchy or not.</param>
    protected void SwitchUser(UserInformation userInformation, bool syncFolderHierarchy)
    {
        ASRMAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

        if (syncFolderHierarchy)
        {
            var folderSyncResponse = ASRMAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

            // Get the CollectionId from FolderSync command response.
            if (string.IsNullOrEmpty(userInformation.InboxCollectionId))
            {
                userInformation.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, Site);
            }

            if (string.IsNullOrEmpty(userInformation.SentItemsCollectionId))
            {
                userInformation.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, Site);
            }
        }
    }

    /// <summary>
    /// User1 sends mail to User2 and does FolderSync in User2's mailbox
    /// </summary>
    /// <param name="templateID">A string that identifies a particular rights policy template to be applied to the outgoing message.</param>
    /// <param name="saveInSentItems">A boolean that represent to save the sent item in SentItems folder or not.</param>
    /// <param name="copyToUserInformation">The UserInformation for the Cc element.</param>
    /// <returns>The subject of the sent message</returns>
    protected string SendMailAndFolderSync(string templateID, bool saveInSentItems, UserInformation copyToUserInformation)
    {
        #region User1 calls method SendMail to send MIME-formatted e-mail messages to User2
        var subject = Common.GenerateResourceName(Site, "subject");
        var content = "The content of the body.";
        SendMailRequest sendMailRequest;

        if (copyToUserInformation != null)
        {
            sendMailRequest = TestSuiteHelper.CreateSendMailRequest(Common.GetMailAddress(UserOneInformation.UserName, UserOneInformation.UserDomain), Common.GetMailAddress(UserTwoInformation.UserName, UserTwoInformation.UserDomain), Common.GetMailAddress(copyToUserInformation.UserName, copyToUserInformation.UserDomain), string.Empty, subject, content, templateID);
        }
        else
        {
            sendMailRequest = TestSuiteHelper.CreateSendMailRequest(Common.GetMailAddress(UserOneInformation.UserName, UserOneInformation.UserDomain), Common.GetMailAddress(UserTwoInformation.UserName, UserTwoInformation.UserDomain), string.Empty, string.Empty, subject, content, templateID);
        }

        if (saveInSentItems)
        {
            sendMailRequest.RequestData.SaveInSentItems = string.Empty;
        }

        var sendMailResponse = ASRMAdapter.SendMail(sendMailRequest);

        var counter = 0;
        var waitTime = int.Parse(Common.GetConfigurationPropertyValue("SSLWaitTime", Site));
        var upperBound = int.Parse(Common.GetConfigurationPropertyValue("SSLRetryCount", Site));

        if (Common.GetConfigurationPropertyValue("TransportType", Site)
            .Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
        {
            while (!string.IsNullOrEmpty(sendMailResponse.ResponseDataXML) && counter < upperBound)
            {
                // Await the SSL configuration take effect.
                Thread.Sleep(waitTime);
                sendMailResponse = ASRMAdapter.SendMail(sendMailRequest);
                counter++;
            }
        }

        Site.Assert.AreEqual<string>(string.Empty, sendMailResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");

        if (saveInSentItems)
        {
            AddCreatedItemToCollection(UserOneInformation, UserOneInformation.SentItemsCollectionId, subject);
        }
        #endregion

        #region Record Cc user name, folder collectionId and item subject that are used in case
        if (copyToUserInformation != null)
        {
            SwitchUser(copyToUserInformation, true);
            AddCreatedItemToCollection(copyToUserInformation, copyToUserInformation.InboxCollectionId, subject);
        }

        #endregion

        #region User2 calls method FolderSync to synchronize the collection hierarchy, returns folder collectionIds
        SwitchUser(UserTwoInformation, true);
        #endregion

        #region Record User2's user name, folder collectionId and item subject that are used in case
        AddCreatedItemToCollection(UserTwoInformation, UserTwoInformation.InboxCollectionId, subject);
        #endregion

        return subject;
    }
    #endregion

    #region Private methods
    /// <summary>
    /// Delete the specified item.
    /// </summary>
    /// <param name="itemsToDelete">The collection of the items to delete.</param>
    private void DeleteCreatedItems(Collection<CreatedItems> itemsToDelete)
    {
        foreach (var itemToDelete in itemsToDelete)
        {
            var syncRequest = Common.CreateInitialSyncRequest(itemToDelete.CollectionId);
            var initSyncResult = ASRMAdapter.Sync(syncRequest);
            var result = SyncChanges(initSyncResult.SyncKey, itemToDelete.CollectionId, false);
            var i = 0;
            if (result.AddElements != null)
            {
                var deletes = new Request.SyncCollectionDelete[result.AddElements.Count];
                foreach (var item in result.AddElements)
                {
                    foreach (var subject in itemToDelete.ItemSubject)
                    {
                        if (item.Email.Subject.Equals(subject))
                        {
                            var delete = new Request.SyncCollectionDelete
                            {
                                ServerId = item.ServerId
                            };
                            deletes[i] = delete;
                        }
                    }

                    i++;
                }

                var syncCollection = TestSuiteHelper.CreateSyncCollection(result.SyncKey, itemToDelete.CollectionId);
                syncCollection.Commands = deletes;

                syncRequest = Common.CreateSyncRequest([syncCollection]);
                var deleteResult = ASRMAdapter.Sync(syncRequest);
                Site.Assert.AreEqual<byte>(
                    1,
                    deleteResult.CollectionStatus,
                    "The value of 'Status' should be 1 which indicates the Sync command executes successfully.");
            }
        }
    }

    #endregion
}