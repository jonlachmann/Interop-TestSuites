namespace Microsoft.Protocols.TestSuites.Common;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using TestTools;

/// <summary>
/// This class is partial class ActiveSyncClient, mainly aims to verify the response XML whether consistent with the schemas.
/// </summary>
public partial class ActiveSyncClient
{
    #region Private Fields

    /// <summary>
    /// Schema list field
    /// </summary>
    private string[] xmlSchemaList = null;

    /// <summary>
    /// The error results of XML Schema validation
    /// </summary>
    private Collection<ValidationEventArgs> xmlValidationErrors;

    /// <summary>
    /// The warning results of XML Schema validation
    /// </summary>
    private Collection<ValidationEventArgs> xmlValidationWarnings;

    /// <summary>
    /// The last XML soap message string
    /// </summary>
    private string lastRawString;

    /// <summary>
    /// The schema validation result, True/False
    /// </summary>
    private bool validationResult;

    #endregion

    #region Public Properties

    /// <summary>
    /// Gets a value indicating whether the last server response if success or not
    /// </summary>
    public bool ValidationResult
    {
        get
        {
            return validationResult;
        }
    }

    /// <summary>
    /// Gets the error events generated during XML validation.
    /// </summary>
    public Collection<ValidationEventArgs> XmlValidationErrors
    {
        get
        {
            return xmlValidationErrors;
        }
    }

    /// <summary>
    /// Gets the warning events generated during XML validation.
    /// </summary>
    public Collection<ValidationEventArgs> XmlValidationWarnings
    {
        get
        {
            return xmlValidationWarnings;
        }
    }

    /// <summary>
    /// Gets the used schema list
    /// </summary>
    public string[] XmlSchemaList
    {
        get
        {
            if (xmlSchemaList == null)
            {
                var list = new List<string>
                {
                    ReadXsdFile("AirSync.xsd"),
                    ReadXsdFile("AirSyncBase.xsd"),
                    ReadXsdFile("AutodiscoverMobileSync.xsd"),
                    ReadXsdFile("AutodiscoverResponse.xsd"),
                    ReadXsdFile("Calendar.xsd"),
                    ReadXsdFile("ComposeMail.xsd"),
                    ReadXsdFile("Contacts.xsd"),
                    ReadXsdFile("Contacts2.xsd"),
                    ReadXsdFile("DocumentLibrary.xsd"),
                    ReadXsdFile("Email.xsd"),
                    ReadXsdFile("Email2.xsd"),
                    ReadXsdFile("FolderCreateResponse.xsd"),
                    ReadXsdFile("FolderDeleteResponse.xsd"),
                    ReadXsdFile("FolderHierarchy.xsd"),
                    ReadXsdFile("FolderSyncResponse.xsd"),
                    ReadXsdFile("FolderUpdateResponse.xsd"),
                    ReadXsdFile("GAL.xsd"),
                    ReadXsdFile("GetHierarchyResponse.xsd"),
                    ReadXsdFile("GetItemEstimateResponse.xsd"),
                    ReadXsdFile("ItemOperations.xsd"),
                    ReadXsdFile("ItemOperationsResponse.xsd"),
                    ReadXsdFile("MeetingResponseResponse.xsd"),
                    ReadXsdFile("MoveItemsResponse.xsd"),
                    ReadXsdFile("Notes.xsd"),
                    ReadXsdFile("PingResponse.xsd"),
                    ReadXsdFile("Provision.xsd"),
                    ReadXsdFile("ProvisionResponse.xsd"),
                    ReadXsdFile("ResolveRecipientsResponse.xsd"),
                    ReadXsdFile("RightsManagement.xsd"),
                    ReadXsdFile("Search.xsd"),
                    ReadXsdFile("SearchResponse.xsd"),
                    ReadXsdFile("SendMailResponse.xsd"),
                    ReadXsdFile("Settings.xsd"),
                    ReadXsdFile("SettingsResponse.xsd"),
                    ReadXsdFile("SmartForwardResponse.xsd"),
                    ReadXsdFile("SmartReplyResponse.xsd"),
                    ReadXsdFile("SyncResponse.xsd"),
                    ReadXsdFile("Tasks.xsd"),
                    ReadXsdFile("ValidateCertResponse.xsd"),
                    ReadXsdFile("Find.xsd"),
                    ReadXsdFile("FindResponse.xsd")
                };

                xmlSchemaList = list.ToArray();
            }

            return xmlSchemaList;
        }
    }

    #endregion

    #region Schema Validation method

    /// <summary>
    /// The schema validation for the XML string.
    /// </summary>
    /// <param name="decodeXMLstring">The XML string which is decoded from WBXML format</param>
    /// <param name="testSite">An instance of interface ITestSite which provides logging, assertions, and adapters for test code onto its execution context.</param>
    public void ValidateResponseSchema(string decodeXMLstring, ITestSite testSite)
    {
        if (string.IsNullOrEmpty(decodeXMLstring))
        {
            validationResult = true;
            return;
        }

        lastRawString = decodeXMLstring;

        // Initialize Validation Result Recorder
        validationResult = false;
        xmlValidationWarnings = new Collection<ValidationEventArgs>();
        xmlValidationErrors = new Collection<ValidationEventArgs>();

        // Prepare settings of XML reader
        var settings = new XmlReaderSettings();
        foreach (var xmlSchema in XmlSchemaList)
        {
            using (var stringReader = new StringReader(xmlSchema))
            {
                settings.Schemas.Add(GetTargetNamespace(xmlSchema), XmlReader.Create(stringReader));
            }
        }

        settings.ValidationType = ValidationType.Schema;
        settings.ConformanceLevel = ConformanceLevel.Document;
        settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
        settings.ValidationEventHandler += new ValidationEventHandler(ValidationCallBack);

        // Load into the memory stream
        using (var ms = new MemoryStream(System.Text.ASCIIEncoding.UTF8.GetBytes(lastRawString)))
        {
            // Create XML reader for SOAP client message
            var xmlReader = XmlReader.Create(ms, settings);

            while (xmlReader.Read())
            {
            }
        }

        validationResult = xmlValidationErrors.Count == 0 && xmlValidationWarnings.Count == 0;

        if (!validationResult)
        {
            string errorInformation = null;

            string warningInformation = null;

            if (xmlValidationErrors.Count > 0)
            {
                if (xmlValidationErrors.Count == 1)
                {
                    errorInformation = string.Format("There is a schema validation error:\r\n");
                }
                else
                {
                    errorInformation = string.Format("There are {0} schema validation errors:\r\n", xmlValidationErrors.Count);
                }

                for (var i = 0; i < xmlValidationErrors.Count; i++)
                {
                    errorInformation = errorInformation + string.Format("Error: {0}\r\n", xmlValidationErrors[i].Message);
                }
            }

            if (xmlValidationWarnings.Count > 0)
            {
                if (xmlValidationWarnings.Count == 1)
                {
                    warningInformation = string.Format("There is a schema validation warning:\r\n");
                }
                else
                {
                    warningInformation = string.Format("There are {0} schema validation warnings:\r\n", xmlValidationWarnings.Count);
                }

                for (var i = 0; i < xmlValidationWarnings.Count; i++)
                {
                    warningInformation = warningInformation + string.Format("Warning: {0}\r\n", xmlValidationWarnings[i].Message);
                }
            }

            testSite.Assert.Fail(errorInformation + warningInformation);
        }
    }

    #endregion

    #region Private Methods
    /// <summary>
    /// Read the xsd file content
    /// </summary>
    /// <param name="fileName">Specific the XSD file path</param>
    /// <returns>The content string of the specified file name</returns>
    private static string ReadXsdFile(string fileName)
    {
        FileStream fs = null;
        try
        {
            fs = File.OpenRead(fileName);
            using (var sr = new StreamReader(fs))
            {
                return sr.ReadToEnd();
            }
        }
        finally
        {
            if (fs != null)
            {
                fs.Dispose();
            }
        }
    }

    /// <summary>
    /// Get the target namespace of the schema
    /// </summary>
    /// <param name="schema">The schema string</param>
    /// <returns>The target name of the give schema</returns>
    private static string GetTargetNamespace(string schema)
    {
        var xd = new XmlDocument();
        xd.LoadXml(schema);

        if (xd.ChildNodes.Count <= 1)
        {
            throw new XmlException("Can't find schema element in " + xd.Name);
        }

        if (xd.ChildNodes[1].Attributes != null)
        {
            var attributeList = xd.ChildNodes[1].Attributes;
            return attributeList["targetNamespace"] == null ? string.Empty : attributeList["targetNamespace"].Value;
        }

        return null;
    }

    /// <summary>
    /// The callback method that will handle XML schema validation events.
    /// </summary>
    /// <param name="sender">The source of the event.</param>
    /// <param name="args"> A ValidationEventArgs containing the event data.</param>
    private void ValidationCallBack(object sender, ValidationEventArgs args)
    {
        if (args.Severity == XmlSeverityType.Error)
        {
            xmlValidationErrors.Add(args);
        }
        else
        {
            xmlValidationWarnings.Add(args);
        }
    }
    #endregion
}