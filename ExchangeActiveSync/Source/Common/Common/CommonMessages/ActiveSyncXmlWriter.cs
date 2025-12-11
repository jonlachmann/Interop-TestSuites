namespace Microsoft.Protocols.TestSuites.Common;

using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

/// <summary>
/// ActiveSync XML writer.
/// </summary>
public class ActiveSyncXmlWriter : XmlTextWriter
{
    /// <summary>
    /// The value indicates whether to skip writing the content.
    /// </summary>
    private bool isSkip;

    /// <summary>
    /// The current element
    /// </summary>
    private string currentElement;

    /// <summary>
    /// Initializes a new instance of the ActiveSyncXmlWriter class.
    /// </summary>
    /// <param name="stream">The steam be write.</param>
    /// <param name="encoding">Represents a character encoding.</param>
    public ActiveSyncXmlWriter(Stream stream, Encoding encoding)
        : base(stream, encoding)
    {
        isSkip = false;
    }

    /// <summary>
    /// Writes the start of an attribute.
    /// </summary>
    /// <param name="prefix">Namespace prefix of the attribute.</param>
    /// <param name="localName">LocalName of the attribute.</param>
    /// <param name="ns">NamespaceURI of the attribute</param>
    public override void WriteStartAttribute(string prefix, string localName, string ns)
    {
        if ("xmlns" == prefix && ("xsi" == localName || "xsd" == localName))
        {
            isSkip = true;
        }
        else
        {
            base.WriteStartAttribute(prefix, localName, ns);
        }
    }

    /// <summary>
    /// Writes the start of an element
    /// </summary>
    /// <param name="prefix">Namespace prefix of the element.</param>
    /// <param name="localName">LocalName of the element.</param>
    /// <param name="ns">NamespaceURI of the element</param>
    public override void WriteStartElement(string prefix, string localName, string ns)
    {
        currentElement = localName;
        base.WriteStartElement(prefix, localName, ns);
    }

    /// <summary>
    /// Writes the given text content.
    /// </summary>
    /// <param name="text">Text to write.</param>
    public override void WriteString(string text)
    {
        if (!isSkip)
        {
            if (IsCDATAValue())
            {
                WriteCData(text);
            }
            else
            {
                base.WriteString(text);
            }
        }
    }

    /// <summary>
    /// Closes the previous System.Xml.XmlTextWriter.WriteStartAttribute(System.String,System.String,System.String) call.
    /// </summary>
    public override void WriteEndAttribute()
    {
        if (!isSkip)
        {
            base.WriteEndAttribute();
        }
        else
        {
            isSkip = false;
        }
    }

    /// <summary>
    /// Override this, true or false will be 1 or 0
    /// </summary>
    /// <param name="value">The value will be change.</param>
    public override void WriteValue(bool value)
    {
        if (value == true)
        {
            WriteValue("1");
        }
        else
        {
            WriteValue("0");
        }
    }

    /// <summary>
    /// Writes raw markup manually from a string.(true or false will be 1 or 0)
    /// </summary>
    /// <param name="data">String containing the text to write.</param>
    public override void WriteRaw(string data)
    {
        switch (data)
        {
            case "true":
                WriteRaw("1");
                break;
            case "false":
                WriteRaw("0");
                break;
            default:
                switch (currentElement)
                {
                    case "DueDate":
                    case "StartDate":
                    case "UtcStartDate":
                    case "UtcDueDate":
                    case "DateCompleted":
                    case "CompleteTime":
                    case "ReminderTime":
                    case "DateReceived":
                    case "ExceptionStartTime":
                    case "StartTime":
                    case "EndTime":
                    case "DtStamp":
                    case "AppointmentReplyTime":
                    case "Until":
                    case "OrdinalDate":
                    case "Start":
                    case "InstanceId":
                    case "RecurrenceId":
                    case "ContentExpiryDate":
                    case "Anniversary":
                    case "Birthday":
                        data = DateTime.Parse(data).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                        break;
                }

                base.WriteRaw(data);
                break;
        }
    }

    /// <summary>
    /// Verifies whether the current element is a CDATA value or not.
    /// </summary>
    /// <returns>A boolean value indicates the element is a CDATA value or not.</returns>
    private bool IsCDATAValue()
    {
        return currentElement.ToLower(CultureInfo.CurrentCulture) == "mime" ||
               currentElement.ToLower(CultureInfo.CurrentCulture) == "conversationid" ||
               currentElement.ToLower(CultureInfo.CurrentCulture) == "conversationindex";
    }
}