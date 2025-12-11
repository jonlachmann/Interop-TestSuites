namespace Microsoft.Protocols.TestSuites.Common;

using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

/// <summary>
/// The ActiveSync response.
/// </summary>
/// <typeparam name="T">The generic type.</typeparam>
public abstract class ActiveSyncResponseBase<T>
{
    /// <summary>
    /// Gets or sets response data.
    /// </summary>
    public T ResponseData { get; set; }

    /// <summary>
    /// Gets or sets status code.
    /// </summary>
    public HttpStatusCode StatusCode { get; set; }

    /// <summary>
    /// Gets or sets status description.
    /// </summary>
    public string StatusDescription { get; set; }

    /// <summary>
    /// Gets or sets web header collection.
    /// </summary>
    public WebHeaderCollection Headers { get; set; }

    /// <summary>
    /// Gets or sets raw body.
    /// </summary>
    public byte[] RawBody { get; set; }

    /// <summary>
    /// Gets or sets response data xml.
    /// </summary>
    public string ResponseDataXML { get; set; }

    /// <summary>
    /// Deserialize response data.
    /// </summary>
    public virtual void DeserializeResponseData()
    {
        if (!string.IsNullOrEmpty(ResponseDataXML))
        {
            if (typeof(T) == typeof(Response.ValidateCert))
            {
                return;
            }

            var stringResponse = EscapeCharactor(ResponseDataXML);

            StringReader stringReader = null;
            try
            {
                stringReader = new StringReader(stringResponse);
                using (var xmlTextReader = new XmlTextReader(stringReader))
                {
                    xmlTextReader.WhitespaceHandling = WhitespaceHandling.All;
                    var xmlSerializer = new XmlSerializer(typeof(T));
                    var deserializedObject = xmlSerializer.Deserialize(xmlTextReader);
                    if (deserializedObject is T)
                    {
                        ResponseData = (T)deserializedObject;
                    }
                }
            }
            finally
            {
                if (stringReader != null)
                {
                    stringReader.Dispose();
                }
            }
        }
    }

    /// <summary>
    /// Replaces CDATA string.
    /// </summary>
    /// <param name="original">The original string</param>
    /// <returns>The replaced string.</returns>
    private string EscapeCharactor(string original)
    {
        var regex = new Regex(@"\<!\[CDATA\[.+?\]\]\>", RegexOptions.Singleline);
        return regex.Replace(original, new MatchEvaluator(RemoveCDataTag));
    }

    /// <summary>
    /// Remove CData Tag.
    /// </summary>
    /// <param name="match">A regular expression</param>
    /// <returns>The sub string</returns>
    private string RemoveCDataTag(Match match)
    {
        return match.Value.Substring(9, match.Length - 12);
    }
}
