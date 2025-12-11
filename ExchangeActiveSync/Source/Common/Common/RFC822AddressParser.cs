namespace Microsoft.Protocols.TestSuites.Common;

using System.Text.RegularExpressions;

/// <summary>
/// This class implements rfc 822 compliant email validator routines.
/// </summary>
public static class RFC822AddressParser
{
    /// <summary>
    /// The constant string for the Escape character
    /// </summary>
    private const string Escape = @"\\";

    /// <summary>
    /// The constant string for the Period
    /// </summary>
    private const string Period = @"\.";

    /// <summary>
    /// The constant string for the Space
    /// </summary>
    private const string Space = @"\040";

    /// <summary>
    /// The constant string for the Tab
    /// </summary>
    private const string Tab = @"\t";

    /// <summary>
    /// The constant string for the open brackets
    /// </summary>
    private const string OpenBr = @"\[";

    /// <summary>
    /// The constant string for the close brackets
    /// </summary>
    private const string CloseBr = @"\]";

    /// <summary>
    /// The constant string for the open parentheses
    /// </summary>
    private const string OpenParen = @"\(";

    /// <summary>
    /// The constant string for the close parentheses
    /// </summary>
    private const string CloseParen = @"\)";

    /// <summary>
    /// The constant string for the Non-ASCII characters
    /// </summary>
    private const string NonAscii = @"\x80-\xff";

    /// <summary>
    /// The constant string for the Ctrl
    /// </summary>
    private const string Ctrl = @"\000-\037";

    /// <summary>
    /// The constant string for the carriage return/line feed
    /// </summary>
    private const string CRLF = @"\n\015";

    /// <summary>
    /// The regex expression for address
    /// </summary>
    private static Regex addreg;

    /// <summary>
    /// Initializes static members of the RFC822AddressParser class
    /// </summary>
    static RFC822AddressParser()
    {
        // Initialize the regex expression
        InitialRegex();
    }

    /// <summary>
    /// Verify whether the specified email address is compliant with RFC822 or not
    /// </summary>
    /// <param name="emailaddress">A string represent a actual email address</param>
    /// <returns>A value indicates whether the address is a valid email address, true if the specified emailaddress is compliant with RFC822, otherwise return false.</returns>
    public static bool IsValidAddress(string emailaddress)
    {
        return addreg.IsMatch(emailaddress);
    }

    /// <summary>
    /// Initialize the regex expression
    /// </summary>
    private static void InitialRegex()
    {
        var qtext = @"[^" + Escape +
                    NonAscii +
                    CRLF + "\"]";
        var dtext = @"[^" + Escape +
                    NonAscii +
                    CRLF +
                    OpenBr +
                    CloseBr + "\"]";

        var quoted_pair = " " + Escape + " [^" + NonAscii + "] ";
        var ctext = @" [^" + Escape +
                    NonAscii +
                    CRLF + "()] ";

        // Nested quoted pairs
        var cnested = string.Empty;
        cnested += OpenParen;
        cnested += ctext + "*";
        cnested += "(?:" + quoted_pair + " " + ctext + "*)*";
        cnested += CloseParen;

        // A comment
        var comment = string.Empty;
        comment += OpenParen;
        comment += ctext + "*";
        comment += "(?:";
        comment += "(?: " + quoted_pair + " | " + cnested + ")";
        comment += ctext + "*";
        comment += ")*";
        comment += CloseParen;

        // x is optional whitespace/comments
        var x = string.Empty;
        x += "[" + Space + Tab + "]*";
        x += "(?: " + comment + " [" + Space + Tab + "]* )*";

        // An email address atom
        var atom_char = @"[^(" + Space + ")<>\\@,;:\\\"." + Escape + OpenBr +
                        CloseBr +
                        Ctrl +
                        NonAscii + "]";

        var atom = string.Empty;
        atom += atom_char + "+";
        atom += "(?!" + atom_char + ")";

        // Double quoted string, unrolled.
        var quoted_str = "(?'quotedstr'";
        quoted_str += "\\\"";
        quoted_str += qtext + " *";
        quoted_str += "(?: " + quoted_pair + qtext + " * )*";
        quoted_str += "\\\")";

        // A word is an atom or quoted string
        var word = string.Empty;
        word += "(?:";
        word += atom;
        word += "|";
        word += quoted_str;
        word += ")";

        // A domain-ref is just an atom
        var domain_ref = atom;

        // A domain-literal is like a quoted string, but [...] instead of "..."
        var domain_lit = string.Empty;
        domain_lit += OpenBr;
        domain_lit += "(?: " + dtext + " | " + quoted_pair + " )*";
        domain_lit += CloseBr;

        // A sub-domain is a domain-ref or a domain-literal
        var sub_domain = string.Empty;
        sub_domain += "(?:";
        sub_domain += domain_ref;
        sub_domain += "|";
        sub_domain += domain_lit;
        sub_domain += ")";
        sub_domain += x;

        // A domain is a list of subdomains separated by dots
        var domain = "(?'domain'";
        domain += sub_domain;
        domain += "(:?";
        domain += Period + " " + x + " " + sub_domain;
        domain += ")*)";

        // A route. A bunch of "@ domain" separated by commas, followed by a colon.
        var route = string.Empty;
        route += "\\@ " + x + " " + domain;
        route += "(?: , " + x + " \\@ " + x + " " + domain + ")*";
        route += ":";
        route += x;

        // A local-part is a bunch of 'word' separated by periods
        var local_part = "(?'localpart'";
        local_part += word + " " + x;
        local_part += "(?:";
        local_part += Period + " " + x + " " + word + " " + x;
        local_part += ")*)";

        // An addr-spec is local@domain
        var addr_spec = local_part + " \\@ " + x + " " + domain;

        // A route-addr is <route? addr-spec>
        var route_addr = string.Empty;
        route_addr += "< " + x;
        route_addr += "(?: " + route + " )?";
        route_addr += addr_spec;
        route_addr += ">";

        // A phrase
        var phrase_ctrl = @"\000-\010\012-\037";

        // Like atom-char, but without listing space, and uses phrase_ctrl. Since the class is negated, this matches the same as atom-char plus space and tab
        var phrase_char = "[^()<>\\@,;:\\\"." + Escape +
                          OpenBr +
                          CloseBr +
                          NonAscii +
                          phrase_ctrl + "]";

        var phrase = string.Empty;
        phrase += word;
        phrase += phrase_char;
        phrase += "(?:";
        phrase += "(?: " + comment + " | " + quoted_str + " )";
        phrase += phrase_char + " *";
        phrase += ")*";

        // A mailbox is an addr_spec or a phrase/route_addr
        var mailbox = string.Empty;
        mailbox += x;
        mailbox += "(?'mailbox'";
        mailbox += addr_spec;
        mailbox += "|";
        mailbox += phrase + " " + route_addr;
        mailbox += ")";

        addreg = new Regex(mailbox, RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace);
    }
}