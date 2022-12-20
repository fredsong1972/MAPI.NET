using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace MAPI.NET
{
    /// <summary>
    /// Rtf Html Converter
    /// </summary>
    public class RtfHtmlConverter
    {
        // Default color table used by VS's IDE.
        static string[] m_colorTable = new string[] 
            {
               // rrGGbb
                "#000000", // default, starts at index 0
                "#000000", // real color table starts at index 1
                "#0000FF",
                "#00ffFF",
                "#00FF00",
                "#FF00FF",
                "#FF0000",
                "#FFFF00",
                "#FFffFF",
                "#000080",
                "#008080",
                "#008000",
                "#800080",
                "#800000",
                "#808000",
                "#808080",
                "#c0c0c0"
            };


        /// <summary>
        /// Escape HTML chars
        /// </summary>
        /// <param name="st">string</param>
        /// <returns>escaped string</returns>
        static string Escape(string st)
        {
            st = st.Replace("&", "&amp");
            st = st.Replace("&", "&amp;");
            st = st.Replace("<", "&lt;");
            st = st.Replace(">", "&gt;");
            return st;
        }

        /// <summary>
        /// check if html in rtf
        /// </summary>
        /// <param name="rtf">rtf string</param>
        /// <returns>true, if html in rtf; otherwise, no html.</returns>
        public static bool IsHTMLinRtf(string rtf)
        {
            // We look for the words "\fromhtml" somewhere in the file.
            // If the rtf encodes text rather than html, then instead
            // it will only find "\fromtext".
            return rtf.Contains(@"\fromhtml");
        }

        /// <summary>
        /// convert body text to html
        /// </summary>
        /// <param name="rtf">rtf string</param>
        /// <returns>html string</returns>
        public static string HtmlFromRtf(string rtf)
        {
            StringBuilder s = new StringBuilder();
            int len = rtf.Length;
            int pos = rtf.IndexOf(@"{\*\htmltag");
            int ignore_tag = -1;

            while (pos < len)
            {
                char c = rtf[pos];
                if (c == '{') ++pos;
                else if (c == '}') ++pos;
                else if (c == '\r' || c == '\n') ++pos;
                else if (c == '\\')
                {
                    ++pos;
                    if (rtf[pos] == '{') { s.Append('{'); ++pos; }
                    else if (rtf[pos] == '}') { s.Append('}'); ++pos; }
                    else if (rtf[pos] == '\'') { ++pos; int hex = Convert.ToInt32(rtf.Substring(pos, 2), 16); s.Append((char)hex); pos += 2; }
                    else if (Match(rtf, @"*\htmltag", ref pos) && pos < len)
                    {
                        int tag = ReadInt(rtf, ref pos);
                        if (tag == ignore_tag)
                            if (!SkipTo(rtf, '}', ref pos))
                                break; // abort if '}' not found
                    }
                    else if (Match(rtf, @"*\mhtmltag", ref pos)) { ignore_tag = ReadInt(rtf, ref pos); }
                    else if (Match(rtf, "par", ref pos)) { s.Append("\r\n"); }
                    else if (Match(rtf, "tab", ref pos)) { s.Append("\t"); }
                    else if (Match(rtf, "li", ref pos)) { ReadInt(rtf, ref pos); }
                    else if (Match(rtf, "fi-", ref pos)) { ReadInt(rtf, ref pos); }
                    else if (Match(rtf, "pntext", ref pos)) { SkipTo(rtf, '}', ref pos); }
                    else if (Match(rtf, "htmlrtf", ref pos) && pos < len)
                    {
                        pos = rtf.IndexOf(@"\htmlrtf0", pos);
                        if (pos < 0)
                            break;
                        pos += 9;
                    }
                    else { s.Append('\\'); s.Append(c); ++pos; }
                }
                else { s.Append(c); ++pos; }
            }

            if (pos < 0)
                return null;

            return s.ToString();
        }

        /// <summary>
        /// Convert plain text to html
        /// </summary>
        /// <param name="text">plain text</param>
        /// <returns>html string</returns>
        public static string HtmlFromPlainText(string text)
        {
            string html = "<html><pre>";
            html += Escape(text);
            html = html.Replace("\r\n", "<br>");
            html += "</pre></html>";
            return html;
        }

        /// <summary>
        /// Convert html to plain text
        /// </summary>
        /// <param name="source">html</param>
        /// <returns>plain text</returns>
        public static string PlainTextFromHtml(string source)
        {
            try
            {
                string result;

                // Remove HTML Development formatting
                // Replace line breaks with space
                // because browsers inserts space
                result = source.Replace("\r", " ");
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\n", " ");
                // Remove step-formatting
                result = result.Replace("\t", string.Empty);
                // Remove repeating spaces because browsers ignore them
                result = System.Text.RegularExpressions.Regex.Replace(result,
                                                                      @"( )+", " ");

                // Remove the header (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*head([^>])*>", "<head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*head( )*>)", "</head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<head>).*(</head>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all scripts (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*script([^>])*>", "<script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*script( )*>)", "</script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                //result = System.Text.RegularExpressions.Regex.Replace(result,
                //         @"(<script>)([^(<script>\.</script>)])*(</script>)",
                //         string.Empty,
                //         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<script>).*(</script>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all styles (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*style([^>])*>", "<style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*style( )*>)", "</style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<style>).*(</style>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert tabs in spaces of <td> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*td([^>])*>", "\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line breaks in places of <BR> and <LI> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*br( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*li( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line paragraphs (double line breaks) in place
                // if <P>, <DIV> and <TR> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*div([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*tr([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*p([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Remove remaining tags like <a>, links, images,
                // comments etc - anything that's enclosed inside < >
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<[^>]*>", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // replace special characters:
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @" ", " ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&bull;", " * ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lsaquo;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&rsaquo;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&trade;", "(tm)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&frasl;", "/",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lt;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&gt;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&copy;", "(c)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&reg;", "(r)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove all others. More can be added, see
                // http://hotwired.lycos.com/webmonkey/reference/special_characters/
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&(.{2,6});", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // for testing
                //System.Text.RegularExpressions.Regex.Replace(result,
                //       this.txtRegex.Text,string.Empty,
                //       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // make line breaking consistent
                result = result.Replace("\n", "\r");

                // Remove extra line breaks and tabs:
                // replace over 2 breaks with 2 and over 4 tabs with 4.
                // Prepare first to remove any whitespaces in between
                // the escaped characters and remove redundant tabs in between line breaks
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\t)", "\t\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\r)", "\t\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\t)", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove redundant tabs
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove multiple tabs following a line break with just one tab
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Initial replacement target string for line breaks
                string breaks = "\r\r\r";
                // Initial replacement target string for tabs
                string tabs = "\t\t\t\t\t";
                for (int index = 0; index < result.Length; index++)
                {
                    result = result.Replace(breaks, "\r\r");
                    result = result.Replace(tabs, "\t\t\t\t");
                    breaks = breaks + "\r";
                    tabs = tabs + "\t";
                }

                // That's it.
                return result;
            }
            catch
            {
                return source;
            }
        }

        /// <summary>
        /// Convert RTF to html
        /// </summary>
        /// <param name="rtf">rtf</param>
        /// <returns>html string</returns>
        public static string HtmlFromRTF2(string rtf)
        {
            StringWriter tw = new StringWriter();

            tw.Write("<html><pre>");
            tw.Write("<span color=black>");
            int i1 = rtf.IndexOf(@"{\colortbl");
            if (i1 <= 0) throw new ArgumentException("Bad input RTF.");
            int i2 = rtf.IndexOf(";}", i1);
            if (i2 <= 0) throw new ArgumentException("Bad input RTF.");
            rtf = rtf.Substring(i2 + 2, rtf.Length - (i2 + 2) - 1);
            rtf = rtf.Replace("\n", " ");
            rtf = rtf.Replace("\r", " ");
            // Example: \fs20 \cf2 using\cf0  System;
            // root --> ('text' '\' ('control word' | 'escaped char'))+
            // 'control word'  --> (alpha)+ (numeric*) space?
            // 'escaped char' = 'x'. Some characters \, {, } are escaped: '\x' --> 'x'
            // @todo - handle embedded groups (begin with '{')

            int idx = 0;
            while (idx < rtf.Length)
            {
                // Get any text up to a '\'. 
                Regex r1 = new Regex(@"(.*?)\\", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                Match m = r1.Match(rtf, idx);
                if (m.Length == 0) break;

                // text will be empty if we have adjacent control words
                string stText = m.Groups[1].ToString();
                tw.Write(Escape(stText));
                idx += m.Length;

                // check for RTF escape characters. According to the spec, these are the only escaped chars.
                char chNext = rtf[idx];
                if (chNext == '{' || chNext == '}' || chNext == '\\')
                {
                    // Escaped char
                    tw.Write(chNext);
                    idx++;
                    continue;
                }

                // Must be a control char. @todo- delimeter includes more than just space, right?
                Regex r2 = new Regex(@"([\{a-z]+)([0-9]*) ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                m = r2.Match(rtf, idx);
                string stCtrlWord = m.Groups[1].ToString();
                string stCtrlParam = m.Groups[2].ToString();

                if (stCtrlWord == "cf")
                {
                    // Set font color.
                    int iColor = Int32.Parse(stCtrlParam);
                    tw.Write("</span>"); // close previous span, and start a new one for the given color.                    
                    tw.Write("<span style=\"color: " + m_colorTable[iColor] + "\">");
                }
                else if (stCtrlWord == "fs")
                {
                    // Sets font size. ignore
                }
                else if (stCtrlWord == "par")
                {
                    // This is a newline. ignore
                    // @todo- I think the only reason we can ignore this is because the \par in our input are always followed by
                    // a '\r\n' and we're accidentally writing that.
                }
                else
                {
                    //throw new ArgumentException("Unrecognized control word '" + stCtrlWord + stCtrlParam + "'after:" + stText);
                }
                idx += m.Length;
            }
            tw.Write(Escape(rtf.Substring(idx))); // rest of string

            tw.Write("</pre></html>");
            return tw.ToString();
        }

        private static bool Match(string str, string match, ref int pos)
        {
            if (str.Length >= match.Length + pos
                && str.Substring(pos, match.Length) == match)
            {
                pos += match.Length;
                return true;
            }
            else
                return false;
        }

        private static int ReadInt(string str, ref int pos)
        {
            int i = 0;
            while (str[pos] >= '0' && str[pos] <= '9')
            {
                i = i * 10 + (str[pos] - '0');
                ++pos;
            }
            if (str[pos] == ' ') ++pos;
            return i;
        }

        private static bool SkipTo(string str, char c, ref int pos)
        {
            pos = str.IndexOf('}', pos);
            if (pos < 0) return false;
            ++pos; return true;
        }

    }
}

