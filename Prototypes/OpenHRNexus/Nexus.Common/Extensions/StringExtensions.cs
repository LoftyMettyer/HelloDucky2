using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

public static class StringExtension
{
    static readonly Regex re = new Regex(@"\{([^\}]+)\}", RegexOptions.Compiled);
    public static string FormatPlaceholder(this string str, Dictionary<string, object> fields)
    {
        if (fields == null)
            return str;

        if (fields.Count == 0)
            return str;

        return re.Replace(str, delegate (Match match)
        {
            if (fields.ContainsKey(match.Groups[1].Value))
                return fields[match.Groups[1].Value].ToString();
            else
                return match.Value;
        });

    }
}