using System.Text.RegularExpressions;

namespace NexarSupplyXll
{
    public static class Extensions
    {
        /// <summary>
        /// Removes all non-alphanumeric characters from the provided string (including
        /// spaces, dashes, underscores, etc), EXCEPT for the wildcard character '*'
        /// </summary>
        /// <param name="str">String to sanitize</param>
        /// <returns>Sanitized string</returns>
        public static string Sanitize(this string str)
        {
            return Regex.Replace(str.ToLower(), @"[^A-Za-z0-9*]+", "");
        }
    }
}
