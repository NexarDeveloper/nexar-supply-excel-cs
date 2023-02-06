using System;
using System.Globalization;
using Nexar.Supply.Api;
using Nexar.Supply.SupplySchema;

namespace ExtensionMethods
{
    public static class Extensions
    {
        /// <summary>
        /// Gets the url of the datasheet; returns first option if available
        /// </summary>
        /// <param name="part">The part as returned by the search</param>
        /// <param name="excludeDatasheets">Datasheets are unauthorized for the querying client</param>
        /// <returns>The url for the 'best' datasheet</returns>
        public static string GetDatasheetUrl(this Part part, bool excludeDatasheets)
        {
            if (excludeDatasheets)
                return "";

            if (part != null && part.BestDatasheet != null)
                return part.BestDatasheet.Url;

            return "ERROR: Datasheet url not found. Please try expanding your search";
        }

        /// <summary>
        /// Find the best price given the price break
        /// </summary>
        /// <param name="offer">The offer to search within</param>
        /// <param name="currency">The 3 character currency code</param>
        /// <param name="qty">The quantity to search for (i.e. QTY required)</param>
        /// <returns>The minimum price available for the specified QTY</returns>
        public static double MinPrice(this Offer offer, string currency, int qty)
        {
            // Force format optional arguments
            if (currency == string.Empty) currency = "USD";
            if (qty == 0) qty = 1;

            double minprice = double.MaxValue;

            try
            {
                string minpricestr = ApiV4.OffersMinPrice(currency, offer, qty, true);
                if (!string.IsNullOrEmpty(minpricestr))
                    minprice = Convert.ToDouble(minpricestr, CultureInfo.CurrentCulture);
            }
            catch (FormatException) { /* Do nothing */ }
            catch (OverflowException) { /* Do nothing */ }

            return minprice;
        }

        /// <summary>
        /// Return an ISO 8601 compliant string in UTC representing the given DateTimeOffset.
        /// For example: 2023-02-05T14:29:20Z. The formatted string is also RFC3339 compliant.
        /// </summary>
        /// <param name="dateTimeOffset">The input DateTimeOffset</param>
        /// <returns>An ISO 8601 compliant string representation in UTC of the input DateTimeOffset</returns>
        public static string ToUtcIso8601String(this DateTimeOffset dateTimeOffset)
        {
            return dateTimeOffset.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ssZ", DateTimeFormatInfo.InvariantInfo);
        }
    }
}
