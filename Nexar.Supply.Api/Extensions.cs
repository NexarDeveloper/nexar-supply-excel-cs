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
        /// <returns>The url for the 'best' datasheet</returns>
        public static string GetDatasheetUrl(this Part part)
        {
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
    }
}
