using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Net;
using Newtonsoft.Json;
using Nexar.Supply.Query;
using Nexar.Supply.SupplySchema;
using RestSharp;

namespace Nexar.Supply.Api
{
    public static class ApiV4
    {
        #region Variables
        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region Classes
        /// <summary>
        /// Encapsulates search results
        /// </summary>
        public class SearchResponse
        {
            /// <summary>
            /// Gets the data returned by the search query
            /// </summary>
            public object Data { get; internal set; }

            /// <summary>
            /// Gets the error message from search query (if available)
            /// </summary>
            public string ErrorMessage { get; internal set; }
        }
        #endregion

        #region Constants
        /// <summary>
        /// The standard limit per Nexar Supply API request (set to 20 as indicated by the underlying Octopart service)
        /// Note: Returning 20 results in up to 20 MPNs. In testing, this means data unrelated to the intended part is being pulled. Commenting out and setting to 1 pending future side effects.
        /// </summary>
        /// public const int RECORD_LIMIT_PER_QUERY = 20;
	    public const int RECORD_LIMIT_PER_QUERY = 1;

        /// <summary>
        /// The start limit for a Nexar Supply API request (guided by the underlying Octopart service)
        /// </summary>
        public const int RECORD_START_MAX = 80;

        private const string NEXAR_BASE_URL = "https://api.nexar.com";
        private const string NEXAR_GRAPHQL = "/graphql";
        private const string NEXAR_VERSION = "0.5";

        #endregion

        #region NexarUtils

#if _
        private static string GetSearchMpnQuery(string mpn)
        {
            return $@"query {{supSearchMpn(q: \""{mpn}\"") {{ results {{ part {{ category {{ parentId id name path }} mpn manufacturer {{ name }} shortDescription descriptions {{ text creditString }} specs {{ attribute {{ name shortname }} displayValue }} }} }} }} }}";
        }
#endif
        
        private static string GetMultiMatchQuery(bool includeDatasheets, bool includeLeadTime)
        {
            string datasheet = includeDatasheets ? "bestDatasheet { url } " : string.Empty;
            string factoryLeadDays = includeLeadTime ? "factoryLeadDays " : string.Empty;

            return "query($queries: [SupPartMatchQuery!]!) {supMultiMatch (queries: $queries) { reference error hits parts { v3uid mpn shortDescription manufacturer { id name homepageUrl } " + datasheet + "octopartUrl sellers { offers { id sku " + factoryLeadDays + "factoryPackQuantity inventoryLevel onOrderQuantity orderMultiple multipackQuantity packaging moq clickUrl updated prices { currency quantity price } } company { id name homepageUrl } isAuthorized } } } }";
        }

        private static string GetResponseErrorMessage(IRestResponse res)
        {
            string contentJson = res.Content;
            var supplyResult = SupplyResult.FromJson(contentJson);

            if (supplyResult?.Errors.Count > 0)
                return supplyResult.Errors.First().Message;
            
            return "Server did not return OK (" + res.ErrorMessage + ")";
        }

        /// <summary>
        /// Returns the version of this Nexar Supply API client
        /// </summary>
        public static string GetVersion()
        {
            return NEXAR_VERSION;
        }

#endregion

#region Methods-Search
        /// <summary>
        /// Execute a part/match endpoint search
        /// </summary>
        /// <param name="pnList">Manufacturing Part Number based search string</param>
        /// <param name="nexarToken">Nexar Supply client token (obtained via Client Id/Secret)</param>
        /// <param name="httpTimeout">The desired timeout for the http request (in ms)</param>
        /// <returns>
        /// Returns a list of parts that were found from the provided search string.
        /// NULL indicates that no parts were found, or there was an error with the request
        /// </returns>
        /// <notes>
        /// Server timeout is defaulted to 5000ms
        /// </notes>
        public static SearchResponse PartsMatch(List<PartsMatchQuery> pnList, string nexarToken, bool includeDatasheets, bool includeLeadTime, int httpTimeout = 5000)
        {
            if ((pnList == null) || (pnList.Count == 0))
                return null;

            var query = new List<Dictionary<string, object>>();
            foreach (PartsMatchQuery pn in pnList)
            {
                query.Add(
                    new Dictionary<string, object>()
                    {
                        { "mpn", pn.q },
                        { "limit", pn.limit },
                        { "start", pn.start },
                        { "reference", pn.q }
                    }
                  );
            }
            string queryString = JsonConvert.SerializeObject(query);
#if _
            string queryString = "";
#endif

            SearchResponse ret = new SearchResponse();

            if (!string.IsNullOrEmpty(nexarToken))
            {
                var client2 = new RestClient(NEXAR_BASE_URL);

                // TODO: The following creates returns a searchMpn search - unused here, but something to explore...
                // var searchQuery = GetSearchMpnQuery(pnList[0].q);
                // string jsonBody = $@"{{""query"": ""{searchQuery}""}}";

                //  For this parts match query, we're gunning for something like this...
                //  https://octopart.com/api/v4/rest/parts/match?apikey=YOUR_API_KEY&include[]=datasheets&queries=[%7B%22mpn%22:%22INA225*%22,%22limit%22:4%7D]&pretty_print=true

                string matchQuery = GetMultiMatchQuery(includeDatasheets, includeLeadTime);
                string variables = "\"queries\": " + queryString;
                string jsonBody = $@"{{ ""query"": ""{matchQuery}"", ""variables"": {{ {variables} }} }}";

                var req2 = new RestRequest(NEXAR_GRAPHQL, Method.POST) { RequestFormat = DataFormat.Json }
                    .AddParameter("application/json", jsonBody, ParameterType.RequestBody)
                    .AddHeader("Content-Type", "application/json")
                    .AddHeader("X-Nexar-Client", "NexarSupply-AddIn")
                    .AddHeader("X-Nexar-Client-Version", NEXAR_VERSION)
                    .AddHeader("Authorization", "Bearer " + nexarToken);

                req2.Timeout = httpTimeout;

                var res2 = client2.Execute(req2);

                if (res2.StatusCode != HttpStatusCode.OK)
                {
                    ret.ErrorMessage = GetResponseErrorMessage(res2);
                    Log.Error(string.Format("{0}", ret.ErrorMessage));
                    return ret;
                }

                string contentJson = res2.Content;
                var supplyResult = SupplyResult.FromJson(contentJson);

                ret.Data = supplyResult.Data;

                if (!string.IsNullOrEmpty(res2.ErrorMessage))
                {
                    ret.ErrorMessage = res2.ErrorMessage;
                }
                else if (supplyResult.Errors?.Count > 0)
                {
                    ret.ErrorMessage = supplyResult.Errors[0].Message;
                }

                return ret;
            }
            else 
            {
                ret.ErrorMessage = "Please login with a call to NEXAR_SUPPLY_LOGIN ";
                Log.Error(string.Format("Unexpected Error (resp == null) '{0}'", ret.ErrorMessage));
                return ret;
            }
        }
    
#endregion

#region Methods-Helper
        /// <summary>
        /// Find the best price given the price break
        /// </summary>
        /// <param name="desiredCurrency">The 3 character currency code</param>
        /// <param name="offer">The offer to search within</param>
        /// <param name="qty">The quantity to search for (i.e. QTY required)</param>
        /// <param name="ignoreMoq">Indicates if the MOQ should be considered when looking up pricing</param>
        /// <returns>The minimum price available for the specified QTY, in Current-Culture string format</returns>
        public static string OffersMinPrice(string desiredCurrency, Offer offer, int qty, bool ignoreMoq)
        {
            if (string.IsNullOrEmpty(desiredCurrency) || (offer == null))
                return string.Empty;

            double priceMin = double.MaxValue;
            List<Price> currencyPrices = offer.Prices.Where(p => p.Currency.Contains(desiredCurrency)).ToList();

            if (currencyPrices != null)
            {
                foreach (Price price in currencyPrices)
                {
                    try
                    {
                        // TODO: The Octopart Add-in had something about a Mouser special case...
                        if ((int)price.Quantity <= qty)
                            if (price.PricePrice < priceMin)
                                priceMin = price.PricePrice;
                    }
                    catch (FormatException) { /* Do nothing */ }
                    catch (OverflowException) { /* Do nothing */ }
                }
            }

            if (priceMin == double.MaxValue)
                return string.Empty;
            else
                return priceMin.ToString("F5", CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Find the preferred currency, or whatever else
        /// </summary>
        /// <param name="desiredCurrency">The desired currency to get the price breaks from</param>
        /// <param name="offer">The offer to search within</param>
        /// <returns>Currency string</returns>
        public static string FindPreferredCurrency(string desiredCurrency, Offer offer)
        {
            string ret = string.Empty;
            foreach (Price price in offer.Prices)
            {
                string currency = price.Currency;
                if (!string.IsNullOrEmpty(currency))
                {
                    if (currency == desiredCurrency)
                    {
                        // We found at least one price with the specified currency, so return it.
                        return desiredCurrency;
                    }
                    else
                    {
                        // Well, we didn't find what we were looking for, but at least give something.
                        ret = currency;
                    }
                }
            }

            return ret;
        }
#endregion
    }
}
