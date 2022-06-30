using ExcelDna.Integration;
using ExtensionMethods;
using IdentityModel.Client;
using Nexar.Supply.SupplySchema;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace NexarSupplyXll
{
    public class NexarSupplyAddIn : IExcelAddIn
    {
        #region Types
        /// <summary>
        /// Whether to consider authorized sellers when filtering search results
        /// </summary>
        public enum AuthorizedSeller
        {
            Any,
            Yes,
            No
        };
        #endregion

        #region Variables
        /// <summary>
        /// A collection of Options that configure the Nexar Supply Add-In
        /// </summary>
        private static readonly Dictionary<string, dynamic> Options = new Dictionary<string, dynamic>()
        {
            {"log", false}
        };

        /// <summary>
        /// Handles all queries to the Nexar Supply API, and caches the results for the Excel session
        /// </summary>
        public static readonly NexarQueryManager QueryManager = new NexarQueryManager();

        /// <summary>
        /// This 'refresh' hack is used to trick excel to refresh the information from the QueryManager
        /// </summary>
        private static string _refreshhack = string.Empty;

        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion

        #region Constructors
        /// <summary>
        /// This function is called when the XLL is installed into an Excel spreadsheet
        /// </summary>
        public void AutoOpen()
        {
            try
            {
                ExcelIntegration.RegisterUnhandledExceptionHandler(
                    ex => "!!! EXCEPTION: " + ex.ToString());
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
            }
        }

        /// <summary>
        /// Gracefully detaches the XLL
        /// </summary>
        public void AutoClose()
        { }
        #endregion

        #region ExcelUdfs
        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the Nexar Supply Details Url")]
        public static object NEXAR_SUPPLY_DETAIL_URL(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part (optional)", Name = "Manufacturer")] string manuf = "")
        {
            // Check to see if a cached version is available (only checks non-error'd queries)
            Part part = GetManuf(mpn_or_sku, manuf);
            if (part != null)
                return part.OctopartUrl?.AbsoluteUri ?? string.Empty;

            // Excel's recalculation engine is based on INPUTS. The main function will be called if:
            // - Inputs are changed
            // - Input cells location are changed (i.e., moving the cell, or deleting a row that impacts this cell)
            // However, the async function will ONLY be run if the inputs are DIFFERENT.
            //   The impact of this is that if last time the function was run it returned an error that was unrelated to the inputs 
            //   (i.e., invalid ApiKey, network was down, etc), then the function would not run again.
            // To fix that issue, whitespace padding is added to the mpn_or_sku. This whitespace is removed anyway, so it has no 
            // real impact other than to generate a refresh.
            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
#if !TEST
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DETAIL_URL", new object[] { mpn_or_sku, manuf }, delegate
            {
#endif
            try
            {
                part = SearchAndWaitPart(mpn_or_sku, manuf);
                if (part == null)
                {
                    string err = QueryManager.GetLastError(mpn_or_sku);
                    if (string.IsNullOrEmpty(err))
                        err = "Query did not provide a result. Please widen your search criteria.";

                    return "ERROR: " + err;
                }

                return part.OctopartUrl?.AbsoluteUri ?? string.Empty;
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
                return "ERROR: " + NexarQueryManager.FATAL_ERROR;
            }
#if !TEST
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                // Still processing...
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (part == null))
            {
                // Regenerate the hack value if an error was received
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }
            
            // Done processing!
            return asyncResult;
#endif
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the Nexar Supply Datasheet Url", HelpTopic = "NexarSupplyAddIn.chm!1002")]
        public static object NEXAR_SUPPLY_DATASHEET_URL(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "")
        {
            Part part = GetManuf(mpn_or_sku, manuf);
            if (part != null)
            {
                // ---- BEGIN Function Specific Information ----
                return part.GetDatasheetUrl(QueryManager.IncludeDatasheets);
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DATASHEET_URL", new object[] { mpn_or_sku, manuf }, delegate
            {
                try
                {
                    part = SearchAndWaitPart(mpn_or_sku, manuf);
                    if (part == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return part.GetDatasheetUrl(QueryManager.IncludeDatasheets);
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (part == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the short description of the part from Nexar Supply")]
        public static object NEXAR_SUPPLY_SHORT_DESCRIPTION(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "")
        {
            Part part = GetManuf(mpn_or_sku, manuf);
            if (part != null)
            {
                // ---- BEGIN Function Specific Information ----
                return part.ShortDescription;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DATASHEET_URL", new object[] { mpn_or_sku, manuf }, delegate
            {
                try
                {
                    part = SearchAndWaitPart(mpn_or_sku, manuf);
                    if (part == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return part.ShortDescription;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (part == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor price from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1003")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_PRICE(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null,
            [ExcelArgument(Description = "Quantity for lookup (optional, default = 1)", Name = "Quantity")] int qty = 1,
            [ExcelArgument(Description = "Currency for lookup (optional, default = USD). Standard currency codes apply (http://en.wikipedia.org/wiki/ISO_4217)", Name = "Currency")] string currency = "USD",
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            List<Offer> offers = GetOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                double minprice = offers.Min(offer => offer.MinPrice(currency, qty));
                if (minprice < double.MaxValue)
                    return minprice;
                else
                    return "ERROR: Query did not provide a result. Please widen your search criteria.";
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
#if !TEST
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_PRICE", new object[] { mpn_or_sku, manuf, distributors, qty, currency }, delegate
            {
#endif
            try
            {
                offers = SearchAndWaitOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
                if ((offers == null) || (offers.Count == 0))
                {
                    string err = QueryManager.GetLastError(mpn_or_sku);
                    if (string.IsNullOrEmpty(err))
                        err = "Query did not provide a result. Please widen your search criteria.";

                    return "ERROR: " + err;
                }

                // ---- BEGIN Function Specific Information ----
                double minprice = offers.Min(offer => offer.MinPrice(currency, qty));
                if (minprice < double.MaxValue)
                    return minprice;
                else
                    return "ERROR: Query did not provide a result. Please widen your search criteria.";
                // ---- END Function Specific Information ----
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
                return "ERROR: " + NexarQueryManager.FATAL_ERROR;
            }
#if !TEST
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
#endif
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the average price from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1004")]
        public static object NEXAR_SUPPLY_AVERAGE_PRICE(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Quantity for lookup (optional, default = 1)", Name = "Quantity")] int qty = 1,
            [ExcelArgument(Description = "Currency for lookup (optional, default = USD). Standard currency codes apply (http://en.wikipedia.org/wiki/ISO_4217)", Name = "Currency")] string currency = "USD",
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller) Parse(authorized);
            List<Offer> offers = GetOffers(mpn_or_sku, manuf, auth);
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                offers = offers.Where(offer => offer.MinPrice(currency, qty) < double.MaxValue).ToList();
                if ((offers != null) && (offers.Count > 0))
                {
                    double price = offers.Average(offer => offer.MinPrice(currency, qty));
                    if (price < double.MaxValue)
                        return price;
                    else
                        return "ERROR: Query did not provide a result. Please widen your search criteria.";
                }
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
#if !TEST
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_AVERAGE_PRICE", new object[] { mpn_or_sku, manuf, qty, currency }, delegate
            {
#endif
            try
            {
                offers = SearchAndWaitOffers(mpn_or_sku, manuf, auth);
                if ((offers == null) || (offers.Count == 0))
                {
                    string err = QueryManager.GetLastError(mpn_or_sku);
                    if (string.IsNullOrEmpty(err))
                        err = "Query did not provide a result. Please widen your search criteria.";

                    return "ERROR: " + err;
                }

                // ---- BEGIN Function Specific Information ----
                offers = offers.Where(offer => offer.MinPrice(currency, qty) < double.MaxValue).ToList();
                if ((offers != null) && (offers.Count > 0))
                {
                    double price = offers.Average(offer => offer.MinPrice(currency, qty));
                    if (price < double.MaxValue)
                        return price;
                }

                return "ERROR: Query did not provide a result. Please widen your search criteria.";
                // ---- END Function Specific Information ----
            }
            catch (Exception ex)
            {
                log.Fatal(ex.ToString());
                return "ERROR: " + NexarQueryManager.FATAL_ERROR;
            }
#if !TEST
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
#endif
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor stock quantity from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1005")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_STOCK(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null,
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            List<Offer> offers = GetOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
            if (offers != null && offers.Count > 0)
            {
                // ---- BEGIN Function Specific Information ----
                return GetOffersStock(offers);
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_STOCK", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return GetOffersStock(offers);
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor MOQ from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1006")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_MOQ(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null,
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            List<Offer> offers = GetOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                offers = offers.Where(offer => offer.Moq > 0.0).ToList();
                if ((offers != null) && (offers.Count > 0))
                {
                    return offers.Min(offer => offer.Moq);
                }
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_MOQ", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    offers = offers.Where(offer => offer.Moq > 0.0).ToList();
                    if ((offers != null) && (offers.Count > 0))
                    {
                        return offers.Min(offer => offer.Moq);
                    }
                    else
                    {
                        string err = "Query did not provide a result. Please widen your search criteria.";
                        return "ERROR: " + err;
                    }
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor order multiple from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1007")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_ORDER_MULTIPLE(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null,
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            List<Offer> offers = GetOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
            if (offers != null && offers.Count > 0)
            {
                // ---- BEGIN Function Specific Information ----
                return offers.Min(offer => offer.OrderMultiple).ToString();
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_ORDER_MULTIPLE", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offers.Min(offer => offer.OrderMultiple).ToString();
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor lead time from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1008")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_LEAD_TIME(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributors for lookup (optional)", Name = "Distributor(s)")] Object[] distributors = null,
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            List<Offer> offers = GetOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
            if ((offers != null) && (offers.Count > 0))
            {
                // ---- BEGIN Function Specific Information ----
                return offers.Min(offer => offer.FactoryLeadDays).ToString();
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_LEAD_TIME", new object[] { mpn_or_sku, manuf, distributors }, delegate
            {
                try
                {
                    offers = SearchAndWaitOffers(mpn_or_sku, manuf, auth, GetDistributors(distributors));
                    if ((offers == null) || (offers.Count == 0))
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offers.Min(offer => offer.FactoryLeadDays).ToString();
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offers == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor packaging style from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1009")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_PACKAGING(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributor for lookup (optional)", Name = "Distributor")] string distributor = "",
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            Offer offer = GetOffer(mpn_or_sku, manuf, auth, distributor);
            if (offer != null)
            {
                // ---- BEGIN Function Specific Information ----
                return offer.Packaging ?? string.Empty;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_PACKAGING", new object[] { mpn_or_sku, manuf, distributor }, delegate
            {
                try
                {
                    offer = SearchAndWaitOffer(mpn_or_sku, manuf, auth, distributor);
                    if (offer == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offer.Packaging ?? string.Empty;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offer == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor url from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1010")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_URL(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributor for lookup (optional)", Name = "Distributor")] string distributor = "",
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            Offer offer = GetOffer(mpn_or_sku, manuf, auth, distributor);
            if (offer != null)
            {
                // ---- BEGIN Function Specific Information ----
                return offer.ClickUrl?.AbsoluteUri ?? string.Empty;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_URL", new object[] { mpn_or_sku, manuf, distributor }, delegate
            {
                try
                {
                    offer = SearchAndWaitOffer(mpn_or_sku, manuf, auth, distributor);
                    if (offer == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offer.ClickUrl?.AbsoluteUri ?? string.Empty;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offer == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Gets the distributor SKU from Nexar Supply", HelpTopic = "NexarSupplyAddIn.chm!1011")]
        public static object NEXAR_SUPPLY_DISTRIBUTOR_SKU(
            [ExcelArgument(Description = "Part Number Lookup", Name = "MPN or SKU")] string mpn_or_sku,
            [ExcelArgument(Description = "Manufacturer of the part to query (optional)", Name = "Manufacturer")] string manuf = "",
            [ExcelArgument(Description = "Distributor for lookup (optional)", Name = "Distributor")] string distributor = "",
            [ExcelArgument(Description = "Only authorized sellers (optional, default = Any)", Name = "Authorized")] string authorized = "Any")
        {
            AuthorizedSeller auth = (AuthorizedSeller)Parse(authorized);
            Offer offer = GetOffer(mpn_or_sku, manuf, auth, distributor);
            if (offer != null)
            {
                // ---- BEGIN Function Specific Information ----
                return offer.Sku;
                // ---- END Function Specific Information ----
            }

            mpn_or_sku = mpn_or_sku.PadRight(mpn_or_sku.Length + _refreshhack.Length);
            object asyncResult = ExcelAsyncUtil.Run("NEXAR_SUPPLY_DISTRIBUTOR_SKU", new object[] { mpn_or_sku, manuf, distributor }, delegate
            {
                try
                {
                    offer = SearchAndWaitOffer(mpn_or_sku, manuf, auth, distributor);
                    if (offer == null)
                    {
                        string err = QueryManager.GetLastError(mpn_or_sku);
                        if (string.IsNullOrEmpty(err))
                            err = "Query did not provide a result. Please widen your search criteria.";

                        return "ERROR: " + err;
                    }

                    // ---- BEGIN Function Specific Information ----
                    return offer.Sku;
                    // ---- END Function Specific Information ----
                }
                catch (Exception ex)
                {
                    log.Fatal(ex.ToString());
                    return "ERROR: " + NexarQueryManager.FATAL_ERROR;
                }
            });

            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
            {
                return NexarQueryManager.PROCESSING;
            }
            else if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)) || (offer == null))
            {
                _refreshhack = string.Empty.PadRight(new Random().Next(0, 100));
            }

            return asyncResult;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Sets the Nexar Application's client Id and Secret to allow Nexar Supply queries", IsVolatile = true)]
        public static object NEXAR_SUPPLY_LOGIN(
            [ExcelArgument(Description = "Client Id", Name = "Client Id")] string clientId,
            [ExcelArgument(Description = "Client Secret", Name = "Client Secret")] string clientSecret,
            [ExcelArgument(Description = "Include Datasheets (optional)", Name = "Include Datasheets")] string datasheets = "",
            [ExcelArgument(Description = "Include Lead Time (optional)", Name = "Include Lead Time")] string leadTime = "")
        {
            bool includeDatasheets = true;
            if (!string.IsNullOrEmpty(datasheets))
                bool.TryParse(datasheets, out includeDatasheets);
            
            bool includeLeadTime = true;
            if (!string.IsNullOrEmpty(leadTime))
                bool.TryParse(leadTime, out includeLeadTime);
            
            bool changed =
                QueryManager.NexarClientId != clientId ||
                QueryManager.NexarClientSecret != clientSecret ||
                QueryManager.IncludeDatasheets != includeDatasheets ||
                QueryManager.IncludeLeadTime != includeLeadTime;
            
            bool renew = QueryManager.NexarTokenRenewing;
            if (renew)
                QueryManager.NexarTokenRenewing = false;

            if (changed || renew)
            {
                QueryManager.NexarClientId = clientId;
                QueryManager.NexarClientSecret = clientSecret;
                QueryManager.IncludeDatasheets = includeDatasheets;
                QueryManager.IncludeLeadTime = includeLeadTime;

                var t = ExcelAsyncUtil.Run("NEXAR_SUPPLY_LOGIN", new object[] { clientId, clientSecret }, delegate
                {
                    QueryManager.NexarToken = GetNexarTokenAsync().Result;
                    if (string.IsNullOrEmpty(QueryManager.NexarToken))
                    {
                        QueryManager.NexarTokenExpires = DateTime.MaxValue;
                        return "Unable to login to Nexar application, check Client Id and Secret";
                    }
                    else
                    {
                        QueryManager.NexarTokenExpires = DateTime.UtcNow + TimeSpan.FromDays(1); // TODO: Better to extract exp from JWT
                        return "The Nexar Supply Add-in is ready!";
                    }
                });
            }

            if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
                return "Please provide your Nexar application Client Id and Secret";
            else if (string.IsNullOrEmpty(QueryManager.NexarToken))
                return "Unable to login to Nexar application, check Client Id and Secret";
            else if (QueryManager.NexarTokenExpires < DateTime.UtcNow)
                return "The access token has expired, please refresh login!";
            else
                return "The Nexar Supply Add-in is ready!";
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Displays the internal Nexar token signifying a successful login", IsHidden = true, IsVolatile = true)]
        public static object NEXAR_SUPPLY_DEV_TOKEN()
        {
            return QueryManager.NexarToken;
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Displays the expiry time in UTC of the internal Nexar token", IsHidden = true, IsVolatile = true)]
        public static object NEXAR_SUPPLY_DEV_TOKEN_EXPIRES()
        {
            return QueryManager.NexarTokenExpires.ToString();
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Displays the internal features requested", IsHidden = true, IsVolatile = true)]
        public static object NEXAR_SUPPLY_FEATURES()
        {
            return "Datasheets: " + (QueryManager.IncludeDatasheets ? "1" : "0") + "; LeadTime: " + (QueryManager.IncludeLeadTime ? "1" : "0");
        }

        [ExcelFunction(Category = "Nexar Supply Queries", Description = "Displays the Nexar Supply API version", IsVolatile = true)]
        public static object NEXAR_SUPPLY_VERSION()
        {
            return Nexar.Supply.Api.ApiV4.GetVersion();
        }
        #endregion

        #region Private-Helper-Methods
        private static async Task<string> GetNexarTokenAsync()
        {
            string authority = "https://identity.nexar.com/";
            var tokenEndpoint = new Uri(authority).AbsoluteUri + "connect/token";

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;  // TODO: Put this in AutoOpen() ?

            using (var request = new ClientCredentialsTokenRequest { Address = tokenEndpoint, ClientId = QueryManager.NexarClientId, ClientSecret = QueryManager.NexarClientSecret})
            {                
                HttpClient httpClient = new HttpClient();
                TokenResponse tokenResponse = await httpClient.RequestClientCredentialsTokenAsync(request);

                string token = tokenResponse.AccessToken;
                if (token == null)
                    token = "";

                return token;
            }
        }

        public static AuthorizedSeller Parse(string authorizedSeller)
        {
            if (!string.IsNullOrEmpty(authorizedSeller))
            {
                switch (char.ToLower(authorizedSeller[0]))
                {
                    case 'y':
                        return AuthorizedSeller.Yes;
                    case 'n':
                        return AuthorizedSeller.No;
                    default:
                        return AuthorizedSeller.Any;
                }
            }

            return AuthorizedSeller.Any;
        }

        private static Part SearchAndWaitPart(string mpnOrSku, string manuf)
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpnOrSku)))
                QueryManager.QueryNext(mpnOrSku);

            List<Part> parts = QueryManager.GetParts(mpnOrSku);
            for (int i = 0; i < 10000 && !QueryManager.IsQueryLimitMaxed(mpnOrSku) && (parts.Count == 0); i++, Thread.Sleep(1))
            {
                parts = QueryManager.GetParts(mpnOrSku);
                if (parts.Count(part => string.IsNullOrEmpty(manuf) || part.Manufacturer.Name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpnOrSku)))
                        break;
                    QueryManager.QueryNext(mpnOrSku);
                }
            }

            return GetManuf(mpnOrSku, manuf);
        }

        private static Offer SearchAndWaitOffer(string mpn_or_sku, string manuf, AuthorizedSeller auth, string distributor = "")
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                QueryManager.QueryNext(mpn_or_sku);

            for (int i = 0; i < 1000 && !QueryManager.IsQueryLimitMaxed(mpn_or_sku); i++, Thread.Sleep(10))
            {
                List<Part> parts = QueryManager.GetParts(mpn_or_sku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.Manufacturer.Name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                        break;
                    QueryManager.QueryNext(mpn_or_sku);
                }
                else if (string.IsNullOrEmpty(distributor))
                {
                    break;
                }
                else
                {
                    // Search for specified distributor
                    var offers = GetAllOffers(parts, auth, distributor);
                    if (offers.Count == 0)
                    {
                        if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                            break;
                        QueryManager.QueryNext(mpn_or_sku);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return GetOffer(mpn_or_sku, manuf, auth, distributor);
        }

        private static List<Offer> SearchAndWaitOffers(string mpn_or_sku, string manuf, AuthorizedSeller auth, string distributor = "")
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                QueryManager.QueryNext(mpn_or_sku);

            for (int i = 0; i < 1000 && !QueryManager.IsQueryLimitMaxed(mpn_or_sku); i++, Thread.Sleep(10))
            {
                List<Part> parts = QueryManager.GetParts(mpn_or_sku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.Manufacturer.Name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                        break;
                    QueryManager.QueryNext(mpn_or_sku);
                }
                else if (string.IsNullOrEmpty(distributor))
                {
                    break;
                }
                else
                {
                    // Search for specified distributor
                    var offers = GetAllOffers(parts, auth, distributor);
                    if (offers.Count == 0)
                    {
                        if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                            break;
                        QueryManager.QueryNext(mpn_or_sku);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return GetOffers(mpn_or_sku, manuf, auth, distributor);
        }

        private static List<Offer> SearchAndWaitOffers(string mpn_or_sku, string manuf, AuthorizedSeller auth, List<string> distributors)
        {
            // Check for errors. If one exists, force a refresh
            if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                QueryManager.QueryNext(mpn_or_sku);

            for (int i = 0; i < 1000 && !QueryManager.IsQueryLimitMaxed(mpn_or_sku); i++, Thread.Sleep(10))
            {
                List<Part> parts = QueryManager.GetParts(mpn_or_sku);
                if (parts.Count(item => string.IsNullOrEmpty(manuf) || item.Manufacturer.Name.Sanitize().Contains(manuf.Sanitize())) == 0)
                {
                    if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                        break;
                    QueryManager.QueryNext(mpn_or_sku);
                }
                else if (distributors == null || distributors.Count == 0)
                {
                    break;
                }
                else
                {
                    // Search for specified distributor
                    var offers = GetAllOffers(parts, auth, distributors);
                    if (offers.Count == 0)
                    {
                        if (!string.IsNullOrEmpty(QueryManager.GetLastError(mpn_or_sku)))
                            break;
                        QueryManager.QueryNext(mpn_or_sku);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return GetOffers(mpn_or_sku, manuf, auth, distributors);
        }

        private static Part GetManuf(string mpnOrSku, string manuf)
        {
            List<Part> parts = QueryManager.GetParts(mpnOrSku);
            return parts.FirstOrDefault(part => string.IsNullOrEmpty(manuf) || part.Manufacturer.Name.Sanitize().Contains(manuf.Sanitize()));
        }

        private static Offer GetOffer(string mpnOrSku, string manuf, AuthorizedSeller auth, string distributor = "")
        {
            List<Part> parts = QueryManager.GetParts(mpnOrSku);
            List<Seller> sellers = parts.SelectMany(offer => offer.Sellers).ToList();
            List<Seller> filteredSellers = FilterSellers(auth, distributor, sellers);
            List<Offer> offers = filteredSellers.SelectMany(seller => seller.Offers).ToList();
            return offers.FirstOrDefault();
        }

        private static List<Offer> GetOffers(string mpnOrSku, string manuf, AuthorizedSeller auth, string distributor = "")
        {
            List<Part> parts = QueryManager.GetParts(mpnOrSku);
            List<Seller> sellers = parts.SelectMany(offer => offer.Sellers).ToList();
            List<Seller> filteredSellers = FilterSellers(auth, distributor, sellers);
            List<Offer> offers = filteredSellers.SelectMany(seller => seller.Offers).ToList();
            return offers;
        }

        private static object GetOffersStock(List<Offer> offers)
        {
            long stock = offers.Max(offer => offer.InventoryLevel);
            switch (stock)
            {
                case -1: return "Non-stocked";
                case -2: return "Yes";
                case -3: return "Unknown";
                case -4: return "RFQ";
                default:
                    return stock;
            }
        }

        private static List<Seller> FilterSellers(AuthorizedSeller auth, string distributor, List<Seller> sellers)
        {
            List<Seller> filteredSellers = sellers.Where(seller => string.IsNullOrEmpty(distributor) || seller.Company.Name.Sanitize().Contains(distributor.Sanitize())).ToList();
            List<Seller> allowedSellers = filteredSellers.Where(seller =>
            {
                switch (auth)
                {
                    case AuthorizedSeller.Yes:
                        return seller.IsAuthorized;
                    case AuthorizedSeller.No:
                        return !seller.IsAuthorized;
                    default:
                        return true;
                }
            }).ToList();
            return allowedSellers;
        }

        private static List<Offer> GetOffers(string mpnOrSku, string manuf, AuthorizedSeller auth, List<string> distributors)
        {
            List<Part> parts = QueryManager.GetParts(mpnOrSku);
            List<Seller> sellers = parts.SelectMany(offer => offer.Sellers).ToList();
            List<Seller> allowedSellers = sellers.Where(
                seller => distributors == null || distributors.Count == 0 || distributors.Any(d => seller.Company.Name.Sanitize().Contains(d.Sanitize()))
            ).ToList();
            List<Seller> filteredSellers = FilterSellers(auth, string.Empty, allowedSellers);
            List<Offer> offers = filteredSellers.SelectMany(seller => seller.Offers).ToList();
            return offers;
        }

        private static List<Offer> GetAllOffers(List<Part> parts, AuthorizedSeller auth, string distributor = "")
        {
            List<Seller> sellers = parts.SelectMany(offer => offer.Sellers).ToList();
            List<Seller> filteredSellers = FilterSellers(auth, distributor, sellers);
            List<Offer> offers = filteredSellers.SelectMany(seller => seller.Offers).ToList();
            return offers;
        }

        private static List<Offer> GetAllOffers(List<Part> parts, AuthorizedSeller auth, List<string> distributors)
        {
            List<Seller> sellers = parts.SelectMany(offer => offer.Sellers).ToList();
            List<Seller> allowedSellers = sellers.Where(
                seller => distributors == null || distributors.Count == 0 || distributors.Any(d => seller.Company.Name.Sanitize().Contains(d.Sanitize()))
            ).ToList();
            List<Seller> filteredSellers = FilterSellers(auth, string.Empty, allowedSellers);
            List<Offer> offers = filteredSellers.SelectMany(seller => seller.Offers).ToList();
            return offers;
        }

        private static List<string> GetDistributors(Object[] distributors)
        {
            List<string> cleanDistributors = distributors.ToList()
                .Where(x => x != ExcelMissing.Value && x != ExcelEmpty.Value)
                .Select(x => x.ToString())
                .ToList();
            return cleanDistributors;
        }
        #endregion
    }
}