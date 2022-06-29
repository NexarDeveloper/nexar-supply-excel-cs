using log4net;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Nexar.Supply.Query;
using Nexar.Supply.SupplySchema;

namespace NexarSupplyXll
{
    /// <summary>
    /// Manages queries and cache to the Nexar API
    /// </summary>
    public class NexarQueryManager
    {
        #region Properties
        /// <summary>
        /// Gets or sets the Nexar API token.
        /// </summary>
        public string NexarToken { get; set; }

        /// <summary>
        /// The UTC time when the access token expires.
        /// </summary>
        public DateTime NexarTokenExpires { get; set; }
        
        /// <summary>
        /// Whether a renew login has been triggered.
        /// </summary>
        public bool NexarTokenRenewing { get; set; }

        /// <summary>
        /// Cached client ID/Secret so we don't get a new token on every update.
        /// </summary>
        public string NexarClientId { get; set; }
        public string NexarClientSecret { get; set; }

        /// <summary>
        /// Whether or not to include specific features when querying on behalf of the client.
        /// A client may not have access if they are a self-serve customer.
        /// </summary>
        public bool IncludeDatasheets { get; set; } = true;

        public bool IncludeLeadTime { get; set; } = true;

        /// <summary>
        /// Gets or sets the lower level Api timeout value
        /// </summary>
        public int HttpTimeout { get; set; }
        #endregion
        
        #region Constants
        /// <summary>
        /// Indicates the count at which time to start the query (maximum of 20 requests per query)
        /// </summary>
        private const int QUERY_COUNT_TRIGGER = 10;

        /// <summary>
        /// Indicates the time delay until starting a query
        /// </summary>
        private const int QUERY_TRIGGER_TIME = 200;

        /// <summary>
        /// Fatal error string
        /// </summary>
        public const string FATAL_ERROR = "An error occurred. Please consider logging a GitHub issue against this problem or contact Nexar directly.";
        
        public const string PROCESSING = "!!! Processing !!!";
        #endregion

        #region Variables
        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// The timer to kick off the web request (if required)
        /// </summary>
        private readonly System.Timers.Timer _queryTimer = new System.Timers.Timer();

        /// <summary>
        /// The query cache
        /// </summary>
        private readonly ConcurrentStack<CacheItem> _queryList = new ConcurrentStack<CacheItem>();
        #endregion

        #region Classes
        /// <summary>
        /// Encapsulates a Cache Item
        /// </summary>
        private class CacheItem
        {
            /// <summary>
            /// Indicates the state of the cache item
            /// </summary>
            public enum ProcessingState
            {
                /// <summary>
                /// item is awaiting to be queried
                /// </summary>
                Await,

                /// <summary>
                /// item is currently being queried
                /// </summary>
                Processing,

                /// <summary>
                /// item was queried, but the response triggered an error
                /// </summary>
                Error,

                /// <summary>
                /// item was queried, and the response has been successfully received
                /// </summary>
                Done
            }

            /// <summary>
            /// Initializes a new instance of the CacheItem class
            /// </summary>
            public CacheItem()
                : this(null, 0, Nexar.Supply.Api.ApiV4.RECORD_LIMIT_PER_QUERY) { }

            /// <summary>
            /// Initializes a new instance of the CacheItem class
            /// </summary>
            /// <param name="mpn">The query string</param>
            public CacheItem(string mpn)
                : this(mpn, 0, Nexar.Supply.Api.ApiV4.RECORD_LIMIT_PER_QUERY) { }

            /// <summary>
            /// Initializes a new instance of the CacheItem class
            /// </summary>
            /// <param name="mpn">The query string</param>
            /// <param name="start">The query item start </param>
            /// <param name="limit">The query item limit</param>
            public CacheItem(string mpn, int start, int limit)
            {
                Query = new PartsMatchQuery {limit = limit};
                Parts = new List<Part>();
                Q = mpn;
                Start = start;
                State = ProcessingState.Await;
            }

            /// <summary>
            /// Gets or sets the query string (maps to the internal query object)
            /// </summary>
            public string Q
            {
                get { return Query.q; }
                private set { Query.q = value; }
            }

            /// <summary>
            /// Gets or sets the start item (maps to the internal query object)
            /// </summary>
            public int Start
            {
                get { return Query.start; }
                private set { Query.start = value; }
            }

            /// <summary>
            /// The query that will be sent to the Nexar Supply API
            /// </summary>
            internal PartsMatchQuery Query { get; private set; }

            /// <summary>
            /// The state of the query
            /// </summary>
            public ProcessingState State { get; internal set; }

            /// <summary>
            /// Indicates an error if one is available
            /// </summary>
            public string Error { get; internal set; }

            /// <summary>
            /// Indicates how many hits the query returned
            /// </summary>
            public int Hits { get; internal set; }

            /// <summary>
            /// Gets a list of parts returned by the query
            /// </summary>
            public List<Part> Parts { get; private set; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the NexarQueryManager
        /// </summary>
        public NexarQueryManager()
        {
            _queryTimer.Elapsed += new System.Timers.ElapsedEventHandler(PerformQueryTick);
            _queryTimer.Interval = QUERY_TRIGGER_TIME;
            _queryTimer.AutoReset = false;
            HttpTimeout = 5000;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Gets a collection of parts that have been retured by completed queries that match the specified
        /// mpn_or_sku keyword
        /// </summary>
        /// <param name="mpn_or_sku">The query string</param>
        /// <returns>A list of parts that have been found from the queries</returns>
        public List<Part> GetParts(string mpn_or_sku)
        {
            var items = _queryList.Where(i => i.Q == mpn_or_sku.Sanitize() && (i.State == CacheItem.ProcessingState.Done));

            if (!items.Any())
            {
                return new List<Part>();
            }
            else
            {
                return items.SelectMany(i => i.Parts).ToList();
            }
        }

        /// <summary>
        /// Gets the error string for the specified query
        /// </summary>
        /// <param name="mpn_or_sku">The query string</param>
        /// <returns>The last error</returns>
        public string GetLastError(string mpn_or_sku)    
        {
            var item = _queryList.FirstOrDefault(i => i.Q == mpn_or_sku.Sanitize() && i.State == CacheItem.ProcessingState.Error);
            if (item != null)
            {
                return item.Error;
            }
            else
            {
                // This should never happen; a query should exist before this function is called
                return string.Empty;
            }
        }

        /// <summary>
        /// Indicates if all known parts have been received related to the query string
        /// </summary>
        /// <param name="mpn_or_sku">The query string</param>
        /// <returns>An indication if all known parts have been found</returns>
        public bool IsQueryLimitMaxed(string mpn_or_sku)
        {
            var items = _queryList.Where(i => i.Q == mpn_or_sku.Sanitize() && i.State == CacheItem.ProcessingState.Done);
            if (!items.Any())
            {
                return false;
            }
            else
            {
                return (GetParts(mpn_or_sku.Sanitize()).Count == items.Max(i => i.Hits));
            }
        }

        /// <summary>
        /// Queries the Nexar Supply API for information related to the query string
        /// </summary>
        /// <remarks>
        ///  - If the query string has already been searched, calling this function will bring up the next set
        ///    of results (usually in increments of 20 items, with a limit of up to 100 items).
        ///  - The query string will have all whitespace removed, and be brought to lower case
        ///  - If a query with the specified query string is already in progress, nothing will be done
        /// </remarks>
        /// <param name="mpn_or_sku">The query string</param>
        public void QueryNext(string mpn_or_sku)
        {
            mpn_or_sku = mpn_or_sku.Sanitize();

            // Verify that there already exists a query for the part. If not, simply create a new query
            if (_queryList.Count(query => query.Q == mpn_or_sku) == 0)
            {
                EnqueueQuery(mpn_or_sku, 0);
            }
            else
            {
                // Only add a new query if there is not already one processing, and we haven't hit max hits
                // and if previously error'd, re-request it
                var existingQueries = _queryList.Where(i => i.Q == mpn_or_sku);
                int maxstart = existingQueries.Max(i => i.Start);
                if (existingQueries.Count(i => i.State == CacheItem.ProcessingState.Error) != 0)
                {
                    existingQueries.ToList().ForEach(i => { i.State = CacheItem.ProcessingState.Await; i.Error = string.Empty; });
                    _queryTimer.Enabled = true;
                }
                else if ((existingQueries.Count(i => i.State == CacheItem.ProcessingState.Processing || i.State == CacheItem.ProcessingState.Await) == 0) && (maxstart < Nexar.Supply.Api.ApiV4.RECORD_START_MAX) && !IsQueryLimitMaxed(mpn_or_sku))
                {
                    EnqueueQuery(mpn_or_sku, maxstart + Nexar.Supply.Api.ApiV4.RECORD_LIMIT_PER_QUERY);
                }
            }
        }

        /// <summary>
        /// Empties the query cache for cached data; parts will need to be requeried again.
        /// </summary>
        /// <remarks>
        public void EmptyQueryCache()
        {
            lock (_queryList)
            {
                _queryList.Clear();
            }
        }
        
        #endregion

        #region Methods-Support
        /// <summary>
        /// Queries the Nexar Supply API for information related to the query string
        /// </summary>
        /// <remarks>
        /// Once the number of queries has reached a predetermined number, the web requested will execute.
        /// In the situation where the query list has stalled and is no longer receiving new requests, a timer
        /// will be started and will eventually trigger the web request.
        /// </remarks>
        /// <param name="mpn_or_sku">The query string</param>
        /// <param name="start">The query start item</param>
        private void EnqueueQuery(string mpn_or_sku, int start)
        {
            lock (_queryList)
            {
                if (_queryList.Count(query => query.Q == mpn_or_sku && query.Start == start) == 0)
                {
                    Log.Debug(string.Format("Adding {0}:{1}:{2} to the queue", mpn_or_sku, start, Nexar.Supply.Api.ApiV4.RECORD_LIMIT_PER_QUERY));
                    CacheItem item = new CacheItem(mpn_or_sku, start, Nexar.Supply.Api.ApiV4.RECORD_LIMIT_PER_QUERY);
                    _queryList.Push(item);

                    if (string.IsNullOrEmpty(NexarToken))
                    {
                        Log.Error("Nexar Supply client is not specified");
                        item.Error = "Nexar Supply client is not specified. Please provide a single call to '=NEXAR_SUPPLY_LOGIN(\"Your client Id\", \"Your client secret\")' somewhere in your worksheet.";
                        item.State = CacheItem.ProcessingState.Error;
                        return;
                    }
                }
                else 
                {
                    // If it exists, and is has error'd, set it to try again
                    _queryList.Where(i => i.Q == mpn_or_sku && i.Start == start && i.State == CacheItem.ProcessingState.Error).ToList()
                        .ForEach(i => { i.State = CacheItem.ProcessingState.Await; i.Error = string.Empty; });
                }
            }

            if (_queryList.Count(i => i.State == CacheItem.ProcessingState.Await) >= QUERY_COUNT_TRIGGER)
            {
                Log.Debug("TRIGGER: Count");
                ProcessQuery();
            }
            else
            {
                _queryTimer.Enabled = true;
            }
        }

        /// <summary>
        /// Kicks off a web request
        /// </summary>
        /// <param name="source">Sender</param>
        /// <param name="e">ElapsedEventArgs</param>
        private void PerformQueryTick(object source, System.Timers.ElapsedEventArgs e)
        {
            Log.Debug("TRIGGER: Timer");
            ProcessQuery();
        }

        /// <summary>
        /// Begins a web request and waits for the response.
        /// </summary>
        /// <remarks>
        /// When the data is received it is checked for errors. Only successful responses are then marked as 'Done' in the cache.
        /// </remarks>
        private void ProcessQuery()
        {
            List<CacheItem> tempList;
            lock (_queryList)
            {
                _queryTimer.Enabled = false;
                tempList = _queryList.Where(i => i.State == CacheItem.ProcessingState.Await).ToList();
                tempList.ForEach(i => i.State = CacheItem.ProcessingState.Processing);
            }

            if (tempList.Count == 0)
            {
                Log.Debug("No parts to search");
                return;
            }
            
            Log.Debug(string.Format("Performing search of {0} items", tempList.Count));
            Nexar.Supply.Api.ApiV4.SearchResponse resp = Nexar.Supply.Api.ApiV4.PartsMatch(tempList.Select(i => i.Query).ToList(), NexarToken, IncludeDatasheets, IncludeLeadTime, HttpTimeout);
            
            // If proper data wasn't provided, quit
            if (resp == null)
            {
                Log.Error("Response error: resp == null");
                tempList.ForEach(i => { i.State = CacheItem.ProcessingState.Error; i.Error = FATAL_ERROR + " (Response is null)"; });
                return;
            }

            if (!string.IsNullOrEmpty(resp.ErrorMessage))
            {
                Log.Debug("Response error:" + resp.ErrorMessage);
                tempList.ForEach(i => { i.State = CacheItem.ProcessingState.Error; i.Error = resp.ErrorMessage; });
                return;
            }

            if (!(resp.Data is Data))
            {
                Log.Error("Response error: resp != SupplySchema.Data");
                tempList.ForEach(i => { i.State = CacheItem.ProcessingState.Error; i.Error = FATAL_ERROR + " (response is not of correct type)"; });
                return;
            }

            Data supplyData = (Data)resp.Data;
            if (supplyData == null)
            {
                Log.Error("Response error: supplyData == null");
                tempList.ForEach(i => { i.State = CacheItem.ProcessingState.Error; i.Error = FATAL_ERROR + " (supplyData is null)"; });
                return;
            }

            if (!(supplyData.SupMultiMatch is List<SupMultiMatch>))
            {
                Log.Error("Response error: supplyData.SupMultiMatch != SupMultiMatch list");
                tempList.ForEach(i => { i.State = CacheItem.ProcessingState.Error; i.Error = FATAL_ERROR + " (SupMultiMatch is not of correct type)"; });
                return;
            }

            List<SupMultiMatch> supMultiMatch = (List<SupMultiMatch>)supplyData.SupMultiMatch;

            if (supMultiMatch.Count == 0)
            {
                Log.Error("Response error: supMultiMatch.count == 0");
                tempList.ForEach(i => { i.State = CacheItem.ProcessingState.Error; i.Error = FATAL_ERROR + " (supplyResultData does not have any results)"; });
                return;
            }

            // Now we're cooking on gas. Acceptable data has been received; include it in the cache
            for (int i = 0; i < supMultiMatch.Count; i++)
            {
                // Graphql doesn't return the request, but we can use multiMatch reference to return the input mpn or sku  
                string key = supMultiMatch[i].Reference ?? string.Empty; 
                
                if (!string.IsNullOrEmpty(supMultiMatch[i].Error))
                {
                    Log.Debug("Repsonse Error: " + supMultiMatch[i].Error);
                    _queryList.First(item => item.Q == key).State = CacheItem.ProcessingState.Error;
                    _queryList.First(item => item.Q == key).Error = supMultiMatch[i].Error;
                }
                else if (supMultiMatch[i].Parts == null)
                {
                    Log.Debug("supMultiMatch[i].Parts is null!");
                    _queryList.First(item => item.Q == key).State = CacheItem.ProcessingState.Error;
                    _queryList.First(item => item.Q == key).Error = "Query did not provide an adequate response";
                }
                else
                {
                    CacheItem querypart = _queryList.First(item => item.Q == key);
                    querypart.Parts.AddRange(supMultiMatch[i].Parts);
                    querypart.Hits = (int)supMultiMatch[i].Hits;
                    if (supMultiMatch[i].Parts.Count == 0)
                    {
                        querypart.Error = "Query did not provide a result. Please widen your search criteria.";
                    }
                    
                    querypart.State = CacheItem.ProcessingState.Done;
                }
            }
        }
        #endregion
    }
}
