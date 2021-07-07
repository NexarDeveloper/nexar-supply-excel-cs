using System;

namespace Nexar.Supply.Query
{
    /// <summary>
    /// Class representing the input query to look-up parts using the Nexar Supply API
    /// </summary>
    public class PartsMatchQuery
    {
        public string brand { get; set; }
        public int limit { get; set; }
        public string mpn { get; set; }
        public string mpn_or_sku { get; set; }
        public string q { get; set; }
        public string reference { get; set; }
        public string seller { get; set; }
        public string sku { get; set; }
        public int start { get; set; }
    }
}
