﻿using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;


// Semi-automatically generated...
// <auto-generated />

/// <summary>
/// A representation of the data returned from the Nexar Supply GraphQL API.
//  Unfortunately, it looks like we Strawberry Shake won't play nicely with the Excel add-in project.
//  So these classes were generated manually using this excellent tool - https://quicktype.io/csharp.
/// </summary>
namespace Nexar.Supply.SupplySchema
{
    public partial class SupplyResult
    {
        [JsonProperty("data")]
        public Data Data { get; set; }
    }

    public partial class Data
    {
        [JsonProperty("supMultiMatch")]
        public List<SupMultiMatch> SupMultiMatch { get; set; }
    }

    public partial class SupMultiMatch
    {
        [JsonProperty("reference")]
        public string Reference { get; set; }

        [JsonProperty("error")]
        public string Error { get; set; }

        [JsonProperty("hits")]
        public long Hits { get; set; }

        [JsonProperty("parts")]
        public List<Part> Parts { get; set; }
    }

    public partial class Part
    {
        [JsonProperty("v3uid")]
        public string V3Uid { get; set; }

        [JsonProperty("mpn")]
        public string Mpn { get; set; }

        [JsonProperty("manufacturer")]
        public Manufacturer Manufacturer { get; set; }

        [JsonProperty("bestDatasheet")]
        public SupDocument BestDatasheet { get; set; }

        [JsonProperty("octopartUrl")]
        public Uri OctopartUrl { get; set; }

        [JsonProperty("sellers")]
        public List<Seller> Sellers { get; set; }
    }

    public partial class Manufacturer
    {
        [JsonProperty("id")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("homepageUrl")]
        public string HomepageUrl { get; set; }
    }

    public partial class SupDocument
    {
        [JsonProperty("url")]
        public string Url { get; set; }
    }

    public partial class Seller
    {
        [JsonProperty("offers")]
        public List<Offer> Offers { get; set; }

        [JsonProperty("company")]
        public Manufacturer Company { get; set; }

        [JsonProperty("isAuthorized")]
        public bool IsAuthorized { get; set; }
    }

    public partial class Offer
    {
        [JsonProperty("id")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long Id { get; set; }

        [JsonProperty("sku")]
        public string Sku { get; set; }

        [JsonProperty("factoryLeadDays")]
        public long? FactoryLeadDays { get; set; }

        [JsonProperty("factoryPackQuantity")]
        public long? FactoryPackQuantity { get; set; }

        [JsonProperty("inventoryLevel")]
        public long InventoryLevel { get; set; }

        [JsonProperty("onOrderQuantity")]
        public long? OnOrderQuantity { get; set; }

        [JsonProperty("orderMultiple")]
        public long? OrderMultiple { get; set; }

        [JsonProperty("multipackQuantity")]
        public long? MultipackQuantity { get; set; }

        [JsonProperty("packaging")]
        public string Packaging { get; set; }

        [JsonProperty("moq")]
        public long? Moq { get; set; }

        [JsonProperty("clickUrl")]
        public Uri ClickUrl { get; set; }

        [JsonProperty("updated")]
        public DateTimeOffset Updated { get; set; }

        [JsonProperty("prices")]
        public List<Price> Prices { get; set; }
    }

    public partial class Price
    {
        [JsonProperty("currency")]
        public string Currency { get; set; }

        [JsonProperty("quantity")]
        public long Quantity { get; set; }

        [JsonProperty("price")]
        public double PricePrice { get; set; }
    }
        
    public partial class SupplyResult
    {
        public static SupplyResult FromJson(string json) => JsonConvert.DeserializeObject<SupplyResult>(json, SupplySchema.Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this SupplyResult self) => JsonConvert.SerializeObject(self, SupplySchema.Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

    internal class ParseStringConverter : JsonConverter
    {
        public override bool CanConvert(Type t) => t == typeof(long) || t == typeof(long?);

        public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null) return null;
            var value = serializer.Deserialize<string>(reader);
            long l;
            if (Int64.TryParse(value, out l))
            {
                return l;
            }
            throw new Exception("Cannot unmarshal type long");
        }

        public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer)
        {
            if (untypedValue == null)
            {
                serializer.Serialize(writer, null);
                return;
            }
            var value = (long)untypedValue;
            serializer.Serialize(writer, value.ToString());
            return;
        }

        public static readonly ParseStringConverter Singleton = new ParseStringConverter();
    }
}
