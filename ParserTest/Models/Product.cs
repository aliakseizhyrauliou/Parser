using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace ParserTest
{
    class Product
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("brand")]
        public string Brand { get; set; }

        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("feedbacks")]
        public int Feedbacks { get; set; }

        [JsonPropertyName("priceU")]
        public double PriceU { get; set; }
    }
}
