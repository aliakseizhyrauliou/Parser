using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace ParserTest
{
    class Responce
    {
        [JsonPropertyName("data")]
        public ProductsList Data { get; set; }
        [JsonPropertyName("metadata")]
        public Metadata Metadata { get; set; }
    }
}
