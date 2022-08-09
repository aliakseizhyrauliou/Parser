using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Newtonsoft.Json;

namespace ParserTest
{
    class Program
    {
        const bool isVivsibleExcel = true;
        public static List<string> badWords = new List<string>();
        static void Main(string[] args)
        {
            var req = new GetRequest();
            var textReader = new TextFileReader();
            var words =  textReader.Read(Path.GetFullPath("../../../ExcelAndKeys/Keys.txt")).Result;
            var dataList = new List<Responce>();
            foreach (var word in words)
            {
                req.Run(word);
                if (req.Response== "{}") 
                {
                    Console.WriteLine($"Cant find '{word}'");
                    badWords.Add(word);
                    continue;
                }
                var responce = JsonConvert.DeserializeObject<Responce>(req.Response);
                dataList.Add(ChangePrices(responce));
            }
            
            ExcelGenerator.DisplayInExcel(DeleteBadRequestWords(badWords, words), dataList, isVivsibleExcel);

        }

        public static List<string> DeleteBadRequestWords(List<string> badWords, List<string> keys)
        {
            badWords.ForEach(x =>
            {
                if (keys.Contains(x))
                {
                    keys.Remove(x);
                }
            });

            return keys;
        }

        public static Responce ChangePrices(Responce responce)
        {
            responce.Data.Products.ForEach(x =>
            {
                x.PriceU = int.Parse(x.PriceU.ToString().Substring(0, x.PriceU.ToString().Length - 2));
            });

            return responce;
        }



    }

}
