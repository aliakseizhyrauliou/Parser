using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ParserTest
{
    public class GetRequest
    {
        private HttpWebRequest _request;
        string _address = "https://search.wb.ru//exactmatch/ru/common/v4/search?appType=1&couponsGeo=12,3,18,15,21&curr=rub&dest=-1029256,-102269,-2162196,-1257786&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&query={param}&reg=0&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,22,66,31,40,1,48,71&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false";

        private Dictionary<string, string> Headers { get; set; }

        public string Response { get; set; }
        public string Accept { get; set; } = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9";
        public string Host { get; set; } = "search.wb.ru";
        public string Referer { get; set; } = $"https://www.wildberries.ru/";
        public string Useragent { get; set; } = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko";

        public GetRequest()
        {
            Headers = new Dictionary<string, string>();
        }

        public void Run(string param)
        {
            _address = $"https://search.wb.ru//exactmatch/ru/common/v4/search?appType=1&couponsGeo=12,3,18,15,21&curr=rub&dest=-1029256,-102269,-2162196,-1257786&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&query={param}&reg=0&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,22,66,31,40,1,48,71&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false";
            _request = (HttpWebRequest)WebRequest.Create(_address);
            _request.Method = "Get";
            _request.CookieContainer = new CookieContainer();
            _request.Accept = Accept;
            _request.Host = Host;
            _request.Referer = Referer;
            _request.UserAgent = Useragent;

            foreach (var pair in Headers)
            {
                _request.Headers.Add(pair.Key, pair.Value);
            }

            try
            {
                HttpWebResponse response = (HttpWebResponse)_request.GetResponse();
                using (var stream = response.GetResponseStream())
                {
                    using (var reader = new StreamReader(stream))
                    {
                        Response = reader.ReadToEnd();
                    }
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
