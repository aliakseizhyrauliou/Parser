using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserTest
{
    class TextFileReader
    {
        public async Task<List<string>> Read(string path) 
        {
            List<string> result = new List<string>();

            using (FileStream fstream = File.OpenRead(path))
            {
                using (StreamReader r = new StreamReader(fstream)) 
                {
                    string line;
                    while ((line = await r.ReadLineAsync()) != null)
                    {
                        if (line.Trim() == "" || IsDigitsOnly(line.Trim())) 
                        {
                            continue;
                        }
                        result.Add(line);                    
                    }
                }
                return result;
            }
        }


        private bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }
    }
}
