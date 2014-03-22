using OpenHDKSharp.V3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenHDKSharp.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            HwpV3Document doc = HwpV3Document.Load("v3.hwp");
            foreach (var eachPara in doc.Paragraph)
                Console.WriteLine(eachPara.ToString());
        }
    }
}
