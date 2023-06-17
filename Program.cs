using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

const string config = "./conf.xml";
Excel2Json.Exprot(config);
Console.WriteLine("输出结束");
Console.ReadKey();