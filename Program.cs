const string conf = "./conf.xml";
const string path = "./path.xml";
Excel2Json.Exprot(conf, path);
Console.WriteLine("输出结束");
Console.ReadKey();

//using Jrainstar.Table;
//using System.IO;

//TableBuilder.onLoad = (string fileName) =>
//{
//    return File.ReadAllText($"C:\\Users\\asus\\Documents\\Excel2Json\\Gen\\json\\{fileName}.json");
//};

//Console.WriteLine(Example.Get(10001, "Alter").name);