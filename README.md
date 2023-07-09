# Excel2Json

```c#
TableBuilder.onLoad = (string fileName) =>
{
    return File.ReadAllText($"C:\\Users\\asus\\Documents\\Excel2Json\\Gen\\json\\{fileName}.json");
};

Console.WriteLine(Example.Get(10001, "Alter").name);
```

- 支持数据类型

```c#
- int 						- int
- float						- float
- string					- string
- int[]						- int[]	
- float[]					- float[]
- string[]					- string[]
- int[][]					- int[][]
- float[][]					- float[][]
- string[][]				- string[][]
- int:int					- Dictionary<int,int>
- int:string				- Dictionary<int,string>
- string:int				- Dictionary<string,int>
- string:string				- Dictionary<string,string>
- int:int[]					- Dictionary<int,int[]>
- int:float[]				- Dictionary<int,float[]>
- int:string[]				- Dictionary<int,string[]>
- string:int[]				- Dictionary<string,int[]>
- string:float[]			- Dictionary<string,float[]>
- string:string[]			- Dictionary<string,string[]>
```

