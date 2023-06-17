# Excel2Json

```c#
// Unity示例
public class Main : MonoBehaviour
{
    void Start()
    {
        TableManager.onLoad = (name) => { return Resources.Load<TextAsset>(name).text; };
        Debug.Log(Pokemon1.Get(10001).name);
    }
}
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

