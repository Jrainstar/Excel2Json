# Excel2Json
Excel2Json

Unity示例

public class Main : MonoBehaviour
{
    void Start()
    {
        TableManager.onLoad = (name) => { return Resources.Load<TextAsset>(name).text; };
        Debug.Log(Pokemon1.Get(10001).name);
    }
}
