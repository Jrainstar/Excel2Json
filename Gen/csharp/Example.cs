using Newtonsoft.Json;
using System.Collections.Generic;

namespace Jrainstar.Table
{
	/// <summary>
	/// 示例
	/// <summary>
	public class Example
	{
		/// <summary>
		/// 编号
		/// </summary>
		public int id { get; set; }

		/// <summary>
		/// 名称
		/// </summary>
		public string name { get; set; }

		/// <summary>
		/// int32
		/// </summary>
		public int int32 { get; set; }

		/// <summary>
		/// 数组
		/// </summary>
		public string[] array { get; set; }

		/// <summary>
		/// 字典
		/// </summary>
		public Dictionary<int,int> dict { get; set; }

		/// <summary>
		/// 交错数组
		/// </summary>
		public int[][] arrayArray { get; set; }

		/// <summary>
		/// 字典数组
		/// </summary>
		public Dictionary<int,int[]> dictArray { get; set; }

		private static Dictionary<int, Example> mainGroup { get; set; }
		private static Dictionary<string, Dictionary<int, Example>> tagGroup { get; set; }

		public static Example Get(int id, string tag = "")
		{
			if (string.IsNullOrEmpty(tag))
			{
				if (mainGroup == null)
				{
					mainGroup = JsonConvert.DeserializeObject<Dictionary<int, Example>>(TableBuilder.Load(typeof(Example).Name));
				}
				mainGroup.TryGetValue(id, out Example value);
				return value;
			}
			else
			{
				if (tagGroup == null)
				{
					tagGroup = new Dictionary<string, Dictionary<int, Example>>();
				}
				if (!tagGroup.ContainsKey(tag))
				{
					tagGroup[tag] = JsonConvert.DeserializeObject<Dictionary<int, Example>>(TableBuilder.Load($"{typeof(Example).Name}_{tag}"));
				}
				tagGroup[tag].TryGetValue(id, out Example value);
				return value;
			}
		}
	}
}
