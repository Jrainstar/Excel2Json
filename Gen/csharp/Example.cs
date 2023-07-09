using Newtonsoft.Json;
using System.Collections.Generic;

namespace Jrainstar.Table
{
	/// <summary>
	/// 示例
	/// <summary>
	public class Example : TableObject<Example>
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

	}
}
