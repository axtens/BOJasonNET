using BOJasonNET;
using System.IO;

namespace BJNTest
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			var boj = new Json();
			string str = File.ReadAllText(@"stuff.json");
			RecursiveDict thing = boj.Parse(ref str);
			object x = boj.GetKey(thing["animal"][0]["type"]);
		}
	}
}