using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JFetch;
using OrionStuff;

namespace Tester {
	class Program {
		static void Main(string[] args) {
			Console.WriteLine("Hello world!");
			try {
				//Console.WriteLine(Orion.GetStratGroups().Result);
				object[,] stratarray = Orion.GetStratGroups().Result;
				List<object> groups = new List<object>();
				for (int i = 0; i < stratarray.GetLength(0); i++) {
					groups.Add(stratarray[i, 0]);
				}
				foreach (object obj in groups) {
					Console.WriteLine(obj);
				}
			}
			catch (Exception ex) {
				Console.WriteLine(ex.ToString());
			}
			Console.Read();
		}
	}
}
