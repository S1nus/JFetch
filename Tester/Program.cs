using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JFetch;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Orion;

namespace Tester {
	class Program {
		static void Main(string[] args) {
			var array = Orion.Orion.FP_Focus("foxas", "8/20/2018").Result;
			for (int i = 0; i<array.GetLength(0); i++) {
				for (int j = 0; j<array.GetLength(1); j++) {
					Console.WriteLine(array[i, j]);
				}
				Console.WriteLine("");
			}
			Console.Read();
		}
	}
}
