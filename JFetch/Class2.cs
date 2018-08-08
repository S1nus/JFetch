using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace JFetch {
	class Class2 {

		public static List<Dictionary<string, string>> table;

		public static Task<int> JFetch(string url) {
			return Task.Factory.StartNew(async () => {
				var client = new HttpClient();
				var response = await client.GetAsync(url).ConfigureAwait(false);
				response.EnsureSuccessStatusCode();
				var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
				var d = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(j);
				table = d;
				return 0;
			}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default).Unwrap();
		}

		public static void Main(string[] args) {
			JFetch("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json");
			while (true) {
				try {
					foreach (KeyValuePair<string, string> kvp in table[0]) {
						Console.WriteLine("Key = {0}, Value = {1}", kvp.Key, kvp.Value);
					}
				}
				catch (Exception e) {
					Console.WriteLine("Not yet bruh");
				}
			}
		}
	}
}
