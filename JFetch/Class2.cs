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

		//public static Task<int> JFetch(string url) {
		//	return Task.Factory.StartNew(async () => {
		//		var client = new HttpClient();
		//		var response = await client.GetAsync(url).ConfigureAwait(false);
		//		response.EnsureSuccessStatusCode();
		//		var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
		//		var d = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(j);
		//		table = d;
		//		return 0;
		//	}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default).Unwrap();
		//}

		public static async Task<List<Dictionary<string, string>>> JFetchAsync(string url) {
			var client = new HttpClient();
			var response = await client.GetAsync(url).ConfigureAwait(false);
			response.EnsureSuccessStatusCode();
			var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
			var d = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(j);
			return d;
		}

		public static async void PrintTblAsync() {
			table = await JFetchAsync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json");
		}

		public static void Main(string[] args) {
			PrintTblAsync();
			while (true) {
				Console.WriteLine(table);
			}
		}
	}
}