using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;

namespace JFetch {
	public static class JFetch {

		static internal T[,] To2D<T>(T[][] source) {
			try {
				int FirstDim = source.Length;
				int SecondDim = source.GroupBy(row => row.Length).Single().Key; // throws InvalidOperationException if source is not rectangular

				var result = new T[FirstDim, SecondDim];
				for (int i = 0; i < FirstDim; ++i)
					for (int j = 0; j < SecondDim; ++j)
						result[i, j] = source[i][j];

				return result;
			}
			catch (InvalidOperationException) {
				throw new InvalidOperationException("The given jagged array is not rectangular.");
			} 
		}

		public static async Task<object[,]> JFetchAsync(string url) {
			var client = new HttpClient();
			var response = await client.GetAsync(url).ConfigureAwait(false);
			response.EnsureSuccessStatusCode();
			var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
			var d = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(j);

			List<List<object>> tbl = new List<List<object>>();
			foreach (Dictionary<string, string> dict in d) {
				List<object> currentRow = new List<object>();
				foreach (KeyValuePair<string, string> kvp in dict) {
					currentRow.Add(kvp.Value);
				}
				tbl.Add(currentRow);
			}

			object[][] result;
			result = tbl.Select(l => l.ToArray()).ToArray();
			object[,] final;
			final = To2D(result);
			return final;
		}

		public static object[,] JFetchSync(string url) {
			var client = new HttpClient();
			var response = client.GetAsync(url).Result;

			if (response.IsSuccessStatusCode) {
				var responseContent = response.Content;
				var j = responseContent.ReadAsStringAsync().Result;
				var d = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(j);

				List<List<object>> tbl = new List<List<object>>();
				foreach (Dictionary<string, string> dict in d) {
					List<object> currentRow = new List<object>();
					foreach (KeyValuePair<string, string> kvp in dict) {
						currentRow.Add(kvp.Value);
					}
					tbl.Add(currentRow);
				}

				object[][] result;
				result = tbl.Select(l => l.ToArray()).ToArray();
				object[,] final;
				final = To2D(result);
				return final;
			}
			else {
				object[,] retError = { { "HTTP error" } };
				return retError;
			}

		}

	}
}
