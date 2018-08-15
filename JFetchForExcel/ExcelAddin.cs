using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Net.Http;
using JFetch;
using System.Collections.Concurrent;
using System.Timers;
using System.Threading;
using Newtonsoft.Json;
using JFetchUtils;

namespace JFetch {
	public static class ExcelAddin {

		private static bool loggedIn = false;
		private static HttpClient client = new HttpClient();
		private static string token = "";

		[ExcelFunction(Description = "Print info about Kings")]
		public static object GetKings() {
			return ExcelAsyncUtil.Run("GetKings", new object[] { }, () => JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client));
		}

		[ExcelFunction(Description = "Get Kings Async and Resize")]
		/*public static object GetKingsResize() {
			return ExcelAsyncUtil.Run("GetKingsResizes", new object[] { }, () => {
				return ArrayResizer.Resize(JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client));	
			});
		}*/

		public static object GetKingsResize() {
			try {
				ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;

				return ExcelAsyncUtil.Observe("GetKingsResize", "just:somestuff", delegate {
					TaskCompletionSource<object> tcs = new TaskCompletionSource<object>();

					Task.Factory.StartNew(async delegate {
						try {
							//tcs.SetResult(/*stuff*/);
							tcs.SetResult(
								await AsyncQueryTopic(caller, new object[] { }).ConfigureAwait(false);
							);
						}
						catch (Exception ex) {
							tcs.SetException(ex);
						}
					}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);

					return new AsyncTaskUtils.ExcelTaskObservable<object>(tcs.Task);
				});
			}
			catch (Exception ex) {
				return new object[,] { { ex.ToString() } };
			}
		}

		private static async Task<object[,]> AsyncQueryTopic(ExcelReference caller, object[] args) {
			var result = await JFetch.JFetch.JFetchAsync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client);

			return (object[,])ArrayResizer.Resize(result, caller);
		}



		/*private async Task<object [,]> DoKingFetchAsync(ExcelReference caller, object[] args) {

		}*/

		/*public static object GetFpFocus() {
			if (!loggedIn) {
				authAsync("intern4@fpcm.net", "Unix15cool");
				return null;
			}

			Dictionary<object, object> toPost = new Dictionary<object, object>();
			toPost["prompts"] = new List<Dictionary<string, string>>();
			Dictionary<string, string> date = new Dictionary<string, string>();

			client.DefaultRequestHeaders.Clear();
			client.DefaultRequestHeaders.Add("Authorization", "Session " + token);
			return ArrayResizer.Resize(JFetch.JFetchSync("https://api.orionadvisor.com/api/v1/reporting/custom/13095/generate/table", client));
		}*/

		/*private static async void authAsync(string username, string password) {
			client.DefaultRequestHeaders.Add("Authentication", "Basic " + Base64Encode(username + ":" + password));
			var response = await client.GetAsync("https://api.orionadvisor.com/api/v1/Security/Token").ConfigureAwait(false);
			response.EnsureSuccessStatusCode();
			var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
			Dictionary<string, string> respDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(j);
			token = respDict["access_token"];
			loggedIn = true;
		}*/

		/*internal static string Base64Encode(string plaintext) {
			var plaintextbytes = System.Text.Encoding.UTF8.GetBytes(plaintext);
			return System.Convert.ToBase64String(plaintextbytes);
		}*/
		
	}
	public class ResizeResult {
		public bool Resize { get; set; }
		public object[,] Result { get; set; }

		public ResizeResult(bool r, object[,] res) {
			Resize = r;
			Result = res;
		}
	}

}