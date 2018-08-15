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
using OrionApiSdk.Utils;

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
			//return ArrayResizer.Resize(JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client));
			ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			return ExcelAsyncUtil.Observe("GetKingsResize", null, delegate {
				TaskCompletionSource<object[,]> tcs = new TaskCompletionSource<object[,]>();

				Task.Factory.StartNew(async delegate {
					try {
						tcs.SetResult((object[,])ArrayResizer.Resize(JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client), caller));
					} catch (Exception ex) {
						tcs.SetResult(new object[,]{ { ex.ToString()} });
					}
				}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);

				return new ExcelTaskObservable<object[,]>(tcs.Task);
			});
		}

		private static string GetCacheHashcode(params object[] args) {
			return Convert.ToString(args[0]) + Convert.ToString(args[1]);
		}

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

	public class ExcelTaskObservable<TResult> : IExcelObservable {
		readonly Task<TResult> _task;
		readonly CancellationTokenSource _cts;

		public ExcelTaskObservable(Task<TResult> task) {
			_task = task;
		}

		public ExcelTaskObservable(Task<TResult> task, CancellationTokenSource cts)
			: this(task) {
			_cts = cts;
		}

		public IDisposable Subscribe(IExcelObserver observer) {
			// Start with a disposable that does nothing
			// Possibly set to a CancellationDisposable later
			IDisposable disp = DefaultDisposable.Instance;

			switch (_task.Status) {
				case TaskStatus.RanToCompletion:
					observer.OnNext(_task.Result);
					observer.OnCompleted();
					break;
				case TaskStatus.Faulted:
					observer.OnError(_task.Exception.InnerException);
					break;
				case TaskStatus.Canceled:
					observer.OnError(new TaskCanceledException(_task));
					break;

				default:
					var task = _task;
					// OK - the Task has not completed synchronously
					// First set up a continuation that will suppress Cancel after the Task completes
					if (_cts != null) {
						var cancelDisp = new CancellationDisposable(_cts);
						task = _task.ContinueWith(t => {
							cancelDisp.SuppressCancel();
							return t;
						}).Unwrap();

						// Then this will be the IDisposable we return from Subscribe
						disp = cancelDisp;
					}
					// And handle the Task completion
					task.ContinueWith(t => {
						switch (t.Status) {
							case TaskStatus.RanToCompletion:
								observer.OnNext(t.Result);
								observer.OnCompleted();
								break;
							case TaskStatus.Faulted:
								observer.OnError(t.Exception.InnerException);
								break;
							case TaskStatus.Canceled:
								observer.OnError(new TaskCanceledException(t));
								break;
						}
					});
					break;
			}

			return disp;
		}
	}

}