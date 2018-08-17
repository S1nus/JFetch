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
using Newtonsoft.Json.Linq;
using OrionStuff;

namespace JFetch {
	public static class ExcelAddin {

		/*
		[ExcelFunction(Description = "Print info about Kings")]
		public static object GetKings() {
			return ExcelAsyncUtil.Run("GetKings", new object[] { }, () => JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client));
		}

		[ExcelFunction(Description = "Get Kings Async and Resize")]
		public static object GetKingsResize() {
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
		*/

		[ExcelFunction(Description = "Attempt at FP_Focus")]
		public static object Fp_Focus_Patch(string groupname, string date) {
			ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			return ExcelAsyncUtil.Observe("Fp_Focus_Patch", null, delegate {
				TaskCompletionSource<object[,]> tcs = new TaskCompletionSource<object[,]>();

				Task.Factory.StartNew(async delegate {
					object[,] fpresult = await OrionStuff.Orion.FP_Focus(groupname, date);
					try {
						tcs.SetResult((object[,])ArrayResizer.Resize(fpresult, caller));
					}
					catch (Exception ex) {
						tcs.SetResult(new object[,] { { ex.ToString() } });
					}
				}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);

				return new ExcelTaskObservable<object[,]>(tcs.Task);
			});
		}

		[ExcelFunction(Description = "Print the account IDs")]
		public static object Fp_Print_Ids() {
			ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			return ExcelAsyncUtil.Observe("Fp_Focus_Patch", null, delegate {
				TaskCompletionSource<object[,]> tcs = new TaskCompletionSource<object[,]>();

				Task.Factory.StartNew(async delegate {
					//object[,] fpresult = await OrionStuff.Orion.FP_Focus(groupname, date);
					await Orion.GetAccountIds();
					object[,] ret = new object[Orion.accountIds.Count, 2];
					for (int i = 0; i<Orion.accountIds.Count; i++) {
						ret[i, 0] = Orion.accountIds.ElementAt(i).Key;
						ret[i, 1] = Orion.accountIds.ElementAt(i).Value;
					}
					try {
						tcs.SetResult((object[,])ArrayResizer.Resize(ret, caller));
					}
					catch (Exception ex) {
						tcs.SetResult(new object[,] { { ex.ToString() } });
					}
				}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);

				return new ExcelTaskObservable<object[,]>(tcs.Task);
			});
		}

		/*
		[ExcelFunction(Description = "Get FP_Focus Data from Orion")]
		public static object Orion_FP_Focus_Patch(string group, string date) {
			var toPost = @"{			
				'prompts': [
					{
						'id' : 17307,
						'code' : '@asof',
						'prompt' : 'As Of Date',
						'promptDescription' : '',
						'promptType' : 'Date',
						'defaultValue' : '{0}',
						'isPromptUser' : true,
						'sortOrder' : null
					},
					{
						'id' : 23342,
						'code' : '@group',
						'prompt' : 'Group',
						'promptDescription' : 'Enter FPSUP, CMSUP, OASUP, ACSUP, CCSUP, ACBALA1, ACBALA2, ACBALA3, MISUP, or EQUITY',
						'promptType' : 'Text',
						'defaultValue' : '{1}',
						'isPromptUser' : true,
						'sortOrder' : null
					}
				]	
			}";
			var toSend = String.Format(toPost, date, group);
			var content = new StringContent(toSend, Encoding.UTF8, "application/json");
			ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			return ExcelAsyncUtil.Observe("GetKingsResize", null, delegate {
				TaskCompletionSource<object[,]> tcs = new TaskCompletionSource<object[,]>();

				Task.Factory.StartNew(async delegate {
					try {
						//tcs.SetResult((object[,])ArrayResizer.Resize(JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client), caller));
						var response = await client.PostAsync("https://api.orionadvisor.com/api/v1/Reporting/custom/13095/Generate/Table", content);
						var j = await response.Content.ReadAsStringAsync();
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
						final = JFetch.To2D(result);
						tcs.SetResult((object[,])ArrayResizer.Resize(final));
					} catch (Exception ex) {
						tcs.SetResult(new object[,]{ { ex.ToString()} });
					}
				}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);

				return new ExcelTaskObservable<object[,]>(tcs.Task);
			});
		}
		*/



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