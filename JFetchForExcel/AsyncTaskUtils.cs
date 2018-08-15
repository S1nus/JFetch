using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;

namespace JFetchUtils {
	// Helpers for creating and using Task-based functions with Excel-DNA's RTD-based IObservable support
	public static class AsyncTaskUtils {
		public static Task ForEachAsync<T>(this IEnumerable<T> source, int dop, Func<T, Task> body) {
			return Task.WhenAll(
				from partition in Partitioner.Create(source).GetPartitions(dop)
				select Task.Factory.StartNew(async delegate {
					using (partition)
						while (partition.MoveNext())
							await body(partition.Current).ConfigureAwait(false);
				}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default).Unwrap());
		}

		public static object RunTask<TResult>(string callerFunctionName, object callerParameters, Func<Task<TResult>> taskSource) {
			// return callerFunctionName + " : " + callerParameters.ToString();
			return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate {
				var task = taskSource();
				return new ExcelTaskObservable<TResult>(task);
			});
		}

		public static object Run_Task<TResult>(string callerFunctionName, object callerParameters, Task<TResult> taskSource) {
			// return callerFunctionName + " : " + callerParameters.ToString();
			return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () => {
				return new ExcelTaskObservable<TResult>(taskSource);
			});
		}

		// Careful - this might only work as long as the task is not shared between calls, since cancellation cancels that task
		public static object RunTaskWithCancellation<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, Task<TResult>> taskSource) {
			return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate {
				var cts = new CancellationTokenSource();
				var task = taskSource(cts.Token);
				return new ExcelTaskObservable<TResult>(task, cts);
			});
		}

		public static object RunAsTask<TResult>(string callerFunctionName, object callerParameters, Func<TResult> function) {
			return RunTask(callerFunctionName, callerParameters, () => Task.Factory.StartNew(function));
		}

		public static object RunAsTaskWithCancellation<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, TResult> function) {
			return RunTaskWithCancellation(callerFunctionName, callerParameters, cancellationToken => Task.Factory.StartNew(() => function(cancellationToken), cancellationToken));
		}

		// Helper class to wrap a Task in an Observable - allowing one Subscriber.
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

	sealed class DefaultDisposable : IDisposable {
		public static readonly DefaultDisposable Instance = new DefaultDisposable();

		// Prevent external instantiation
		DefaultDisposable() { }


		public void Dispose() {
			// no op
		}
	}

	sealed class CancellationDisposable : IDisposable {
		bool _suppress;
		readonly CancellationTokenSource _cts;

		public CancellationDisposable(CancellationTokenSource cts) {
			_cts = cts ?? throw new ArgumentNullException("cts");
		}

		public CancellationDisposable()
			: this(new CancellationTokenSource()) {
		}

		public void SuppressCancel() {
			_suppress = true;
		}

		public CancellationToken Token {
			get { return _cts.Token; }
		}

		public void Dispose() {
			if (!_suppress) _cts.Cancel();
			_cts.Dispose();  // Not really needed...
		}
	}

	// This is not a very elegant IObservable implementation - should not be public.
	// It basically represents a Subject 
	public class ThreadPoolDelegateObservable<TResult> : IExcelObservable {
		//readonly ExcelFunc _func;
		Task<TResult> _func;
		bool _subscribed;

		public ThreadPoolDelegateObservable(Task<TResult> func) {
			_func = func;
		}

		public IDisposable Subscribe(IExcelObserver observer) {
			if (_subscribed) throw new InvalidOperationException("Only single Subscription allowed.");
			_subscribed = true;

			ThreadPool.QueueUserWorkItem(async delegate (object state) {
				try {
					object result = await (Task<TResult>)state;
					observer.OnNext(result);
					observer.OnCompleted();
				} catch (Exception ex) {
					// TODO: Log somehow?
					observer.OnError(ex);
				}
			}, _func);

			return ThreadPoolDisposable.Instance;
		}

		class ThreadPoolDisposable : IDisposable {
			public static readonly ThreadPoolDisposable Instance = new ThreadPoolDisposable();

			private ThreadPoolDisposable() { }

			public void Dispose() { }
		}

	}
}