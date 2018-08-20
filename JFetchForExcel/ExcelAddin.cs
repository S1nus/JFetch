using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace JFetchForExcel {
	public static class ExcelAddin {

		private static NonLockingHashMap<string, object[,]> ResultCache { get; set; }
		private static NonLockingHashMap<string, FunctionRange> FunctionCounters { get; set; }
		private static object[,] EmptyArray { get; set; }

		[ExcelFunction(Description = "Attempt to Patch FP_Focus")]
		public static object FP_Focus_Patch([ExcelArgument(Name = "Shortname", AllowReference = true)] object group,
												[ExcelArgument(Name = "Date", AllowReference = true)] object date) {
			ExcelReference caller = null;
			string _group = "";
			string _date = "";

			try {
				if (check_params_single(group) is int && check_params_single(date) is int) {
					_group = Convert.ToString(((ExcelReference)group).GetValue());
					_date = ParseDate(Convert.ToString(((ExcelReference)date).GetValue()));
				}
				//Paramaters have all been checked now locate hash bucket
				caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			} catch (Exception ex) {
				return new object[,] { { ex.ToString() } };
			}

			return ExcelAsyncUtil.Observe("FP_Focus_Path", GetCacheHashCode(caller, _group, _date), delegate {
				TaskCompletionSource<object> tcs = new TaskCompletionSource<object>();
				Task.Factory.StartNew(async delegate {
					try {
						tcs.SetResult(await AsyncQueryTopic(caller, new object[]{_group, _date}).ConfigureAwait(false));
					}
					catch (Exception ex) {
						tcs.SetResult(new object[,] { { ex.ToString() } });
					}
				}, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);
				return new ExcelTaskObservable<object>(tcs.Task);
			});
		}

		private static async Task<object[,]> AsyncQueryTopic(ExcelReference caller, object[] args) {
			string key = GetCacheHashCode(args);

			if (!ResultCache.TryGetValue(key, out object[,] result) || IsEmptyArray(result)) {
				//result = await BatchRunner(args).ConfigureAwait(false);
				result = await Orion.Orion.FP_Focus((string)args[0], (string)args[1]).ConfigureAwait(false);
				result = result ?? EmptyArray;

				if (!ResultCache.ContainsKey(key)) {
					ResultCache.TryAdd(key, result);
				}

			}

			ResizeResult _ret = CalculateResize(caller, result);
			return (_ret.Resize) ? (object[,])ArrayResizer.Resize(_ret.Result, caller) : _ret.Result;
		}

		private static object check_params_single(object data_point) {

			string data_point_name = "";
			if (data_point is ExcelReference) {
				if (((ExcelReference)data_point).GetValue() == ExcelEmpty.Value) {
					return ExcelEmpty.Value;
				}
				try {
					data_point_name = Convert.ToString(((ExcelReference)data_point).GetValue());
				} catch (Exception ex) { /*ErrorLog.AppendException(ex);*/ }

			} else if (data_point is string) {
				return Convert.ToString(data_point);
			}

			if (data_point_name == "")
				return ExcelEmpty.Value;

			if (data_point_name.Equals("0"))
				return ExcelEmpty.Value;

			return 1;
		}

		private static string ParseDate(string _date) {
			bool good = true;
			string date = "";
			DateTime test = new DateTime();
			try {
				good = DateTime.TryParse(_date, out test);
			} catch (Exception ex) { /*ErrorLog.AppendException(ex);*/ }

			if (good)
				date = test.ToString("MM/dd/yy");

			return date;
		}

		public class ResizeResult {
			public bool Resize { get; set; }
			public object[,] Result { get; set; }

			public ResizeResult(bool r, object[,] res) {
				Resize = r;
				Result = res;
			}
		}

		private static ResizeResult CalculateResize(ExcelReference caller, object[,] result) {
			//If the array is null, then use the empty_array
			if (result == null)
				result = EmptyArray;

			//Size of the actual array function results
			int rows = result.GetLength(0);
			int columns = result.GetLength(1);

			//Search Hash in function counters
			if (TryLookup(caller, out FunctionRange counter_range)) {
				bool _resize;
				//If result has changed size, request a resize
				if (_resize = !counter_range.CheckDimensions(rows, columns)) {
					counter_range.ResetDimensions(rows, columns);
					//counter_range.DecrementResize();
				} else
					counter_range.DecrementResize();

				//If Resize is equal to one, that means a resize has occured and the done flag should be set
				if (counter_range.CheckResize(1))
					counter_range.DoneFlag = 1;

				return new ResizeResult(_resize, result);
			} else {
				counter_range = new FunctionRange(caller);
				counter_range.ResetDimensions(rows, columns);

				//Add to function counter Map
				FunctionCounters.TryAdd(counter_range.HashCode, counter_range);

				return new ResizeResult(false, result);
			}
		}

		private static bool IsEmptyArray(object[,] arr) {
			if (arr == null)
				return false;

			if (!(arr.GetLength(0) == EmptyArray.GetLength(0) && arr.GetLength(1) == EmptyArray.GetLength(1)))
				return false;

			for (int r = 0; r < arr.GetLength(0); r++) {
				for (int c = 0; c < arr.GetLength(1); c++) {
					if (!arr[r, c].Equals(EmptyArray[r, c]))
						return false;
				}
			}

			return true;
		}

		private static string GetCacheHashCode(params object[] args) {
			return Convert.ToString(args[0]) + Convert.ToString(args[1]);
		}

		private static bool TryLookup(ExcelReference caller, out FunctionRange range) {
			List<string> _keys = FunctionCounters.Keys.ToList();
			for (int i = 0; i < _keys.Count; i++) {
				if (FunctionCounters[_keys[i]].Contains(caller)) {
					range = FunctionCounters[_keys[i]];
					return true;
				}
			}

			range = null;
			return false;
		}

		public class FunctionRange {
			//Function to get hashcode of a FunctionRange
			public static string GetHashCode(int row_first, int row_last, int col_first, int col_last, int sheetId) {
				return Convert.ToString(row_first) + Convert.ToString(row_last) + Convert.ToString(col_first) + Convert.ToString(col_last) + sheetId.ToString();
			}
			//Set when the function has returned after being Resized
			public int DoneFlag { get; set; } = 0;
			//Caller Range Variables
			public int RowFirst { get; set; } = 0;
			public int RowLast { get; set; } = 0;
			public int ColumnFirst { get; set; } = 0;
			public int ColumnLast { get; set; } = 0;
			public int SheetID { get; set; } = 0;
			//Counts the number of times this FunctionRange has been resized
			public int Resize = 0;
			public int Area { get; set; } = 0;
			//Stores last key pased to the RTD, to know when to for a Resize
			public string LastKey { get; set; } = "";

			//Constructor that takes a caller and stores its initial values
			public FunctionRange(ExcelReference caller) {
				RowFirst = caller.RowFirst;
				RowLast = caller.RowLast;
				ColumnFirst = caller.ColumnFirst;
				ColumnLast = caller.ColumnLast;
				Area = (caller.RowLast - caller.RowFirst + 1) * (caller.ColumnLast - caller.ColumnFirst + 1);
				SheetID = caller.SheetId.ToInt32();

				Resize = 2;
			}

			//Test if caller is contained by this Range
			public bool Contains(ExcelReference caller) {
				if (caller.SheetId.ToInt32() != SheetID)
					return false;
				else if (!(caller.RowFirst >= RowFirst && caller.RowFirst <= RowLast))
					return false;
				else if (!(caller.ColumnFirst >= ColumnFirst && caller.ColumnFirst <= ColumnLast))
					return false;
				else
					return true;
			}

			//Test if Function range is contained in this Range
			public bool Contains(FunctionRange caller) {
				if (caller.SheetID != SheetID)
					return false;
				else if (!(caller.RowFirst >= RowFirst && caller.RowLast <= RowLast))
					return false;
				else if (!(caller.ColumnFirst >= ColumnFirst && caller.ColumnLast <= ColumnLast))
					return false;
				else
					return true;
			}

			//Called when results array has changed dimensions.
			public void ResetDimensions(int rows, int columns) {
				RowLast = RowFirst + rows - 1;
				ColumnLast = ColumnFirst + columns - 1;
				Area = rows * columns;
				Resize = 2;
				DoneFlag = 0;
			}

			//Following two methods were made when locks were needed, and are largely trivial
			public bool CheckDimensions(int rows, int columns) {
				return Area == rows * columns;
			}

			//Check Resize against a number
			public bool CheckResize(int num) {
				return num == Resize;
			}

			//Attomically Decrement Resize
			public void DecrementResize() {
				Interlocked.Decrement(ref Resize);
			}
			//Attomically Increment Resize
			public void IncrementResize() {
				Interlocked.Increment(ref Resize);
			}

			//Equals override to check two formula ranges
			public override bool Equals(object obj) {
				//If parameter is null, return false.
				if (obj is null)
					return false;

				return this.GetHashCode() == (obj as FunctionRange).GetHashCode();
			}
			//HashCode as String property
			public string HashCode { get => GetHashCode(RowFirst, RowLast, ColumnFirst, ColumnLast, SheetID); }
			//Hash Code to use for functions
			public override int GetHashCode() {
				return GetHashCode(RowFirst, RowLast, ColumnFirst, ColumnLast, SheetID).GetHashCode();
			}
		}

		public class ExcelTaskObservable<TResult> : IExcelObservable {
            readonly Task<TResult> _task;
            readonly CancellationTokenSource _cts;

            public ExcelTaskObservable(Task<TResult> task)
            {
                _task = task;
            }

            public ExcelTaskObservable(Task<TResult> task, CancellationTokenSource cts)
                : this(task)
            {
                _cts = cts;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                // Start with a disposable that does nothing
                // Possibly set to a CancellationDisposable later
                IDisposable disp = DefaultDisposable.Instance;

                switch (_task.Status)
                {
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
                        if (_cts != null)
                        {
                            var cancelDisp = new CancellationDisposable(_cts);
                            task = _task.ContinueWith(t =>
                            {
                                cancelDisp.SuppressCancel();
                                return t;
                            }).Unwrap();

                            // Then this will be the IDisposable we return from Subscribe
                            disp = cancelDisp;
                        }
                        // And handle the Task completion
                        task.ContinueWith(t =>
                        {
                            switch (t.Status)
                            {
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

		sealed class DefaultDisposable : IDisposable {
			public static readonly DefaultDisposable Instance = new DefaultDisposable();

			// Prevent external instantiation
			DefaultDisposable()
			{}


			public void Dispose() {
				// no op
			}
		}

		sealed class CancellationDisposable : IDisposable {
			bool _suppress;
			readonly CancellationTokenSource _cts;

			public CancellationDisposable(CancellationTokenSource cts)
			{
				_cts = cts ?? throw new ArgumentNullException("cts");
			}

			public CancellationDisposable()
				: this(new CancellationTokenSource())
			{
			}

			public void SuppressCancel()
			{
				_suppress = true;
			}

			public CancellationToken Token
			{
				get { return _cts.Token; }
			}

			public void Dispose()
			{
				if (!_suppress) _cts.Cancel();
				_cts.Dispose();  // Not really needed...
			}
		}


	}
}
