using ExcelDna.Integration;
using JFetchUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OrionApiSdk.Utils
{
    /// <summary>
    /// Factory Class for the Excel Asynchronous Array Function Returned through the RTD Server.
    /// <summary>
    public class FunctionFactory
    {
        //Named Tuple for the results of the generator method
        public class ResizeResult
        {
            public bool Resize { get; set; }
            public object[,] Result { get; set; }

            public ResizeResult(bool r, object[,] res)
            {
                Resize = r;
                Result = res;
			}
		}

		//Object Representing the Excel Range of the Array Function, that performs record keeping
		public class FunctionRange
        {
            //Function to get hashcode of a FunctionRange
            public static string GetHashCode(int row_first,int row_last, int col_first, int col_last, int sheetId)
            {
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
            public FunctionRange(ExcelReference caller)
            {
                RowFirst = caller.RowFirst;
                RowLast = caller.RowLast;
                ColumnFirst = caller.ColumnFirst;
                ColumnLast = caller.ColumnLast;
                Area = (caller.RowLast - caller.RowFirst + 1) * (caller.ColumnLast - caller.ColumnFirst + 1);
                SheetID = caller.SheetId.ToInt32();

                Resize = 2;
            }

            //Test if caller is contained by this Range
            public bool Contains(ExcelReference caller)
            {
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
            public bool Contains(FunctionRange caller)
            {
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
            public void ResetDimensions(int rows, int columns)
            {
                RowLast = RowFirst + rows - 1;
                ColumnLast = ColumnFirst + columns - 1;
                Area = rows * columns;
                Resize = 2;
                DoneFlag = 0;
            }

            //Following two methods were made when locks were needed, and are largely trivial
            public bool CheckDimensions(int rows, int columns)
            {
                return Area == rows * columns;
            }

            //Check Resize against a number
            public bool CheckResize(int num)
            {
                return num == Resize;
            }

            //Attomically Decrement Resize
            public void DecrementResize()
            {
                Interlocked.Decrement(ref Resize);
            }
            //Attomically Increment Resize
            public void IncrementResize()
            {
                Interlocked.Increment(ref Resize);
            }       

            //Equals override to check two formula ranges
            public override bool Equals(object obj)
            {
                //If parameter is null, return false.
                if (obj is null)
                    return false;

                return this.GetHashCode() == (obj as FunctionRange).GetHashCode();
            }
            //HashCode as String property
            public string HashCode { get => GetHashCode(RowFirst, RowLast, ColumnFirst, ColumnLast, SheetID); }
            //Hash Code to use for functions
            public override int GetHashCode()
            {
                return GetHashCode(RowFirst,RowLast,ColumnFirst, ColumnLast, SheetID).GetHashCode();
            }
        }

        //Represents the Actual Function itself
        public class Function 
        {            
            public object[,] EmptyArray { get; set; }  //Empty Array, used for null results, or errors
            private NonLockingHashMap<string, FunctionRange> FunctionCounters { get; set; } //Map that contains all the FormulaRanges
            private NonLockingHashMap<string, object[,]> ResultCache { get; set; } //Map of all proccessed Array Results
            private NonLockingHashMap<string, int> FunctionLengths = new NonLockingHashMap<string, int>();

            private readonly Func<object[], Task<object[,]>> BatchRunner; //Function that makes the request
            private readonly Func<object[], string> RawHashFunction; //Hash function that hashes the <param object[] args>

            /// <summary>
            /// Wraps a <see cref="Func{Object[],Task{object[,]}}"/> and hashes the result with a <see cref="Func{object[],string}"/>.
            /// This class acts as a generator for the Do Resize Macro.
            /// </summary>
            /// <param name="batchRunner">The <see cref="Func{object[],Task{object[,]}}"/> to run on each batch.</param>
            /// <param name="hashFunction">The paramter hashing function returning a <see cref="string"/></param>
            /// <param name="emptyArray">The defualt <see cref="object[,]"/> return value when the request is <c>null</c> <c>0</c> sized.</param>
            public Function(Func<object[], Task<object[,]>> batchRunner,Func<object[], string> hashFunction, object[,] emptyArray = null)
            {
                BatchRunner = batchRunner;
                EmptyArray = emptyArray ?? new object[,] { { "#ORION N/A" } };
                RawHashFunction = hashFunction;
                FunctionCounters = new NonLockingHashMap<string, FunctionRange>();
                ResultCache = new NonLockingHashMap<string, object[,]>();
            }

            //Hash Function that provides the topic string for the RTD Call. If the topic string has changed, then the current time
            //is added to hash to bypass the RTD Marshall Cache, and rerun the <BatchRunner> in order to request a resize.
            private string HashFunction(ExcelReference caller, object[] args)
            {
                string hash = RawHashFunction(args);

                //Add Pending Result to Hash Length Table
                //AddLengthResult(hash, -3);

                //Lookup caller in Function Counter Cache
                if (TryLookup(caller, out FunctionRange counter_range))
                {
                    //Hash Key has changed or, the DoneFlag is set, in either case, run <BatchRunner>
                    if (!hash.Equals(counter_range.LastKey) || counter_range.DoneFlag == 1)
                    {
                        //Reset Done Flag, so it will be Resized
                        if (counter_range.DoneFlag == 1)
                            counter_range.DoneFlag = 0;

                        counter_range.LastKey = hash;
                        hash += DateTime.Now.GetHashCode().ToString();
                    }
                }

                return hash;
            }

            //Main Public method. Similar to behavior of AsyncBatch.Run
            public object Run(string functionName, ExcelReference caller, params object[] args)
            {
                //Return an ExcelObservable that is completed with the RTD Server
                return ExcelAsyncUtil.Observe(functionName, HashFunction(caller, args), delegate
                {
                    TaskCompletionSource<object> tcs = new TaskCompletionSource<object>();
                    // Start a background task that will complete tcs.Task
                    Task.Factory.StartNew(async delegate 
                    {
                        try
                        {                           
                            tcs.SetResult(await AsyncQueryTopic(caller, args).ConfigureAwait(false)); //Await batch Runner
                        }
                        catch (Exception ex)
                        {
                            tcs.SetException(ex);
                        }

                    }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);

                    return new AsyncTaskUtils.ExcelTaskObservable<object>(tcs.Task);
                });
            }
         
            //Calculate Whether or not to run the resize macro for this particular calller
            private async Task<object[,]> AsyncQueryTopic(ExcelReference caller, object[] args)
            {
                //Use hash function to search ResultCache for previous result
                string key = RawHashFunction(args);

				//If key is not found, or found value is the empty array, make request again
				if (!ResultCache.TryGetValue(key, out object[,] result) || IsEmptyArray(result)) {
					result = await BatchRunner(args).ConfigureAwait(false);
					result = result ?? EmptyArray;

					//Add Length of result to storage dict
					//AddLengthResult(key, result.GetLength(0));

					//Add result to ResultCache or, add EmptyArray if the BatchRunner returned null
					if (!ResultCache.ContainsKey(key))
                        ResultCache.TryAdd(key, result);
                }

                //Calculate whether or not to resize
                ResizeResult _ret = CalculateResize(caller, result);

                //If a resize has been calculated, then run DoResize macro
                return (_ret.Resize) ? (object[,])JFetch.ArrayResizer.Resize(_ret.Result, caller) : _ret.Result;
            }        
            
            //Calculate whether or not the Function Range is due for a Resize
            private ResizeResult CalculateResize(ExcelReference caller, object[,] result)
            {
                //If the array is null, then use the empty_array
                if (result == null)
                    result = EmptyArray;

                //Size of the actual array function results
                int rows = result.GetLength(0);
                int columns = result.GetLength(1);

                //Search Hash in function counters
                if (TryLookup(caller,out FunctionRange counter_range))
                {                               
                    bool _resize;
                    //If result has changed size, request a resize
                    if (_resize = !counter_range.CheckDimensions(rows, columns))
                    {
                        counter_range.ResetDimensions(rows, columns);
                        //counter_range.DecrementResize();
                    }                                      
                    else   
                        counter_range.DecrementResize();                                        

                    //If Resize is equal to one, that means a resize has occured and the done flag should be set
                    if (counter_range.CheckResize(1))
                        counter_range.DoneFlag = 1;

                    return new ResizeResult(_resize, result);
                }
                else
                {
                    counter_range = new FunctionRange(caller);
                    counter_range.ResetDimensions(rows, columns);
                    
                    //Add to function counter Map
                    FunctionCounters.TryAdd(counter_range.HashCode, counter_range);

                    return new ResizeResult(false, result);
                }
            }

            //Check if all Formula Ranges for a workseet have a set done flag
            public int FindPendingRanges(int index)
            {
                int pending = 0;
                List<string> _keys = FunctionCounters.Keys.ToList();
                for (int i = 0; i < _keys.Count; i++)
                {
                    if (FunctionCounters[_keys[i]].SheetID == index && FunctionCounters[_keys[i]].DoneFlag != 1)
                        pending++;
                }

                return pending;
            }

            //Following two methods respectively store and add topic array lengths
            private void AddLengthResult(string func_hash, int _len)
            {
                if (func_hash == null || func_hash == "" || !(func_hash.Length > 0))
                    return;

                if (FunctionLengths.TryGetValue(func_hash, out int _entry))
                {
                    if (_len > _entry)
                        FunctionLengths.TryUpdate(func_hash, _len,_entry);
                }
                else
                    FunctionLengths.TryAdd(func_hash, _len);
            }

            public int GetTopicLength(string topic)
            {
                if (FunctionLengths.TryGetValue(topic, out int length))
                    return length;
                else
                    return -1;
            }

            //Helper method to lookup the caller ExcelReference in the Function Counter Map
            private bool TryLookup(ExcelReference caller, out FunctionRange range)
            {
                List<string> _keys = FunctionCounters.Keys.ToList();
                for (int i = 0; i < _keys.Count; i++)
                {
                    if (FunctionCounters[_keys[i]].Contains(caller))
                    {
                        range = FunctionCounters[_keys[i]];
                        return true;
                    }
                }

                range = null;
                return false;
            }

            //Helper method to test array against empty array
            private bool IsEmptyArray(object[,] arr)
            {
                if (arr == null)
                    return false;

                if (!(arr.GetLength(0) == EmptyArray.GetLength(0) && arr.GetLength(1) == EmptyArray.GetLength(1)))
                    return false;

                for (int r = 0; r < arr.GetLength(0); r++)
                {
                    for (int c = 0; c < arr.GetLength(1); c++)
                    {
                        if (!arr[r, c].Equals(EmptyArray[r, c]))
                            return false;
                    }
                }

                return true;
            }
        }
    }
}