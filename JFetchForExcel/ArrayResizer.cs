using ExcelDna.Integration;
using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Timers;

namespace JFetchForExcel {
	public static class ArrayResizer
    {
        private static ConcurrentQueue<ExcelReference> ResizeJobs = new ConcurrentQueue<ExcelReference>();

        private static int TimerInit = 0;

        // Will run on a ThreadPool thread
        private static void TimerElapsed(object sender, ElapsedEventArgs e) => DoResizing();

        //Schedules the resize macro
        private static System.Timers.Timer BatchTimer = new System.Timers.Timer()
        {
            AutoReset = false,
            SynchronizingObject = null,
            Interval = TimeSpan.FromMilliseconds(500).TotalMilliseconds           
        };

        //Add array and caller to the Resize Queue
        public static object Resize(object[,] array, ExcelReference caller)
        {
            if (caller == null)
                return array;

            int rows = array.GetLength(0);
            int columns = array.GetLength(1);

            //Check for Size problem: enqueue job, call async update and return #N/A
            if ((caller.RowLast - caller.RowFirst + 1 != rows) || (caller.ColumnLast - caller.ColumnFirst + 1 != columns))
                EnqueueResize(caller, rows, columns);

            //Size is already OK - just return result
            return array;
        }

        private static void EnqueueResize(ExcelReference caller, int rows, int columns)
        {
            //Add new ExcelReference to the Queue to be resized
            ResizeJobs.Enqueue(new ExcelReference(caller.RowFirst, caller.RowFirst + rows - 1, caller.ColumnFirst, caller.ColumnFirst + columns - 1, caller.SheetId));

            if(Interlocked.CompareExchange(ref TimerInit,0,1) == 0)
            {
                BatchTimer.Elapsed += TimerElapsed;
            }

            if (!BatchTimer.Enabled)
                BatchTimer.Start();
        }


        private static void DoResizing()
        {
            ConcurrentQueue<ExcelReference> _batch = Interlocked.Exchange(ref ResizeJobs, new ConcurrentQueue<ExcelReference>());

            ExcelAsyncUtil.QueueAsMacro(delegate (object state)
            {
                while (_batch.Count > 0)
                {
                    bool good = ((ConcurrentQueue<ExcelReference>)state).TryDequeue(out ExcelReference _local);
                    if (good && _local != null)
                        DoResize(_local);
                }

            }, _batch);
        }

        public static void DoResize(ExcelReference target)
        {
            // Get the current state for reset later
            object oldCalculationMode = null;
            object oldEcho = null;
            try
            {
                //Set up environment
                oldEcho = XlCall.Excel(XlCall.xlfGetWorkspace, 40);
                XlCall.Excel(XlCall.xlcEcho, false);
                oldCalculationMode = XlCall.Excel(XlCall.xlfGetDocument, 14);
                XlCall.Excel(XlCall.xlcOptionsCalculation, 3);

                //Start of DoResize
                ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);

                // Get the formula in the first cell of the target
                string formula = (string)XlCall.Excel(XlCall.xlfGetCell, 41, firstCell);
                bool isFormulaArray = (bool)XlCall.Excel(XlCall.xlfGetCell, 49, firstCell);
                if (isFormulaArray)
                {
                    // Select the sheet and firstCell - needed because we want to use SelectSpecial.
                    object oldSelectionOnActiveSheet = null;
                    object oldActiveCellOnActiveSheet = null;

                    object oldSelectionOnRefSheet = null;
                    object oldActiveCellOnRefSheet = null;
                    try
                    {
                        // Remember old selection state on the active sheet
                        oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);
                        oldActiveCellOnActiveSheet = XlCall.Excel(XlCall.xlfActiveCell);
                        // Switch to the sheet we want to select
                        string refSheet = (string)XlCall.Excel(XlCall.xlSheetNm, firstCell);
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { refSheet }); //---------------------------------------------------------
                        // record selection and active cell on the sheet we want to select
                        oldSelectionOnRefSheet = XlCall.Excel(XlCall.xlfSelection);
                        oldActiveCellOnRefSheet = XlCall.Excel(XlCall.xlfActiveCell);
                        // make the selection
                        XlCall.Excel(XlCall.xlcFormulaGoto, firstCell);

                        // Extend the selection to the whole array and clear
                        XlCall.Excel(XlCall.xlcSelectSpecial, 6);
                        ExcelReference oldArray = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

                        oldArray.SetValue(ExcelEmpty.Value);
                    }
                    catch (Exception ex) { /*ErrorLog.AppendException(ex);*/ }
                    finally
                    {
                        // Reset the selection on the target sheet
                        XlCall.Excel(XlCall.xlcSelect, oldSelectionOnRefSheet, oldActiveCellOnRefSheet);
                        // Reset the sheet originally selected
                        string oldActiveSheet = (string)XlCall.Excel(XlCall.xlSheetNm, oldSelectionOnActiveSheet); //----------------------------------
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { oldActiveSheet });
                        // Reset the selection in the active sheet (some bugs make this change sometimes too)
                        XlCall.Excel(XlCall.xlcSelect, oldSelectionOnActiveSheet, oldActiveCellOnActiveSheet);
                    }

                }
                // Get the formula and convert to R1C1 mode
                bool isR1C1Mode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
                string formulaR1C1 = formula;
                if (!isR1C1Mode)
                {
                    object formulaR1C1Obj;
                    XlCall.XlReturn formulaR1C1Return = XlCall.TryExcel(XlCall.xlfFormulaConvert, out formulaR1C1Obj, formula, true, false, ExcelMissing.Value, firstCell);
                    if (formulaR1C1Return != XlCall.XlReturn.XlReturnSuccess || formulaR1C1Obj is ExcelError)
                    {
                        string firstCellAddress = (string)XlCall.Excel(XlCall.xlfReftext, firstCell, true);
                        XlCall.Excel(XlCall.xlcAlert, "Cannot resize array formula at " + firstCellAddress + " - formula might be too long when converted to R1C1 format.");
                        firstCell.SetValue("'" + formula);
                        return;
                    }
                    formulaR1C1 = (string)formulaR1C1Obj;
                }
                // Must be R1C1-style references
                object ignoredResult;
                XlCall.XlReturn formulaArrayReturn = XlCall.TryExcel(XlCall.xlcFormulaArray, out ignoredResult, formulaR1C1, target);

                // TODO: Find some dummy macro to clear the undo stack
                if (formulaArrayReturn != XlCall.XlReturn.XlReturnSuccess)
                {
                    string firstCellAddress = (string)XlCall.Excel(XlCall.xlfReftext, firstCell, true);
                    XlCall.Excel(XlCall.xlcAlert, "Cannot resize array formula at " + firstCellAddress + " - result might overlap another array.");
                    // Might have failed due to array in the way.
                    firstCell.SetValue("'" + formula);
                }
            }
            catch (Exception ex) { /*ErrorLog.AppendException(ex);*/ } //MessageBox.Show("Error in array-resize\n" + ex.ToString());
            finally
            {
                if (oldCalculationMode != null)
                    XlCall.Excel(XlCall.xlcOptionsCalculation, oldCalculationMode);
                if (oldEcho != null)
                    XlCall.Excel(XlCall.xlcEcho, oldEcho);
            }
        }

        //Old Depreciated
        static void ___DoResize(ExcelReference target)
        {
            object oldEcho = XlCall.Excel(XlCall.xlfGetWorkspace, 40);
            object oldCalculationMode = XlCall.Excel(XlCall.xlfGetDocument, 14);
            try
            {
                // Get the current state for reset later
                XlCall.Excel(XlCall.xlcEcho, false);
                XlCall.Excel(XlCall.xlcOptionsCalculation, 3);

                // Get the formula in the first cell of the target
                string formula = (string)XlCall.Excel(XlCall.xlfGetCell, 41, target);
                ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);

                bool isFormulaArray = (bool)XlCall.Excel(XlCall.xlfGetCell, 49, target);
                if (isFormulaArray)
                {
                    object oldSelectionOnActiveSheet = null;
                    object oldActiveCellOnActiveSheet = null;

                    object oldSelectionOnRefSheet = null;
                    object oldActiveCellOnRefSheet = null;
                    try
                    {
                        // Select First_Cell-----------------------------------------------
                        // Remember old selection state on the active sheet
                        oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfSelection);
                        oldActiveCellOnActiveSheet = XlCall.Excel(XlCall.xlfActiveCell);
                        // Switch to the sheet we want to select
                        string refSheet = (string)XlCall.Excel(XlCall.xlSheetNm, firstCell);
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { refSheet });
                        // record selection and active cell on the sheet we want to select
                        oldSelectionOnRefSheet = XlCall.Excel(XlCall.xlfSelection);
                        oldActiveCellOnRefSheet = XlCall.Excel(XlCall.xlfActiveCell);
                        // make the selection
                        XlCall.Excel(XlCall.xlcFormulaGoto, firstCell);
                        //----------------------------------------------------------------
                        // Extend the selection to the whole array and clear
                        XlCall.Excel(XlCall.xlcSelectSpecial, 6);
                        ExcelReference oldArray = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

                        oldArray.SetValue(ExcelEmpty.Value);
                    }
                    catch (Exception ex) { /*ErrorLog.AppendException(ex);*/ } //MessageBox.Show("Error in selection change-array-resize\n" + ex.ToString());
                    finally
                    {
                        // Reset the selection on the target sheet
                        XlCall.Excel(XlCall.xlcSelect, oldSelectionOnRefSheet, oldActiveCellOnRefSheet);
                        // Reset the sheet originally selected
                        string oldActiveSheet = (string)XlCall.Excel(XlCall.xlSheetNm, oldSelectionOnActiveSheet);                     
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { oldActiveSheet });
                        // Reset the selection in the active sheet (some bugs make this change sometimes too)
                        XlCall.Excel(XlCall.xlcSelect, oldSelectionOnActiveSheet, oldActiveCellOnActiveSheet);
                    }

                }
                // Get the formula and convert to R1C1 mode
                bool isR1C1Mode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
                string formulaR1C1 = formula;
                if (!isR1C1Mode)
                {
                    // Set the formula into the whole target
                    formulaR1C1 = (string)XlCall.Excel(XlCall.xlfFormulaConvert, formula, true, false, ExcelMissing.Value, firstCell);
                }
                // Must be R1C1-style references
                object ignoredResult;
                //Debug.Print("Resizing START: " + target.RowLast);
                XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlcFormulaArray, out ignoredResult, formulaR1C1, target);
                //Debug.Print("Resizing FINISH");
                // TODO: Dummy action to clear the undo stack

                if (retval != XlCall.XlReturn.XlReturnSuccess)
                {
                    // TODO: Consider what to do now!?
                    // Might have failed due to array in the way.
                    firstCell.SetValue("'" + formula);
                }
            }
            finally
            {
                XlCall.Excel(XlCall.xlcEcho, oldEcho);
                XlCall.Excel(XlCall.xlcOptionsCalculation, oldCalculationMode);
            }
        }
    }
}