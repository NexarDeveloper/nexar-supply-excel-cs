using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using NexarSupplyXll;

namespace NexarSupplyXll
{
    [ComVisible(true)]
    public class NexarSupplyRibbon : ExcelRibbon
    {
        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        /// <summary>
        /// Constructor for the Nexar Supply Ribbon toolbar
        /// </summary>
        public NexarSupplyRibbon()
        { }

        /// <summary>
        /// Formats all Nexar Supply URL queries so that they have a clickable hyperlink
        /// </summary>
        /// <param name="control"></param>
        public void HyperlinkUrlQueries(IRibbonControl control)
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic cellsToCheck = xlApp.ActiveSheet.Cells.SpecialCells(XlCellType.xlCellTypeFormulas);

                if (cellsToCheck != null)
                {
                    Worksheet ws = xlApp.ActiveSheet;

                    foreach (Range cell in cellsToCheck.Cells)
                    {
                        string formula = (string)cell.Formula;
                        if (formula.Contains("NEXAR_SUPPLY_") && formula.Contains("URL"))
                        {
                            if (cell.Value.ToString().Contains("http"))
                            {
                                ws.Hyperlinks.Add(cell, (string)cell.Value, Type.Missing, "Click for more information", Type.Missing);
                            }
                            else if (cell.Hyperlinks.Count > 0)
                            {
                                cell.Hyperlinks.Delete();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }

        /// <summary>
        /// Visit the nexar.com website
        /// </summary>
        /// <param name="control"></param>
        public void VisitNexarQueries(IRibbonControl control)
        {
            try
            {
                Process.Start("https://nexar.com");
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }

        // <summary>
        // Performance a refresh of the 'NEXAR_SUPPLY' queries
        /// </summary>
        /// <param name="control">Ribbon control button</param>
        private void doRefresh(IRibbonControl control, bool forceAll)
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic cellsToCheck = xlApp.ActiveSheet.Cells.SpecialCells(XlCellType.xlCellTypeFormulas);

                if (cellsToCheck != null)
                {
                    if (forceAll)
                    {
                        // This neat trick will 're-formulate' the cell, changing nothing, but causing a refresh!
                    cellsToCheck.Replace(
                        @"NEXAR_SUPPLY_",
                        @"NEXAR_SUPPLY_",
                        XlLookAt.xlPart,
                        XlSearchOrder.xlByRows,
                        true,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);
                }
                    else
                    {
                        // Check each cell's value for error - if so reset the value and This neat trick will 're-formulate' the cell, changing nothing, but causing a refresh!
                        foreach (Range cell in cellsToCheck.Cells)
                        {
                            string formula = (string)cell.Formula;
                            if (formula.Contains("NEXAR_SUPPLY_") && cell.Value.ToString().ToLower().StartsWith("error"))
                            {
                                cell.Value = NexarQueryManager.PROCESSING;
                                if (cell.Hyperlinks.Count > 0)
                                    cell.Hyperlinks.Delete();

                                cell.Formula = formula;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }

        /// <summary>
        /// Refreshes all 'NEXAR_SUPPLY' queries in which there is an error value reported
        /// </summary>
        /// <param name="control">Ribbon control button</param>
        public void RetryErrors(IRibbonControl control)
        {
            doRefresh(control, false);
        }

        /// <summary>
        /// Forces a refresh of all 'NEXAR_SUPPLY' queries
        /// </summary>
        /// <param name="control">Ribbon control button</param>
        public void ForceRefreshAll(IRibbonControl control)
        {
            NexarSupplyAddIn.QueryManager.EmptyQueryCache();
            doRefresh(control, true);
        }

    }
}
