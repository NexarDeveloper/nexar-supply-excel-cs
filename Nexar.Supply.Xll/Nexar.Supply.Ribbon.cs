using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

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
                        if (formula.Contains("=NEXAR_SUPPLY") && formula.Contains("URL"))
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

        /// <summary>
        /// Forces a refresh of all 'NEXAR_SUPPLY' queries
        /// </summary>
        /// <param name="control">Ribbon control button</param>
        public void RefreshAllQueries(IRibbonControl control)
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic cellsToCheck = xlApp.ActiveSheet.Cells.SpecialCells(XlCellType.xlCellTypeFormulas);

                // This neat trick will 're-formulate' the cell, changing nothing, but causing a refresh!
                if (cellsToCheck != null)
                {
                    cellsToCheck.Replace(
                        @"=NEXAR_SUPPLY_",
                        @"=NEXAR_SUPPLY_",
                        XlLookAt.xlPart,
                        XlSearchOrder.xlByRows,
                        true,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }
    }
}
