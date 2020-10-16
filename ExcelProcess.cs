using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_To_PDF
{
    class ExcelProcess
    {
        private Excel.Application excelApp = null; 
        private Excel.Workbook workbook = null;
        private Excel.Worksheet worksheet = null;
        private Excel.PageSetup pageSetup = null;
        private Range range = null;
        private Range rangeFinalCol = null;
        private Range rangePathToFolder = null;
        private string defaultPrinter = "";
        private object missing = System.Reflection.Missing.Value;

        private string PathToExcel { get; set; }
        private string PathToExcelNew { get; set; }
        private string PathToFolder { get; set; }
        public ExcelProcess(string _pathToExcel, string _pathToFolder)
        {
            try
            {
                PathToExcel = _pathToExcel;
                PathToFolder = _pathToFolder;
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                //
                if (copyExcel())
                {
                    workbook = excelApp.Workbooks.Open(PathToExcelNew);
                    worksheet = workbook.Worksheets[1];
                    SetupPage();
                    WorkRange();
                    RelaeseMemoryCom();
                    Console.WriteLine("Конвертация завершена");
                }
                else
                    Console.WriteLine("Указанного Excel файла не существует");            
            }
            catch (Exception e)
            {
                RelaeseMemoryCom();
                Console.WriteLine(e.Message);
            }
        }
        public void SetupPage()
        {
            pageSetup = worksheet.PageSetup;
            pageSetup.Orientation = XlPageOrientation.xlLandscape;
            pageSetup.LeftMargin = 0;
            pageSetup.RightMargin = 0;
            pageSetup.Zoom = false;
            pageSetup.FitToPagesTall = 1;
            pageSetup.FitToPagesWide = 1;
            //
            workbook.SaveAs(PathToExcelNew, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        public void WorkRange()
        {
            int row = CountRow();
            int col = CountCol();
            int actualRow = row;
            while (actualRow >= 4)
            {
                int actualCol = FinalCount(actualRow, col);
                //
                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[4, actualCol]];
                SettingBorder();
                range.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, GetPathToFolder(4), XlFixedFormatQuality.xlQualityStandard, true);
                RelaeseMemoryRange();
                range = worksheet.Rows[4, missing];
                range.Delete(XlDeleteShiftDirection.xlShiftUp);
                RelaeseMemoryRange();
                actualRow--;
            }
            //
        }
        public string GetPathToFolder(int actualRow)
        {
            rangePathToFolder = worksheet.Cells[actualRow, 2];
            string nameFile = GenerateRandom();
            if (rangePathToFolder.Value2 != null)
            {
                nameFile = rangePathToFolder.Value2.ToString().Trim();
            }    
            string pathToFolder = $"{PathToFolder}\\{nameFile}.pdf";
            if (rangePathToFolder != null)
            {
                Marshal.ReleaseComObject(rangePathToFolder);
            }
            return pathToFolder;
        }
        public string GenerateRandom()
        {
            Random rnd = new Random();
            int value = rnd.Next();
            return value.ToString();
        }
        public int FinalCount(int row, int col)
        {
            int actualCol = col;
            rangeFinalCol = worksheet.Cells[4, col];
            string finalCol = "";
            if (rangeFinalCol.Value2 != null)
            {
                finalCol = rangeFinalCol.Value2.ToString().Trim();
            }
            if (finalCol == "")
            {
                actualCol = col - 1;
            }
            if (rangeFinalCol != null)
            {
                Marshal.ReleaseComObject(rangeFinalCol);
            }
            return actualCol;
        }
        public void SettingBorder()
        {
            var BordersIndex = XlBordersIndex.xlEdgeBottom;
            range.Borders[BordersIndex].Weight = XlBorderWeight.xlThin;
            range.Borders[BordersIndex].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[BordersIndex].ColorIndex = 0;
        }
        public bool copyExcel()
        {
            if (File.Exists(PathToExcel))
            {
                string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                PathToExcelNew = $"{path}\\copyFile.xlsx";
                FileDelete();
                File.Copy(PathToExcel, PathToExcelNew, true);
                return true;
            }
            return false;
        }

        public int CountRow()
        {
            return worksheet.Cells.Find("*", missing, missing, missing,
                          Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                          false, missing, missing).Row;
        }
        public int CountCol()
        {
            return worksheet.Cells.Find("*", missing, missing, missing,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                false, missing, missing).Column;
        }
        public void FileDelete()
        {
            if (File.Exists(PathToExcelNew))
                File.Delete(PathToExcelNew);
        }
        public void RelaeseMemoryRange()
        {
            if (range != null)
            {
                Marshal.ReleaseComObject(range);
            }
        }
        public void RelaeseMemoryCom()
        {
            if (worksheet != null)
                Marshal.ReleaseComObject(worksheet);
            if (range != null)
                Marshal.ReleaseComObject(range);
            if (rangeFinalCol != null)
                Marshal.ReleaseComObject(rangeFinalCol);
            if (rangePathToFolder != null)
                Marshal.ReleaseComObject(rangePathToFolder);
            if (pageSetup != null)
                Marshal.ReleaseComObject(pageSetup);
            if (workbook != null)
            {
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
            }
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            //FileDelete();
        }
    }
}
