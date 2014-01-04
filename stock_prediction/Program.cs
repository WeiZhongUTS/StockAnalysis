using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace stock_prediction
{
	class MainClass
	{
		public static void Main (string[] args)
		{  
			int yearToStart = 2009;
            int yearToEnd = 2013;

            Excel.Application oApp = new Excel.Application();
            Excel.Workbook oBook;
            object misValue = System.Reflection.Missing.Value;
            oBook = oApp.Workbooks.Add(misValue);
            
            StreamReader reader = new StreamReader(File.OpenRead(@".\stock_codes.csv"));

            int sheetIndex = 1;
            DataAnalysis dataAnalysis = new DataAnalysis();
            while (!reader.EndOfStream)
            {
                Console.WriteLine(sheetIndex);
                try
                {
                    Excel.Worksheet oSheet;

                    if (sheetIndex > 2)
                    {
                        var xlSheets = oBook.Sheets as Excel.Sheets;
                        var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetIndex], Type.Missing, Type.Missing, Type.Missing);
                    }

                    oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(sheetIndex);
                    string stockCode = reader.ReadLine();

                    Console.WriteLine("Downloading " + stockCode);
                    List<HistoricalStockRecord> data = HistoricalStockDownloader.DownloadData(stockCode, yearToStart, yearToEnd);
                    Console.WriteLine("Saving " + stockCode);
                    dataAnalysis.fillInExcelSheet(stockCode, data, oSheet);

                    sheetIndex++;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            dataAnalysis.addMacroToExcel(oBook, yearToStart, yearToEnd);

            oBook.SaveAs(string.Format("{0} - {1}.xlsm", yearToStart, yearToEnd), Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, false, false, null, null, null);

            oBook.Close(true, misValue, misValue);
            oApp.Quit();
		}
	}
}
