using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;

namespace stock_prediction
{
    public enum PredictionType
    {
        Sell, Buy, Hold
    }

    public class DataAnalysis
    {
        //add the macro to the excel workbook here.
        public void addMacroToExcel(Excel.Workbook xlWorkBook, int startYear, int endYear)
        {
            Microsoft.Vbe.Interop.VBComponent xlMod;
            xlMod = xlWorkBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
            xlMod.Name = "Module1";

            xlMod.CodeModule.AddFromString(getMacro(startYear, endYear));
        }

        public string getMacro(int startYear, int endYear)
        {
            int yearSpan = endYear - startYear + 1;
            char lastColumnChar = (char)((int)'A' + yearSpan);
            StringBuilder sb = new StringBuilder();

            sb.Append("Sub Auto_Open()" + "\n");
            sb.Append("  Dim SHt As Worksheet " + "\n");
            sb.Append("  For Each SHt In ThisWorkbook.Worksheets " + "\n");
            sb.Append("    SHt.Activate " + "\n");
            sb.Append("    Range(\"A1\").Select " + "\n");
            sb.Append("    Range(\"A1:" + lastColumnChar + "366\").Select " + "\n");
            sb.Append("    ActiveSheet.Shapes.AddChart.Select " + "\n");
            sb.Append("    ActiveChart.ChartType = xlLine " + "\n");
            sb.Append("    ActiveChart.Parent.Height = 593.8897637795 " + "\n");
            sb.Append("    ActiveChart.Parent.Width = 837.3228346457 " + "\n");
            sb.Append("    ActiveChart.Parent.Left = 450 " + "\n");
            sb.Append("    ActiveChart.Parent.Top = 20 " + "\n");
            sb.Append("  Next SHt " + "\n");
            sb.Append("  Sheets(1).Select " + "\n");
            sb.Append("End Sub");
 
            return sb.ToString();
        }
        
    
        public void fillInExcelSheet(string stockCode, List<HistoricalStockRecord> records, Excel.Worksheet oSheet)
        {
            oSheet.Name = stockCode;
            for (int j = 1; j <= records.Count; j++)
            {
                oSheet.Cells[1, j + 1] = records[j-1].Year;
            }

            for (int i = 0; i < HistoricalStockDownloader.DAYSINONEYEAR; i++)
            {
                oSheet.Cells[i + 2, 1] = i + 1;

                for (int j = 1; j <= records.Count; j++)
                {
                    oSheet.Cells[i + 2, j + 1] = records[j - 1].Quotes[i];
                }
            }
        }

        public List<HistoricalStockDerivative> generateStockDerivative(string code, List<HistoricalStockRecord> records, int daysInterval)
        {
            List<HistoricalStockDerivative> stockDerivatives = new List<HistoricalStockDerivative>();
            foreach (var record in records)
            {
                StringBuilder sb = new StringBuilder();

                sb.AppendLine(string.Format("{0},{1}", "DayInYear", "Derivative"));

                HistoricalStockDerivative stockDerivative = new HistoricalStockDerivative();
                stockDerivative.Year = record.Year;
                stockDerivative.Derivatives = new double[HistoricalStockDownloader.checkLeapYear(record.Year) ? HistoricalStockDownloader.DAYSINONEYEAR + 1 : HistoricalStockDownloader.DAYSINONEYEAR];
                stockDerivatives.Add(stockDerivative);

                for (int i = 0; i < record.Quotes.Length - 1; )
                {
                    double date1 = record.Quotes[i];
                    double date2 = record.Quotes[i + daysInterval];

                    double derivative = Math.Derivative(0, (double)daysInterval, date1, date2);
                    stockDerivative.Derivatives[i] = derivative;
                    sb.AppendLine(string.Format("{0},{1}", i + 1, derivative));

                    i = i + daysInterval;
                }
                System.IO.File.WriteAllText(string.Format("{0} - Derivative - {1}.csv", code, record.Year), sb.ToString());
            }
            return stockDerivatives;
        }

        public AvgHistoricalStockRecord generateStockDerivativeAvg(string code, List<HistoricalStockDerivative> derivatives)
        {
            int count = derivatives.Count;

            AvgHistoricalStockRecord avgStockDerivative = new AvgHistoricalStockRecord();
            avgStockDerivative.StartYear = derivatives[count - 1].Year;
            avgStockDerivative.EndYear = derivatives[0].Year;
            avgStockDerivative.AvgQuotes = new double[HistoricalStockDownloader.DAYSINONEYEAR];

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("{0},{1}", "DayInYear", "DerivativeAvg"));

            for (int i = 0; i < HistoricalStockDownloader.DAYSINONEYEAR; i++)
            {
                double sum = 0;
                for (int j = 0; j < count; j++)
                {
                    double derivate = derivatives[j].Derivatives[i];
                    sum += derivate;
                }

                double avg = sum / (avgStockDerivative.EndYear - avgStockDerivative.StartYear + 1);

                sb.AppendLine(string.Format("{0},{1}", i + 1, avg));

                avgStockDerivative.AvgQuotes[i] = avg;
            }

            System.IO.File.WriteAllText(string.Format("{0} - DerivativeAvg - {1} - {2}.csv", code, avgStockDerivative.StartYear, avgStockDerivative.EndYear), sb.ToString());

            return avgStockDerivative;
        }

        public double[] generatePrediction(string code, int daysInterval, AvgHistoricalStockRecord avgStockDerivative)
        {

            int length = avgStockDerivative.AvgQuotes.Length - daysInterval;

            double[] predictionResults = new double[length];

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("{0},{1}", "DayInYear", "Prediction"));

            for (int i = 0; i < length; i++)
            {
                double spanSum = 0.0, spanAvg = 0.0;

                for (int j = 0; j < daysInterval; j++)
                {
                    spanSum += avgStockDerivative.AvgQuotes[i + j];
                }

                spanAvg = spanSum / daysInterval;
                predictionResults[i] = spanAvg;

                sb.AppendLine(string.Format("{0},{1}", i + 1, spanAvg));
            }

            System.IO.File.WriteAllText(string.Format("{0} - {1} - {2} - prediction.csv", code, avgStockDerivative.StartYear, avgStockDerivative.EndYear), sb.ToString());

            return predictionResults;
        }


		

		//generate prediction report to CSV file
		public void generateReport(string code, int yearToStart, int yearToEnd, int yearToPredict, int daysOfAvgDerivativeInterval, int daysThreshold, 
			double valueThreshold, double[] predictionResults, double[] evaluatedData)
		{
			StringBuilder sb = new StringBuilder();
			sb.AppendLine(string.Format("{0},{1},{2},{3}", "Start Year", "End Year", "Prediction Year","Days Interval"));

			PredictionResult result = new PredictionResult();
			result.StartYear = yearToStart;
			result.EndYear = yearToEnd;
			result.PredictedYear = yearToPredict;
			result.DaysInterval = daysOfAvgDerivativeInterval;

			double totalScore = 0;
			PredictionType[] types = new PredictionType[predictionResults.Length];

			sb.AppendLine(string.Format("{0},{1},{2},{3}", result.StartYear, result.EndYear, result.PredictedYear, result.DaysInterval));
			sb.AppendLine(string.Format("{0},{1},{2}", "Day", "Prediction Type", "Evalution Result"));

			for (int i = 0; i < predictionResults.Length; i++)
			{
				PredictionType type = PredictionType.Hold;
				for (int j = 0; j < daysThreshold && i + j < predictionResults.Length; j++)
				{
					double value = predictionResults[i + j];
					PredictionType currentType;

					if (value > valueThreshold)
					{
						currentType = PredictionType.Buy;
					}
					else if (value < valueThreshold)
					{
						currentType = PredictionType.Sell;
					}
					else
					{
						currentType = PredictionType.Hold;
					}


					if (j > 0 && type != currentType)
					{
						type = PredictionType.Hold;
						break;
					}
					else
					{
						type = currentType;
					}

				}

				types[i] = type;
				bool evaluationResult = getEvalutionResult(evaluatedData[i], evaluatedData[i + daysOfAvgDerivativeInterval], type);

				if (evaluationResult)
				{
					totalScore++;
				}

				sb.AppendLine(string.Format("{0},{1},{2}", i + 1, getTypeString(type), evaluationResult));
			}

			sb.AppendLine(string.Format("{0}", totalScore/predictionResults.Length));

			System.IO.File.WriteAllText(string.Format("{0} - {1} - {2} - prediction report.csv", code, yearToStart, yearToEnd), sb.ToString());

		}

		// Convert type to corresponding string
		private string getTypeString(PredictionType type)
		{
			string typeString = "";

			switch (type)
			{
				case PredictionType.Buy:
					typeString = "Buy";
					break;
				case PredictionType.Sell:
					typeString = "Sell";
					break;
				case PredictionType.Hold:
				default:
					typeString = "Hold";
					break;
			}

			return typeString;
		}

		private bool getEvalutionResult(double startDayQuote, double endDayQuote, PredictionType type)
		{
			if (endDayQuote - startDayQuote > 0)
			{
				return type == PredictionType.Buy;
			}
			else if (endDayQuote - startDayQuote < 0)
			{
				return type == PredictionType.Sell;
			}
			else
			{
				return type == PredictionType.Hold;
			}
		}


        //generate prediction report to CSV file
        public void generateReport(string code, int yearToStart, int yearToEnd, int daysThreshold, double valueThreshold, double[] predictionResults)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("{0},{1},{2}", "Start Day", "End Day", "Recommendation Type"));

            for (int i = 0; i < predictionResults.Length; i++)
            {
                PredictionType type = PredictionType.Hold;
                for (int j = 0; j < daysThreshold && i + j < predictionResults.Length; j++)
                {
                    double value = predictionResults[i + j];
                    PredictionType currentType;

                    if (value > valueThreshold)
                    {
                        currentType = PredictionType.Buy;
                    }
                    else if (value < valueThreshold)
                    {
                        currentType = PredictionType.Sell;
                    }
                    else
                    {
                        currentType = PredictionType.Hold;
                    }


                    if (j > 0 && type != currentType)
                    {
                        type = PredictionType.Hold;
                        break;
                    }
                    else
                    {
                        type = currentType;
                    }

                }

                sb.AppendLine(string.Format("{0},{1},{2}", i + 1, i + daysThreshold + 1, getTypeString(type)));
            }

            System.IO.File.WriteAllText(string.Format("{0} - {1} - {2} - prediction report.csv", code, yearToStart, yearToEnd), sb.ToString());

        }

       

        private void convertCSVtoExcel()
        {

            // Create new CSV file.
            //var csvFile = new ExcelFile();

            //// Load data from CSV file.
            //csvFile.LoadCsv(fileName + ".csv", CsvType.CommaDelimited);

            //// Save CSV file to XLS file.
            //csvFile.SaveXls(fileName + ".xls");
            //Excel.Application excel = new Excel.Application();
            //Excel.Workbook workBook = excel.Workbooks.Add();
            //Excel.Worksheet sheet = workBook.ActiveSheet;


        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}

