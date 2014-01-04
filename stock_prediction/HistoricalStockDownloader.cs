using System;
using System.Collections.Generic;
using System.Net;

namespace stock_prediction
{
	public class HistoricalStockDownloader
	{
        public static int DAYSINONEYEAR = 365;

		public static List<HistoricalStockRecord> DownloadData(string ticker, int yearToStartFrom, int yearToEnd)
		{
			List<HistoricalStockRecord> records = new List<HistoricalStockRecord>();

			using (WebClient web = new WebClient())
			{
				string data = web.DownloadString(string.Format("http://ichart.finance.yahoo.com/table.csv?s={0}&c={1}", ticker, yearToStartFrom));

                //System.IO.File.WriteAllText(string.Format("raw data.csv"), data);

				data =  data.Replace("r","");

				string[] rows = data.Split(new string[] {"\n", "\r\n"}, StringSplitOptions.RemoveEmptyEntries);

				int tempYear = 0;

				//First row is headers so Ignore it
				for (int i = 1; i < rows.Length; i++)
				{
					if (rows[i].Replace("n","").Trim() == "") continue;

					string[] cols = rows[i].Split(',');

                    DateTime currentDate = Convert.ToDateTime(cols[0]);

                    int currentYear = currentDate.Year;

					HistoricalStockRecord record;

					if (currentYear >= yearToStartFrom && currentYear <= yearToEnd)
					{

                        //create new record when a new year data is coming
						if (tempYear != currentYear)
						{
							tempYear = currentYear;
							record = new HistoricalStockRecord();
							records.Add(record);
                            record.Quotes = new double[checkLeapYear(currentYear) ? DAYSINONEYEAR + 1 : DAYSINONEYEAR];
                            record.Year = currentYear;
						}
						else
						{
							record = records[records.Count - 1];
						}


//						hs.Open = Convert.ToDouble(cols[1]);
//						hs.High = Convert.ToDouble(cols[2]);
//						hs.Low = Convert.ToDouble(cols[3]);
//						hs.Close = Convert.ToDouble(cols[4]);
//						hs.Volume = Convert.ToDouble(cols[5]);
//                      hs.AdjClose = Convert.ToDouble(cols[6]);

                        double value = Convert.ToDouble(cols[6]);
                        record.Quotes[currentDate.DayOfYear-1] = value;

					}


				}

                return fillEmptyQuote(records);
			}
		}

        /*Check if a year is leap year*/
        public static bool checkLeapYear(int year)
        {
            return (year%4 == 0 && year%100 != 0) || (year%400 == 0);
        }

        // fill empty row with quote data from the close day
        private static List<HistoricalStockRecord> fillEmptyQuote(List<HistoricalStockRecord> records)
        {
            double lastQuote;

            if (records.Count > 0)
            {

                foreach (var record in records)
                {
                    lastQuote = 0;

                    for (int i = 0; i < (checkLeapYear(record.Year) ? DAYSINONEYEAR + 1 : DAYSINONEYEAR); i++)
                    {
                        double currentQuote = record.Quotes[i];

                        if (currentQuote != 0)
                        {
                            lastQuote = currentQuote;
                        }
                        else if (currentQuote == 0 && lastQuote != 0)
                        {
                            record.Quotes[i] = lastQuote;
                        }
                        else
                        {
                            int firstValuedQuoteIndex = 0;

                            for (int j = 0; j < (checkLeapYear(record.Year) ? DAYSINONEYEAR + 1 : DAYSINONEYEAR); j++)
                            {
                                if (record.Quotes[j] != 0)
                                {
                                    firstValuedQuoteIndex = j;
                                    break;
                                }
                            }

                            double firstValuedQuoteValue = record.Quotes[firstValuedQuoteIndex];
                            record.Quotes[i] = firstValuedQuoteValue;
                            lastQuote = firstValuedQuoteValue;
                        }

                    }

                }

            }

            return records;
        }
	}
}

