using System;
using System.Collections.Generic;

namespace stock_prediction
{
	public class HistoricalStock
	{
			public DateTime Date { get; set; }
			public double Open { get; set; }
			public double High { get; set; }
			public double Low { get; set; }
			public double Close { get; set; }
			public double Volume { get; set; }
			public double AdjClose { get; set; }
	}

	public class HistoricalStockRecord
	{
		public int Year { get; set;}
        public double[] Quotes { get; set; }
	}

    public class AvgHistoricalStockRecord
    {
        public int StartYear { get; set;}
        public int EndYear { get; set;}
        public double[] AvgQuotes { get; set; }
    }

    public class HistoricalStockDerivative
    {
        public int Year { get; set; }
        public double[] Derivatives { get; set; }
    }

	public class PredictionResult
	{
		public int StartYear { get; set;}
		public int EndYear { get; set;}
		public int PredictedYear { get; set;}
		public int DaysInterval { get; set; }
		public double Accuracy { get; set; }
		public PredictionType[] records { get; set; }
	}

	public class PredictionResultRecord
	{
		PredictionType Type { get; set; }
		public bool EvaluationResult { get; set; }
	}
}

