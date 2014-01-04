using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace stock_prediction
{
	public static class Math
    {
        public static double Derivative(double x1, double x2, double y1, double y2)
        {
            double result = (y2 - y1) / (x2 - x1);
            return result;
        }
    }
}
