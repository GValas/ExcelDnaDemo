using ExcelDna.Integration;

namespace ExcelDnaDemo
{
    public static class ExcelFunctions
    {

        [ExcelFunction(Description = "Monte Carlo Black&Scholes Call price")]
        public static double BsCall(double spot, double vol, double rate, double timeToMaturity, double strike)
        {
            return new BlackScholesHelper().CallPrice(spot, vol, rate, timeToMaturity, strike);
        }
    }
}


/************************************************************
=BsCall(100;0,3;0,08;1;100)
**************************************************************/