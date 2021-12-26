using ExcelDna.Integration;

namespace ExcelDnaDemo
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object Clock(int refreshRate = 2000)
        {
            // Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
            // Note that the topic information needs at least one string - it's not used in this sample
            return XlCall.RTD(RtdClockServer.ServerProgId, null, refreshRate.ToString());
        }
    }
}

/*************************************
 =Clock(1000)
**************************************/
