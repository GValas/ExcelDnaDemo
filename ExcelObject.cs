using ExcelDna.Integration;
using ExcelDna.ComInterop;
using System.Runtime.InteropServices; 
//
namespace ExcelDnaDemo
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("ExcelDnaDemo")]
    public class InterfaceFunctions
    {
        public double MclBsCall(double spot, double vol, double rate, double timeToMaturity, double strike)
        {
            return new BlackScholesHelper().CallPrice(spot, vol, rate, timeToMaturity, strike);
        }
    }
    //
    [ComVisible(false)]
    class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }
}

/************************************************************
  
Option Explicit
Sub tester()
    Dim lib As Object
    Set lib = CreateObject("ExcelDnaDemo")
    Debug.Print lib.MclBsCall(100, 0.3, 0.08, 1, 100)
    Set lib = Nothing
End Sub

**************************************************************/