using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace SolidWorksAPI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        public static void Main()
        {
            //Get Active Solidworks Session (2022)
            SldWorks sldWorkdApp = Marshal.GetActiveObject("SldWorks.Application.30") as SldWorks;

            sldWorkdApp.SendMsgToUser("Welcome to Solidworks");
        }
    }
}
