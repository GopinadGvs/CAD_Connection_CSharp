using System;
using System.Runtime.InteropServices;
using NXOpen;
using NXOpenUI;

namespace NXAPI
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]        

        public static void Main()
        {
            //Get Active NX Session
            Session theSession = Session.GetSession();
            UI uiSession = UI.GetUI();

            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;

            uiSession.NXMessageBox.Show("MyTitle", NXMessageBox.DialogType.Information, "Welcome to NX");            
        }

        public static int GetUnloadOption(string dummy) { return (int)NXOpen.Session.LibraryUnloadOption.Immediately; }
    }
}
