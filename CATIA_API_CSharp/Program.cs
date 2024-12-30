using System;
using System.Runtime.InteropServices;
using INFITF;
//using MECMOD;
//using PARTITF;

namespace CATIAAPI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        public static void Main()
        {
            //Get Active CATIA Session 

            //Application catiaApp = Marshal.GetActiveObject("CATIA.Application") as Application;

            //int count = catiaApp.Documents.Count;

            // Get or launch CATIA application instance
            Application catiaApp = null;
            try
            {
                // Try to get an active CATIA session
                catiaApp = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
            }
            catch
            {
                Console.WriteLine("No running instance of CATIA found.");
                return;
            }

            if (catiaApp != null)
            {
                Console.WriteLine("Connected to CATIA!");
                // Now you can interact with CATIA documents, parts, etc.
                Documents docs = catiaApp.Documents;

                // For example, open a CATPart
                // docs.Open(@"C:\Path\To\Your\Part.CATPart");

                // Or create a new part
                Document partDoc = docs.Add("Part");
                //Part part = (Part)partDoc.Product;
            }

        }
    }
}
