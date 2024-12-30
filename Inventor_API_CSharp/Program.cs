using System;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace InvnetorAPI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            //Get Active Inventor Session
            Inventor.Application application = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;

            MessageBox.Show("Welcome to Inventor", "MyTitle");
        }
    }
}
