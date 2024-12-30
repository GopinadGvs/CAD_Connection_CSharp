using System;
using pfcls;

namespace testCase
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            //Get Active CREO Session
            CCpfcAsyncConnection classAsyncConnection = new CCpfcAsyncConnection();
            IpfcAsyncConnection connection = classAsyncConnection.Connect(null, null, null, null);
            IpfcBaseSession session = (IpfcBaseSession)connection.Session;

            //Display Message
            (session as IpfcSession).UIShowMessageDialog("Welcome to CREO", null);
        }
    }
}
