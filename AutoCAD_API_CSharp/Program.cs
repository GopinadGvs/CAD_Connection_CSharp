using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ComponentModel;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;


namespace AutoCADAPI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [CommandMethod("Connection")]
        static void Connection()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            ed.WriteMessage("Welcome to AutoCAD");
        }
    }
}
