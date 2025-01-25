using SolidEdgeFramework;
using System;
using System.Runtime.InteropServices;
using SolidEdgePart;
using SolidEdgeDraft;

namespace SolidEdgeAPI
{
    public class Program
    {
        static void Main(string[] args)
        {

            #region Connect to Session

            Application application = null;

            try
            {
                application = Marshal.GetActiveObject("SolidEdge.Application") as Application;

            }
            catch (Exception)
            {

            }

            if (application == null)
                application = Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application")) as Application;
            application.Visible = true;

            #endregion

            #region Application Properties

            string caption = application.Caption;
            string version = application.Version;
            string userName = application.UserName;

            Documents documents = application.Documents;
            SolidEdgeDocument activeDoc = application.ActiveDocument;

            #endregion

            #region Documents Properties

            string docName = activeDoc.Name;

            string docFullPath = activeDoc.FullName;

            DocumentTypeConstants documentType = application.ActiveDocumentType;

            if (documentType != DocumentTypeConstants.igAssemblyDocument) return;

            //documents.Close();

            #endregion

            #region Create Document

            SolidEdgeDocument solidEdgeDocument1 = documents.Add("SolidEdge.PartDocument");
            //SolidEdgeDocument solidEdgeDocument2 = documents.Add("SolidEdge.AssemblyDocument");
            //SolidEdgeDocument solidEdgeDocument3 = documents.Add("SolidEdge.DraftDocument");

            string SavePath = @"E:\Personal\SolidEdge\Testing\test123.par";

            //solidEdgeDocument1.Save();

            solidEdgeDocument1.SaveAs(SavePath);

            //solidEdgeDocument1.Close();

            //documents.Open(SavePath);

            #endregion

            #region Drawing Document

            DraftDocument draftDocument = activeDoc as DraftDocument;

            Sheet activeSheet = draftDocument.ActiveSheet;

            string sheetName = activeSheet.Name;
            int sheetNumber = activeSheet.Number;

            DrawingViews drawingViews = activeSheet.DrawingViews;

            DrawingView drawingView = drawingViews.Item(1);

            string viewName = drawingView.Name;

            ModelLink modelLink = draftDocument.ModelLinks.Add(SavePath);

            DrawingView drawingView1 = activeSheet.DrawingViews.AddPartView(modelLink, ViewOrientationConstants.igFrontView, 1, 0, 0, PartDrawingViewTypeConstants.sePartDesignedView);

            PartsLists partsLists = draftDocument.PartsLists;

            PartsList partsList = partsLists.Add(drawingView1, "", 1, 1);

            #endregion

        }
    }
}
