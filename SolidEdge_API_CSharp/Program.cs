using SolidEdgeFramework;
using System;
using System.Runtime.InteropServices;
using SolidEdgePart;
using SolidEdgeDraft;
using SolidEdgeFrameworkSupport;
using OfficeOpenXml;
using System.IO;

namespace SolidEdgeAPI
{
    public class Program
    {
        static void Main(string[] args)
        {
            Application app = new APIMethods().Connect();

            DraftDocument drawing = (DraftDocument)app.ActiveDocument;

            //ModelLink model = drawing.ModelLinks.Add(@"E:\Personal\SolidEdge\Testing\Asm5.asm");

            //DrawingView frontView = drawing.ActiveSheet.DrawingViews.AddPartView(model, ViewOrientationConstants.igFrontView, 1, 0.1, 0.1, PartDrawingViewTypeConstants.sePartDesignedView);

            Sheet activeSheet = drawing.ActiveSheet;

            PartsList partsList = drawing.PartsLists.Item(1);

            TableRows rows = partsList.Rows;

            TableColumns columns = partsList.Columns;


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo("PartsList.xlsx")))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Parts List");

                // Write data rows
                for (int row = 1; row <= rows.Count; row++)
                {
                    for (int col = 1; col <= columns.Count; col++)
                    {
                        TableRow tableRow = rows.Item(row);

                        worksheet.Cells[row + 1, col].Value = rows.Item(row).Cells.Item(col).Value;
                    }
                }

                // Save the Excel file
                package.Save();

            }

            //string name = activeSheet.Name;

            //Line2d circle = activeSheet.Lines2d.Item(1);

            //Circle2d circle2D = activeSheet.Circles2d.Item(1);

            //Dimensions dims = (drawing.ActiveSheet.Dimensions as Dimensions);

            //Dimension diaDimension = dims.AddLength(circle);

            //DrawingView frontView = activeSheet.DrawingViews.Item(1);


            //activeSheet.DrawingViews.AddByFold(frontView, FoldTypeConstants.igFoldRight, 345 / 1000, 260 / 1000);


            ////Dimensions dims = (frontView.Parent as Sheet).Dimensions as Dimensions;


            //DVLine2d line1 = frontView.DVLines2d.Item(1);
            //DVLine2d line2 = frontView.DVLines2d.Item(2);

            ////line1.GetKeyPoint(1, out double x1, out double y1, out double z1, out KeyPointType keypointType1, out int a1);
            ////line1.GetKeyPoint(1, out double x2, out double y2, out double z2, out KeyPointType keypointTyp2, out int a2);

            //Dimension dim = dims.AddDistanceBetweenObjects(line1.Reference, 0, 0, 0, false, line2.Reference, 0, 0, 0, false);

            //dim.GetTextOffsets(out double x1, out double y1);

            //dim.SetTextOffsets(0.04, 0.174);

            //dim.GetTextOffsets(out double x2, out double y2);

            //DrawingView sideView = drawing.ActiveSheet.DrawingViews.AddPartView(model, ViewOrientationConstants.igLeftView, 1, 0.3, 0.5, PartDrawingViewTypeConstants.sePartDesignedView);

            //drawing.SelectSet.Add(frontView);
            //drawing.SelectSet.Add(sideView);

            //drawing.ActiveSheet.DrawingViews.Align();


            //PartsList list = drawing.PartsLists.Add(drawing.ActiveSheet.DrawingViews.Item(1), "", 1, 1);

            //list.GetOrigin(out double x, out double y);

            //list.SetOrigin(0.45, 0.39);

            //SheetWindow sheetWindow = drawing.Windows.Item(1);
            //sheetWindow.FitEx(SheetFitConstants.igFitAll);

        }
    }

    public class APIMethods
    {
        public Application Connect()
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

            return application;
        }
    }

    public class Testing
    {
        public void Test()
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
            SolidEdgeDocument activeDoc = (SolidEdgeDocument)application.ActiveDocument;

            #endregion

            #region Documents Properties

            string docName = activeDoc.Name;

            string docFullPath = activeDoc.FullName;

            DocumentTypeConstants documentType = application.ActiveDocumentType;

            if (documentType != DocumentTypeConstants.igAssemblyDocument) return;

            //documents.Close();

            #endregion

            #region Create Document

            SolidEdgeDocument solidEdgeDocument1 = (SolidEdgeDocument)documents.Add("SolidEdge.PartDocument");
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

            #region PartsList

            PartsList list = draftDocument.PartsLists.Add(drawingViews.Item(1), "", 1, 1);

            list.GetOrigin(out double x, out double y);

            list.SetOrigin(0.45, 0.39);

            #endregion
        }
    }
}
