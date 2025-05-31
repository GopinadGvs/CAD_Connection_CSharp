using System;
using System.Runtime.InteropServices;
using NXOpen;
using NXOpen.Assemblies;
using NXOpen.Features;
using NXOpenUI;
using NXOpen.UF;
using NXOpen.Drawings;
using static NXOpen.CAE.Post;
using System.ComponentModel.Design;
//using NXOpen.PDM;
//using NXOpen.CAE;

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
			
			//PdmSession pdmSession = theSession.PdmSession;
            //pdmSession.GetCheckedoutStatusOfAllObjectsInSession()

            //PdmPart pdmPart = pdmSession.getpdmpar            

            uiSession.NXMessageBox.Show("MyTitle", NXMessageBox.DialogType.Information, "Welcome to NX");

            UFSession uFSession = UFSession.GetUFSession();

            uFSession.Curve.AskArcData();
            uFSession.Curve.AskLineData();
            uFSession.Curve.AskCentroid();

            uFSession.Obj.SetLayer();
            uFSession.Obj.SetName();
            uFSession.Obj.AskTypeAndSubtype();

            uFSession.Disp.SetHighlight();
            uFSession.Disp.SetHighlights();

            uFSession.View.AskVisibleObjects();
            uFSession.View.AskZoomScale();

            uFSession.Part.Open();
            uFSession.Part.AskDisplayPart();

            uFSession.Assem.AddPartToAssembly();
            uFSession.Assem.AskComponentData();
            uFSession.Assem.AskCompPosition();

            uFSession.Draw.AddOrthographicView();
            uFSession.Draw.AskBoundaryCurves();

            uFSession.Drf.AskBoundaries();


            //uFSession.Assem.SuppressArray();

            //uFSession.Obj.AskTypeAndSubtype();

            //uFSession.Disp.SetHighlight();

            // Define what types of NX objects can be selected
            Selection.MaskTriple[] maskTriples = new Selection.MaskTriple[2];

            //// First mask: Faces
            maskTriples[0] = new Selection.MaskTriple();
            maskTriples[0].Type = UFConstants.UF_solid_type;
            maskTriples[0].Subtype = UFConstants.UF_solid_face_subtype;
            maskTriples[0].SolidBodySubtype = 0; // 0 means all types of faces


            //uiSession.SelectionManager.SetSelectionMask()

            ListingWindow ls = theSession.ListingWindow;

            Part workPart = theSession.Parts.Work;
            Part displayPart = theSession.Parts.Display;

            ComponentAssembly assy = workPart.ComponentAssembly;

            Component component1 = assy.RootComponent;
            //component1.Name;

            BasePart part = assy.OwningPart;
            Part prt1 = component1.Prototype as Part;

            component1.GetChildren();

            //assy.OwningComponent;

            Face face = null;
            Component comp = face.OwningComponent;


            //uiSession.SelectionManager.SelectTaggedObject("","",Selection.SelectionScope.WorkPart,false, Selection.SelectionType.,
            //uiSession.SelectionManager.SelectTaggedObjects("","",Selection.SelectionScope.WorkPart,false, false,out )


            Component component = null;
            component.Suppress();
           

            ModelingViewCollection modelingViewCollection = workPart.ModelingViews;

            modelingViewCollection.WorkView.Fit();
            workPart.ModelingViews.WorkView.Orient(NXOpen.View.Canned.Isometric, NXOpen.View.ScaleAdjustment.Fit);
            //workPart.ModelingViews.WorkView.RenderingStyle = NXOpen.View.RenderingStyleType.StaticWireframe;

            Matrix3x3 viewMatrix = workPart.ModelingViews.WorkView.Matrix;

            //NXOpen.Matrix3x3 rotMatrix1 = new NXOpen.Matrix3x3();
            //rotMatrix1.Xx = 1;
            //rotMatrix1.Xy = 0;
            //rotMatrix1.Xz = 0;
            //rotMatrix1.Yx = 0;
            //rotMatrix1.Yy = 1;
            //rotMatrix1.Yz = 0;
            //rotMatrix1.Zx = 0;
            //rotMatrix1.Zy = 0;
            //rotMatrix1.Zz = 1;


            //// Convert degrees to radians
            //double radians = 15 * Math.PI / 180.0;

            //// Create rotation matrix around Z-axis
            //Matrix3x3 rotationMatrix = new Matrix3x3();
            //rotationMatrix.Xx = Math.Cos(radians);
            //rotationMatrix.Xy = -Math.Sin(radians);
            //rotationMatrix.Xz = 0.0;

            //rotationMatrix.Yx = Math.Sin(radians);
            //rotationMatrix.Yy = Math.Cos(radians);
            //rotationMatrix.Yz = 0.0;

            //rotationMatrix.Zx = 0.0;
            //rotationMatrix.Zy = 0.0;
            //rotationMatrix.Zz = 1.0;



            //workPart.ModelingViews.WorkView.Orient(rotationMatrix);

            //uFSession.View.AskVisibleObjects(orientedView.Tag, out int visiblecount, out Tag[] visibleObjects, out int clippedCount, out Tag[] clippedObjects);

            View orientedView = workPart.ModelingViews.WorkView;

            DisplayableObject[] displayableObjects = orientedView.AskVisibleObjects();

            //displayableObjects[0].OwningComponent;

            theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");


            DraftingDrawingSheet drawingSheet = null;
            DraftingDrawingSheetBuilder draftingDrawingSheetBuilder = workPart.DraftingDrawingSheets.CreateDraftingDrawingSheetBuilder(drawingSheet);
            draftingDrawingSheetBuilder.Number = "1";
            draftingDrawingSheetBuilder.MetricSheetTemplateLocation = "F:\\Program Files\\NX1953\\DRAFTING\\templates\\Drawing-A0-Size2D-template.prt";


            NXOpen.NXObject nXObject2;
            nXObject2 = draftingDrawingSheetBuilder.Commit();
            draftingDrawingSheetBuilder.Destroy();

            NXOpen.Drawings.BaseView nullNXOpen_Drawings_BaseView = null;
            NXOpen.Drawings.BaseViewBuilder baseViewBuilder1;
            baseViewBuilder1 = workPart.DraftingViews.CreateBaseViewBuilder(nullNXOpen_Drawings_BaseView);

            baseViewBuilder1.Placement.Associative = true;
            baseViewBuilder1.SelectModelView.SelectedView = (ModelingView)orientedView;
            baseViewBuilder1.Scale.Numerator = 1;
            baseViewBuilder1.Scale.Denominator = 5;

            NXOpen.Point3d point1 = new NXOpen.Point3d(456.19106706323305, 426.82189129726197, 0.0);
            //baseViewBuilder1.Placement.Placement.SetValue(null, workPart.Views.WorkView, point1);
            baseViewBuilder1.Placement.Placement.SetValue(null, orientedView, point1);

            NXOpen.NXObject nXObject3;
            nXObject3 = baseViewBuilder1.Commit();
            baseViewBuilder1.Destroy();

            TaggedObject taggedObject = NXOpen.Utilities.NXObjectManager.Get(nXObject3.Tag);
            DraftingView draftingView = (taggedObject as DraftingView);

            //uFSession.Draw.AskBoundaryCurves();

            //draftingView.SetScale()

            //draftingView.Update();

            nullNXOpen_Drawings_BaseView.AskVisibleObjects();





            return;


            //NXOpen.Point3d translation1 = new NXOpen.Point3d(-1169.1166216538256, 532.27181521555599, -2507.4890653418188);
            //workPart.ModelingViews.WorkView.SetRotationTranslationScale(rotMatrix1, translation1, 0.044461098658023737);


            ModelingView object0 = modelingViewCollection.FindObject("Isometric") as ModelingView;

            NXOpen.Point3d scaleAboutPoint2 = new NXOpen.Point3d(-586.16731660571156, 1816.1022063622067, 0.0);
            NXOpen.Point3d viewCenter2 = new NXOpen.Point3d(586.1673166057127, -1816.1022063622056, 0.0);
            workPart.ModelingViews.WorkView.ZoomAboutPoint(0.80000000000000004, scaleAboutPoint2, viewCenter2);

            ls.Open();
            ls.WriteLine(workPart.Name);
            ls.Close();

            Face face = null;

            Direction direction = workPart.Directions.CreateDirection(face, Sense.Forward, SmartObject.UpdateOption.DontUpdate);
            //direction.Origin;
            //direction.Vector;

            theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

            //workPart.Drafting.EnterDraftingApplication();

            //theSession.ApplicationSwitchImmediate("UG_APP_MODELING");

            NXOpen.Drawings.DraftingDrawingSheet nullNXOpen_Drawings_DraftingDrawingSheet = null;
            NXOpen.Drawings.DraftingDrawingSheetBuilder draftingDrawingSheetBuilder1;
            draftingDrawingSheetBuilder1 = workPart.DraftingDrawingSheets.CreateDraftingDrawingSheetBuilder(nullNXOpen_Drawings_DraftingDrawingSheet);


            //draftingDrawingSheetBuilder1.MetricSheetTemplateLocation = "F:\\Program Files\\NX1953\\DRAFTING\\templates\\Drawing-A0-Size2D-template.prt";
            //draftingDrawingSheetBuilder1.ScaleNumerator = 1.0;
            //draftingDrawingSheetBuilder1.ScaleDenominator = 1.0; 
            //draftingDrawingSheetBuilder1.Units = NXOpen.Drawings.DrawingSheetBuilder.SheetUnits.Metric;
            //draftingDrawingSheetBuilder1.ProjectionAngle = NXOpen.Drawings.DrawingSheetBuilder.SheetProjectionAngle.Third;
            //NXOpen.NXObject nXObject1;
            //nXObject1 = draftingDrawingSheetBuilder1.Commit();

            //draftingDrawingSheetBuilder1.Destroy();

            //NXOpen.Drawings.BaseView nullNXOpen_Drawings_BaseView = null;
            //NXOpen.Drawings.BaseViewBuilder baseViewBuilder1;
            //baseViewBuilder1 = workPart.DraftingViews.CreateBaseViewBuilder(nullNXOpen_Drawings_BaseView);


            //baseViewBuilder1.Placement.Associative = true;

            NXOpen.ModelingView modelingView1 = ((NXOpen.ModelingView)workPart.ModelingViews.FindObject("Top"));
            baseViewBuilder1.SelectModelView.SelectedView = modelingView1;

            baseViewBuilder1.SecondaryComponents.ObjectType = NXOpen.Drawings.DraftingComponentSelectionBuilder.Geometry.PrimaryGeometry;


            baseViewBuilder1.Style.ViewStyleBase.Part = workPart;

            baseViewBuilder1.Style.ViewStyleBase.PartName = "E:\\Personal\\NX\\NX-Testing\\01_Data\\F5800010214830101.prt";

            bool loadStatus1;
            loadStatus1 = workPart.IsFullyLoaded;

            baseViewBuilder1.SelectModelView.SelectedView = modelingView1;

            

            //uFSession.Part.AskDisplayPart();
            //uFSession.Part.Open();

            //uFSession.Assem.AddPartToAssembly();
            //uFSession.Assem.AskCompPosition();
            //uFSession.Assem.AskHiddenComps();

            //uFSession.Draw.AskBoundaryCurves();
            //uFSession.Draw.AskViewScale();
            //uFSession.Draw.AskViewAngle();

            //uFSession.Disp.SetHighlight();

            //uFSession.Attr.GetUserAttribute();
            //uFSession.Attr.SetUserAttribute();


            //uFSession.Modl.AskBoundingBox();

            //uFSession.Assem.AddPartToAssembly();

            //PdmSession pdmSession = theSession.PdmSession;
            //pdmSession.GetCheckedoutStatusOfAllObjectsInSession()

            //PdmPart pdmPart = pdmSession.getpdmpar

            //string filePath = @"E:\Personal\Job\ORNNOVA_WIPRO\Data\Part1.prt";

            //PartLoadStatus partLoadStatus = null;
            //Part OpenedPart = theSession.Parts.OpenDisplay(filePath, out partLoadStatus);

            //Part newPart = theSession.Parts.NewDisplay(filePath, Part.Units.Millimeters);

            

            //workPart.DrawingSheets.CurrentDrawingSheet.SheetDraftingViews.CreateBaseView();

            //BaseViewBuilder baseViewBuilder = workPart.DraftingViews.CreateBaseViewBuilder();

            

            //PartSaveStatus partSaveStatus = workPart.Save(BasePart.SaveComponents.True, BasePart.CloseAfterSave.False);
            //partSaveStatus.Dispose();

            //ComponentAssembly componentAssembly = workPart.ComponentAssembly;
            //Component component = componentAssembly.RootComponent;
            

            //foreach (Component item in component.GetChildren())
            //{
            //    ls.WriteLine(item.Name);
            //}


            //Feature[] features = workPart.Features.ToArray();

            //foreach (Feature feature in features)
            //{
            //    ls.WriteLine(feature.Name);

            //    foreach (Expression expression in feature.GetExpressions())
            //    {
            //        ls.WriteLine(expression.Name + "," + expression.Value);                    
            //    }
            //}

            //ls.Close();



            //uiSession.NXMessageBox.Show("MyTitle", NXMessageBox.DialogType.Information, "Welcome to NX");





        }

        public static int GetUnloadOption(string dummy) { return (int)NXOpen.Session.LibraryUnloadOption.Immediately; }
    }    
}
