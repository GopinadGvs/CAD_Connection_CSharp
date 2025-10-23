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
using NXOpen.Annotations;
using NXOpen.BlockStyler;
using NXOpen.Utilities;
using NXOpen.PDM;
using NXOpen.Die;
//using NXOpen.CAE;

namespace NXAPI
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            //BatchProcessDemo();

            JournalCode();

            //Get Active NX Session from dot net
            //Session session = Session.GetSession();

            ////Get UI
            //UI uI = UI.GetUI();

            ////Get Active NX Session from UFunc
            //UFSession uFSession = UFSession.GetUFSession();

            ////uFSession.Ugmgr.AskNewPartNo()
            ////uFSession.Ugmgr.AskNewPartRev()

            ////PdmSession pdmSession = Session.GetSession().PdmSession;

            ////Get display Part
            //Part displayPart = session.Parts.Display;

            ////Get Work Part
            //Part workPart = session.Parts.Work;

            //ComponentAssembly compAssy = workPart.ComponentAssembly;

            //Component root = compAssy.RootComponent;
            //root.GetChildren();

            //Face face = null;
            //face.OwningPart;

            //Using dot net
            //Point myPoint = workPart.Points.CreatePoint(new Point3d(100, 100, 100));
            //myPoint.SetVisibility(SmartObject.VisibilityOption.Visible);

            ////Using UFunc
            //uFSession.Curve.CreatePoint(new double[] { 200, 200, 200 }, out Tag pointTag);


            //TaggedObject myUfuncPoint = NXObjectManager.Get(pointTag);

            //if(myUfuncPoint is Point)
            //{
            //    myUfuncPoint.GetType();
            //}

            //uFSession.Modl.CreatePointsFeature(new double[] { 200, 200, 200 }, out Tag pointTag);

            //Using Builder
            //Point point = workPart.Points.CreatePoint(new Point3d(0, 0, 10));
            //NXOpen.Features.Feature nullNXOpen_Features_Feature = null;
            //PointFeatureBuilder pointFeatureBuilder = session.Parts.Work.BaseFeatures.CreatePointFeatureBuilder(nullNXOpen_Features_Feature);
            //pointFeatureBuilder.Point = point;
            //NXObject mypointObject = pointFeatureBuilder.Commit();

            //Point mypp = mypointObject as Point;

            //pointFeatureBuilder.Destroy();

            //NXRemotableObject
            //TaggedObject
            //NXObject
            //DisplayableObject

            //uI.SelectionManager.SelectTaggedObjects()

            //ClearanceSetCollection clearanceSets = workPart.ComponentAssembly.ClearanceSets;

        }


        public static void BatchProcessDemo()
        {
            string path = @"E:\Personal\NX_Data\Part_BatchProcess\Part_For_BatchProcess_Demo.prt";

            Session session = Session.GetSession();

            UFSession uFSession = UFSession.GetUFSession();

            uFSession.Part.Open(path, out Tag part, out UFPart.LoadStatus status);

            string exp = "Length = 200";

            uFSession.Modl.EditExp(exp);

            uFSession.Part.Save();
        }

        public static void JournalCode()
        {
            NXOpen.Session theSession = NXOpen.Session.GetSession();
            NXOpen.Part workPart = theSession.Parts.Work;
            NXOpen.Part displayPart = theSession.Parts.Display;
            // ----------------------------------------------
            //   Menu: Tools->Expressions...
            // ----------------------------------------------
            theSession.Preferences.Modeling.UpdatePending = false;

            NXOpen.Session.UndoMarkId markId1;
            markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start");

            theSession.SetUndoMarkName(markId1, "Expressions Dialog");

            NXOpen.Session.UndoMarkId markId2;
            markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Create Expression");

            NXOpen.Unit unit1 = ((NXOpen.Unit)workPart.UnitCollection.FindObject("MilliMeter"));
            NXOpen.Expression expression1;
            expression1 = workPart.Expressions.CreateNumberExpression("thk=20", unit1);

            NXOpen.Session.UndoMarkId markId3;
            markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Check Circular");

            NXOpen.NXObject[] objects1 = new NXOpen.NXObject[1];
            objects1[0] = expression1;
            theSession.UpdateManager.MakeUpToDate(objects1, markId3);

            expression1.EditComment("");

            NXOpen.Session.UndoMarkId markId4;
            markId4 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Update Expression");

            int nErrs1;
            nErrs1 = theSession.UpdateManager.DoUpdate(markId4);

            theSession.DeleteUndoMark(markId4, "Update Expression");

            NXOpen.Session.UndoMarkId markId5;
            markId5 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Expressions");

            theSession.DeleteUndoMark(markId5, null);

            NXOpen.Session.UndoMarkId markId6;
            markId6 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Expressions");

            NXOpen.Session.UndoMarkId markId7;
            markId7 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Make Up to Date");

            NXOpen.Session.UndoMarkId markId8;
            markId8 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "NX update");

            int nErrs2;
            nErrs2 = theSession.UpdateManager.DoUpdate(markId8);

            theSession.DeleteUndoMark(markId8, "NX update");

            theSession.DeleteUndoMark(markId7, null);

            theSession.DeleteUndoMark(markId6, null);

            theSession.SetUndoMarkName(markId1, "Expressions");

            NXOpen.PartSaveStatus partSaveStatus1;
            partSaveStatus1 = workPart.SaveAs("E:\\Personal\\NX_Data\\DllSignTest\\Part_After_dll_RunPart_UpdateExp_ByDll.prt");

            partSaveStatus1.Dispose();
        }


        public static int GetUnloadOption(string dummy) { return (int)NXOpen.Session.LibraryUnloadOption.Immediately; }
    }
}
