using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.IO;
using Newtonsoft.Json;
using System.Data.SQLite;
using System.Data;
using System.Diagnostics;

namespace Inv_Test
{
    internal class Samples
    {
        static void Main(string[] args)
        {
            Application application = Marshal.GetActiveObject("Inventor.Application") as Application;

            PartDocument partDocument = application.Documents.Add(DocumentTypeEnum.kPartDocumentObject) as PartDocument;

            PartComponentDefinition compDef = partDocument.ComponentDefinition;

            PlanarSketch sketch = compDef.Sketches.Add(compDef.WorkPlanes[3]);

            sketch.SketchCircles.AddByCenterRadius(application.TransientGeometry.CreatePoint2d(0, 0), 5.0);

            Profile profile = sketch.Profiles.AddForSolid();

            ExtrudeFeature extrude = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                profile, 10.0, PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kJoinOperation);

            PropertySet set = partDocument.PropertySets["Inventor User Defined Properties"];

            Property property = set[""];
            property.Value = "MyVal";

            set.Add("MyVal", "MyProp");

            partDocument.Update();

            //Document document = application.ActiveDocument as Document;

            //DocumentTypeEnum documentTypeEnum = document.DocumentType;

            //if (documentTypeEnum != DocumentTypeEnum.kPartDocumentObject) return;

            //PartDocument partDocument = document as PartDocument;

            //PartFeatures features = partDocument.ComponentDefinition.Features;

            //foreach (PartFeature feature in features)
            //{
            //    Debug.Print($"Feature Name: {feature.Name}, Feature Extended Name: {feature.ExtendedName}");
            //}
        }

        public void AddDimToSelectedEdgesIn3d(Application invApp)
        {
            Edge edge1 = invApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Select Edge1") as Edge;
            AttributeSet myset1 = edge1.AttributeSets.Add("MySet");
            myset1.Add("EDGE_ID", ValueTypeEnum.kStringType, "EDGE_1");

            Edge edge2 = invApp.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Select Edge2") as Edge;
            AttributeSet myset2 = edge2.AttributeSets.Add("MySet");
            myset2.Add("EDGE_ID", ValueTypeEnum.kStringType, "EDGE_2");

            List<string> attribValues = new List<string> { "EDGE_1", "EDGE_2" };

            string drawingPath = @"E:\CAD-TEAM\Gopi\2024\19032024\21-31155-0550-766-A\Workspaces\Workspace\GM- EV\2021\31155\0550\766-A_Rotor Test\003-Lower tooling\M\21-31155-0550-766-A-003-0003_SHT2.dwg";

            DrawingDocument drawingDocument = invApp.Documents.Open(drawingPath) as DrawingDocument;

            DrawingView drawingView = drawingDocument.ActiveSheet.DrawingViews[1];

            List<DrawingCurve> curvestoDimension = new List<DrawingCurve>();

            foreach (DrawingCurve item in drawingView.DrawingCurves)
            {
                Edge modelEdge = item.ModelGeometry as Edge;

                if (modelEdge == null)
                    continue;

                string attribVal = string.Empty;

                try
                {
                    attribVal = modelEdge.AttributeSets["MySet"]["EDGE_ID"].Value;

                    //AttributeSets sets = modelEdge.AttributeSets;

                    //foreach (AttributeSet attset in sets)
                    //{
                    //    if(attset.Name == "MySet")
                    //    {
                    //        attribVal = attset["EDGE_ID"].Value;
                    //    }
                    //}
                }

                catch
                {

                }

                if (attribValues.Contains(attribVal))
                {
                    curvestoDimension.Add(item);
                }
            }
        }

        public void AddAttributesToFace()
        {
            Application invApp = (Application)Marshal.GetActiveObject("Inventor.Application");
            Face face = invApp.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, "Select Face") as Face;
            AttributeSet myset = face.AttributeSets.Add("MySet");
            myset.Add("FACE_ID", ValueTypeEnum.kStringType, "FACE_1");
        }

        public void ChangeCamera(Application invApp)
        {
            Camera camera = invApp.ActiveView.Camera;
            camera.Eye = invApp.TransientGeometry.CreatePoint(0, 0, 0);
            camera.Target = invApp.TransientGeometry.CreatePoint(0, 0, 0);
            camera.UpVector = invApp.TransientGeometry.CreateUnitVector(1, 0, 0);
            camera.Apply();
        }

        public void RenameAssembly(Application invApp)
        {
            ProgressBar progressBar = invApp.CreateProgressBar(false, 5, "Test Progress Bar");
            progressBar.UpdateProgress();
            progressBar.Close();            
        }
    }
}
