using System;
using Inventor;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;

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
            //Notes:
            //1. Defintion - Document.ComponentDefinition
            //1. API Units In Inventor = cm
            //2. Reference Key, Transient Key, Internal Name
            //3. Nifty Attributes App to see attributes of face or edge
            //4. Sheet.CreateGeometryIntent
            //5. Transient Geometry - Application level
            //6. Transient Objects - Application level
            //7. view.ReferencedDocumentDescriptor.ReferencedDocument
            //8. Progress Bar
            //9. Camera - eye, target and upVector

            //Get Active Inventor Session
            Inventor.Application application = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //MessageBox.Show("Welcome to Inventor", "MyTitle");

            #region ProgressBar

            Inventor.ProgressBar progressBar = application.CreateProgressBar(false, 5, "Test Progress Bar");
            progressBar.UpdateProgress();
            progressBar.Close();

            #endregion

            #region Create Part Doc & Extrude Feature

            PartDocument partDoc = application.Documents.Add(DocumentTypeEnum.kPartDocumentObject) as PartDocument;

            PartComponentDefinition componentDefinition = partDoc.ComponentDefinition;

            WorkPlane workPlane = componentDefinition.WorkPlanes[3];

            PlanarSketch sketch = componentDefinition.Sketches.Add(workPlane);

            sketch.SketchCircles.AddByCenterRadius(application.TransientGeometry.CreatePoint2d(0, 0), 5);

            Profile profile = sketch.Profiles.AddForSolid();

            componentDefinition.Features.ExtrudeFeatures.AddByDistanceExtent(profile, 10, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);

            PropertySet set = partDoc.PropertySets["Inventor User Defined Properties"];
            set.Add("val", "name");

            partDoc.Update();

            #endregion

            #region Create Loft Feature

            // Get active part document
            PartDocument partDoc1 = (PartDocument)application.ActiveDocument;
            PartComponentDefinition compDef = partDoc.ComponentDefinition;
            TransientGeometry tg = application.TransientGeometry;

            // Create sketch 1 on XY Plane (circle dia 10)
            PlanarSketch sketch1 = compDef.Sketches.Add(compDef.WorkPlanes[3]); // XY Plane
            sketch1.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(0, 0), 5); // Radius 5 mm

            // Create offset workplane 20 mm above XY
            WorkPlane offsetPlane = compDef.WorkPlanes.AddByPlaneAndOffset(compDef.WorkPlanes[3], 2.0); // 2 cm = 20 mm
            offsetPlane.Visible = true;

            // Create sketch 2 on offset plane (circle dia 8)
            PlanarSketch sketch2 = compDef.Sketches.Add(offsetPlane);
            sketch2.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(0, 0), 4); // Radius 4 mm

            // Create loft sections
            ObjectCollection loftSections = application.TransientObjects.CreateObjectCollection();
            loftSections.Add(sketch1.Profiles.AddForSolid());
            loftSections.Add(sketch2.Profiles.AddForSolid());

            // Create the loft feature
            LoftDefinition loftDef = compDef.Features.LoftFeatures.CreateLoftDefinition(
                loftSections,
                PartFeatureOperationEnum.kJoinOperation
            );
            compDef.Features.LoftFeatures.Add(loftDef);

            #endregion

            #region Extract Bodies, Faces and Edges

            // Get active part document
            PartDocument partDoc2 = (PartDocument)application.ActiveDocument;
            PartComponentDefinition compDef2 = partDoc.ComponentDefinition;

            // Get all surface and solid bodies
            SurfaceBodies bodies = compDef.SurfaceBodies;

            int bodyCount = 0;
            foreach (SurfaceBody body in bodies)
            {
                bodyCount++;
                Console.WriteLine($"Body {bodyCount}:");

                // List faces
                int faceCount = 0;
                foreach (Face face in body.Faces)
                {
                    faceCount++;

                    Console.WriteLine($"  Face {faceCount}: Type = {face.SurfaceType}");
                }

                // List edges
                int edgeCount = 0;
                foreach (Edge edge in body.Edges)
                {
                    edgeCount++;
                    double[] startPoint = new double[2];
                    double[] endPoint = new double[2];
                    edge.Evaluator.GetEndPoints(ref startPoint, ref endPoint);

                    //Console.WriteLine($"  Edge {edgeCount}: Length = {length:F2} cm");
                }

                Console.WriteLine();
            }

            Console.WriteLine($"Total Bodies: {bodyCount}");

            #endregion

            #region Adjacent Faces

            int bodyIndex = 0;
            foreach (SurfaceBody body in bodies)
            {
                bodyIndex++;
                Console.WriteLine($"Body {bodyIndex}:");

                int faceIndex = 0;
                foreach (Face face in body.Faces)
                {
                    faceIndex++;
                    Console.WriteLine($"  Face {faceIndex}: Type = {face.SurfaceType}");

                    // Use a HashSet to avoid duplicate adjacent faces
                    var adjacentFaces = new System.Collections.Generic.HashSet<Face>();

                    foreach (Edge edge in face.Edges)
                    {
                        Faces edgeFaces = edge.Faces;
                        foreach (Face adjacent in edgeFaces)
                        {
                            if (!adjacent.Equals(face))
                            {
                                adjacentFaces.Add(adjacent);
                            }
                        }
                    }

                    // Output adjacent face info
                    int adjCount = 0;
                    foreach (Face adj in adjacentFaces)
                    {
                        adjCount++;
                        Console.WriteLine($"    Adjacent Face {adjCount}: Type = {adj.SurfaceType}");
                    }
                }
            }

            #endregion

            #region Selection - Need Nifty Attributes to See

            Edge edge1 = application.CommandManager.Pick(SelectionFilterEnum.kPartEdgeFilter, "Select Edge") as Edge;

            AttributeSet set1 = edge1.AttributeSets.Add("MySet");
            set1.Add("Name", ValueTypeEnum.kStringType, "Value");

            #endregion

            #region Read Assembly

            AssemblyDocument asmDoc = application.ActiveDocument as AssemblyDocument;

            ComponentOccurrences compOccs = asmDoc.ComponentDefinition.Occurrences;

            foreach (ComponentOccurrence ComponentOccurrence in compOccs)
            {
                ObjectTypeEnum objectTypeEnum = ComponentOccurrence.Definition.Type;

                //Document doc = ComponentOccurrence.Definition.Document;
                //DocumentTypeEnum documentTypeEnum = doc.DocumentType;

                //ComponentOccurrence.Replace();
            }

            #endregion

            #region Create Drawing Doc & View

            DrawingDocument drawingDocument = application.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject) as DrawingDocument;

            Sheet sheet = drawingDocument.Sheets.Add(DrawingSheetSizeEnum.kA3DrawingSheetSize, PageOrientationTypeEnum.kLandscapePageOrientation);

            DrawingView drawingView = sheet.DrawingViews.AddBaseView(partDoc as _Document, application.TransientGeometry.CreatePoint2d(0, 0), 1, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            List<DrawingCurve> curvestoDimension = new List<DrawingCurve>();

            foreach (DrawingCurve item in drawingView.DrawingCurves)
            {
                Edge modelEdge = item.ModelGeometry as Edge;
            }

            //sheet.DrawingDimensions.GeneralDimensions.AddLinear();

            #endregion

            #region Create Notes

            //sheet.DrawingNotes.LeaderNotes.Add();
            //sheet.DrawingNotes.GeneralNotes.AddFitted();

            #endregion

            #region Change Camera

            Camera camera = application.ActiveView.Camera;
            camera.Eye = application.TransientGeometry.CreatePoint(0, 0, 0);
            camera.Target = application.TransientGeometry.CreatePoint(0, 0, 0);
            camera.UpVector = application.TransientGeometry.CreateUnitVector(1, 0, 0);
            camera.Apply();

            #endregion

        }
    }
}
