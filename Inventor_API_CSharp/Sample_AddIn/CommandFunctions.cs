// TODO: This module exists as a convenient location for the code that does the real
// work when a command is executed.  If you're converting VBA macros into add-in 
// commands you can copy the macros here, convert them to CSharp, 
// and change any references to "ThisApplication" to "Globals.invApp".

using Inventor;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace My_CSharp_AddIn
{
    public static class CommandFunctions
    {
        public static Inventor.Application iApp;

        public static void CreateRectangle()
        {
            iApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            Document iDoc = iApp.ActiveDocument;


            PartDocument iPart = (PartDocument)iApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, "", true);

            // PartDocument  iPart;
            iPart = (PartDocument)iApp.ActiveDocument;

            TransientGeometry oTranGeom = iApp.TransientGeometry;

            PartComponentDefinition iPartComp = (PartComponentDefinition)iPart.ComponentDefinition;

            // Create a new sketch on the X-Y work plane.

            PlanarSketch oSketch = iPartComp.Sketches.Add(iPartComp.WorkPlanes[3]);

            Point2d oPlacePoint = oTranGeom.CreatePoint2d(0, 0);

            // create rect
            Point2d point1 = oTranGeom.CreatePoint2d(0, 0);

            Point2d point2 = oTranGeom.CreatePoint2d(5, 5);

            Object see = oSketch.SketchLines.AddAsTwoPointCenteredRectangle(point1, point2);

            Profile iSolidProf = oSketch.Profiles.AddForSolid();

            ExtrudeDefinition oExtrudeDef = iPartComp.Features.ExtrudeFeatures.CreateExtrudeDefinition(iSolidProf, PartFeatureOperationEnum.kJoinOperation);

            oExtrudeDef.SetDistanceExtent(10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);

            iPartComp.Features.ExtrudeFeatures.Add(oExtrudeDef);

            oSketch = iPartComp.Sketches.Add(iPartComp.WorkPlanes[3]);

            oSketch.SketchCircles.AddByCenterRadius(oPlacePoint, 0.5);

            iSolidProf = oSketch.Profiles.AddForSolid();

            oExtrudeDef = iPartComp.Features.ExtrudeFeatures.CreateExtrudeDefinition(iSolidProf, PartFeatureOperationEnum.kCutOperation);

            oExtrudeDef.SetDistanceExtent(10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);

            iPartComp.Features.ExtrudeFeatures.Add(oExtrudeDef);
        }
    }
}
