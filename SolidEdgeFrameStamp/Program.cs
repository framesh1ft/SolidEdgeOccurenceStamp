using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFramework;
using SolidEdgeAssembly;
using System.Runtime.InteropServices;
using System.Web;

namespace SolidEdgeFrameStamp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var app = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

            var doc = app.ActiveDocument as AssemblyDocument;
            if (doc == null) 
            { 
                Console.WriteLine("Active document is not an assembly."); 
                return; 
            }

            var frames = doc.StructuralFrames;
            if(frames == null || frames.Count == 0)
            {
                Console.WriteLine("No frames found.");
                return;
            }

            foreach (StructuralFrame frame in frames)
            {
                int sectionCount = 0;
                Array sections = Array.CreateInstance(typeof(object), 0);
                frame.GetSections(out sectionCount, ref sections);

                for (int i = 0; i < sectionCount; i++)
                {
                    int secID = (int)sections.GetValue(i);

                    // Cut Length - NOTE: Solid Edge internal units are meters.
                    double cutLength;
                    frame.SegmentCutLength(secID, out cutLength, out _);
                    double lengthMm = Math.Round(cutLength * 1000.0, 0); // Determine rounding policy.

                    // Miter angles — radians internally; round per your policy
                    // (see EndFaceEndAngle / SideFaceEndAngle on StructuralFrame —
                    //  exact signatures are in the SDK help; they take the section id)

                    // Get the occurrence behind this section
                    object occID; object occObj;
                    frame.ReturnOccurrenceForGivenSectionID(secID, out occID, out occObj);
                    var occ = (Occurrence)occObj;

                    // Profile identity: read the stock part number you stamped
                    // into the .par file properties (occ.OccurrenceDocument → SummaryInfo/Custom props)


                    // Build the deterministic key
                    double endAngle1, endAngle2, sideAngle1, sideAngle2;
                    occ.EndFaceEndAngle(out endAngle1, out endAngle2);
                    occ.SideFaceEndAngle(out sideAngle1, out sideAngle2);
                    string key = $"{secID}-L{lengthMm:0}-M{endAngle1:0}{endAngle2:0}{sideAngle1:0}{sideAngle2:0}";

                    // Stamp it on the occurrence so CADLink can read it
                    // → custom occurrence properties on the Occurrence object.
                    //   Verify the exact member in the SDK help ("occurrence properties",
                    //   available ST4+); overwrite if present so re-runs are idempotent.
                    StampOccurrence(occ, "ErpPartKey", key);
                    StampOccurrence(occ, "CutLengthMm", lengthMm.ToString("0"));
                }

            }
            doc.Save();
            Console.WriteLine("Frame members stamped.");
        }

        public static void StampOccurrence(Occurrence occ, string occ_property, string value)
        {
            occ.CustomPropertyValue[occ_property] = value;
        }
    }
}
