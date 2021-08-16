using System;
using System.Collections.Generic;
using System.Linq;
using JPMorrow.Measurements;
using PdfSharp.Pdf.Annotations;

namespace JPMorrow.Pdf.Bluebeam.FireAlarm
{
    /// <summary>
    /// takes poly lines representing conduit in bluebeam and extracts the total length of them
    /// </summary>
    public class BluebeamConduitPackage
    {
        private static string[] conduitSizes = new string[] { "1/2\"", "3/4\"", "1\"", "1 1/2\""};
        private Dictionary<string, List<double>> TotalEmtLength { get; set; } = new Dictionary<string, List<double>>()
        {
            {conduitSizes[0], new List<double>()},
            {conduitSizes[1], new List<double>()},
            {conduitSizes[2], new List<double>()},
            {conduitSizes[3], new List<double>()},
        };

        private Dictionary<string, List<double>> TotalPvcLength { get; set; } = new Dictionary<string, List<double>>() 
        {
            {conduitSizes[0], new List<double>()},
            {conduitSizes[1], new List<double>()},
            {conduitSizes[2], new List<double>()},
            {conduitSizes[3], new List<double>()},
        };

        private Dictionary<string, List<double>> TotalMcCableLength { get; set; } = new Dictionary<string, List<double>>() 
        {
            {conduitSizes[0], new List<double>()},
            {conduitSizes[1], new List<double>()},
            {conduitSizes[2], new List<double>()},
            {conduitSizes[3], new List<double>()},
        };

        /// <summary>
        /// Conduit Length Totals
        /// </summary>
        
        public string GetTotalEmtLengthFeetIn(string conduit_size) 
        {
            bool s = TotalEmtLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalEmtLength in feet and inches for size " + conduit_size);
            return Measure.LengthFromDbl(doubles.Sum());
        }

        public string GetTotalPvcLengthFeetIn(string conduit_size)
        {
            bool s = TotalPvcLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalPvcLength in feet and inches for size " + conduit_size);
            return Measure.LengthFromDbl(doubles.Sum());
        }

        public string GetTotalMcCableLengthFeetIn(string conduit_size)
        {
            bool s = TotalMcCableLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalMcCableLength in feet and inches for size " + conduit_size);
            return Measure.LengthFromDbl(doubles.Sum());
        }

        /// <summary>
        /// Coupling totals
        /// </summary>
        
        public int GetTotalEmtCouplings(string conduit_size)
        {
            bool s = TotalMcCableLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not get total connectors for emt of size " + conduit_size);
            var coupling_count = Math.Ceiling(doubles.Sum()) / 10;
            return (int)coupling_count;
        }

        public int GetTotalPvcCouplings(string conduit_size)
        {
            bool s = TotalPvcLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not get total connectors for pvc of size " + conduit_size);
            var coupling_count = Math.Ceiling(doubles.Sum()) / 10;
            return (int)coupling_count;
        }

        public int GetTotalMcCableCouplings(string conduit_size)
        {
            bool s = TotalMcCableLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not get total connectors for mc cable of size " + conduit_size);
            var coupling_count = Math.Ceiling(doubles.Sum()) / 10;
            return (int)coupling_count;
        }

        private static string[] fireAlarmSubjectTags = new string[] {
            "conduit - emt - fire alarm",
            "conduit - pvc - fire alarm",
            "conduit - mc - fire alarm"
         }; 

        private BluebeamConduitPackage(IEnumerable<PdfAnnotation> annotations)
        {
            foreach(var a in annotations)
            {
                if(a == null) throw new Exception("The annotation provided is null");
                StoreConduitSizeAndLength(GetSubject(a), GetContents(a));
            }
        }

        public static BluebeamConduitPackage PackageFromPolyLines(IEnumerable<PdfAnnotation> annotations)
        {
            var conduit_poly_line_annots = new List<PdfAnnotation>();
            foreach(var a in annotations)
            {
                if(IsPolyLine(a) && HasFireAlarmSubject(a))
                    conduit_poly_line_annots.Add(a);
            }

            return new BluebeamConduitPackage(conduit_poly_line_annots);
        }

        private static bool IsPolyLine(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_sub_type = elements.TryGetString("/SubType", out string sub_type);
            if(!has_sub_type || !sub_type.ToLower().Equals("/PolyLine")) return false;
            return true;
        }

        private static bool HasFireAlarmSubject(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Subj", out string subject);
            if(!has_subj_name) return false;
            if(!fireAlarmSubjectTags.Any(x => subject.ToLower().Contains(x))) return false;
            return true;
        }

        private static string GetContents(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Contents", out string contents);

            if(!has_subj_name)
                throw new Exception("The provided annotation does not have any contents");

            return contents;
        }

        private static string GetSubject(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Subj", out string subject);
            
            if(!has_subj_name)
                 throw new Exception("The provided annotation has no subject");

            return subject;
        }

        private void StoreConduitSizeAndLength(string subject, string contents)
        {
            var conduit_size = subject.Split(" - ").First();
            conduit_size = conduit_size.Trim(' ');

            if(!conduitSizes.Any(x => x.Equals(conduit_size)))
                throw new Exception("The provided poly line is not the corrent type");

            var length = Measure.LengthDbl(contents);

            if(length < 0)
                throw new Exception("Conduit length resolved to a negative value");

            string s = subject.ToLower();

            if(s.Contains(fireAlarmSubjectTags[0])) TotalEmtLength[conduit_size].Add(length);
            else if(s.Contains(fireAlarmSubjectTags[1])) TotalPvcLength[conduit_size].Add(length);
            else if(s.Contains(fireAlarmSubjectTags[2])) TotalMcCableLength[conduit_size].Add(length);
        }
    }

    /// <summary>
    /// THIS CLASS WILL TAKE IN A SET OF RECTANGLE ANNOTATIONS AND 
    /// RETURN A TOTAL COUNT OF ALL THE DIFFERENT SIZED FIRE ALARM BOXES
    /// AND THEIR KNOCKOUTS.
    /// </summary>



    public class BlubeamFireAlarmBoxPackage
    {
        /// <summary>
        /// Incriment this to match fire alarm markup subject line version
        /// </summary>
        private static string fireAlarmBoxMarkupVersion = "1.0.0";

        private string[] fireAlarmBoxSizes = new string[] {

        };

        private static string[] fireAlarmBoxIdentifierTags = new string[] {
            "4\" fire alarm box - jm" + fireAlarmBoxMarkupVersion,
            "4 11/16\" fire alarm box - jm" + fireAlarmBoxMarkupVersion,
        };

        private static Dictionary<string, int> boxPropertys = new Dictionary<string, int> {
            {"1/2\" KOs",       0},
            {"3/4\" KOs",       0},
            {"1\" KOs",         0},
            {"1 1/2\" KOs",     0}
        };



        /* private BlubeamFireAlarmBoxPackage()
        {

        } */

        public static BlubeamFireAlarmBoxPackage BoxPackageFromRectangleAnnotations(IEnumerable<PdfAnnotation> annotations)
        {
            var filtered_annots = new List<PdfAnnotation>();

            foreach (var a in annotations)
            {
                bool passed = IsRectangle(a) || HasFireAlarmBoxSubject(a, out string subject);
                if (!passed) continue;

            }

            return null;
        }

        private static bool IsRectangle(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subtype", out string sub_type);
            return s && sub_type.Equals("/Rectangle");
        }

        private static bool HasFireAlarmBoxSubject(PdfAnnotation a, out string ret_subject)
        {
            ret_subject = string.Empty;
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subj", out string subject);
            if (s) ret_subject = subject;
            return s && fireAlarmBoxIdentifierTags.Any(x => x.Equals(subject.Trim().ToLower()));
        }
    }
}