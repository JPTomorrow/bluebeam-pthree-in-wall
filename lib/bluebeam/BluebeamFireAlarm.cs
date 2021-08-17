using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using JPMorrow.Measurements;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;

namespace JPMorrow.Pdf.Bluebeam.FireAlarm
{
    /// <summary>
    /// takes poly lines representing conduit in bluebeam and extracts the total length of them
    /// </summary>
    public class BluebeamConduitPackage
    {
        private static string[] conduitSizes = new string[] { "1/2\"", "3/4\"", "1\"", "1 1/4\"", "1 1/2\"", "2\"" };
        private Dictionary<string, List<double>> TotalEmtLength { get; set; } = new Dictionary<string, List<double>>()
        {
            {conduitSizes[0], new List<double>()},
            {conduitSizes[1], new List<double>()},
            {conduitSizes[2], new List<double>()},
            {conduitSizes[3], new List<double>()},
            {conduitSizes[4], new List<double>()},
            {conduitSizes[5], new List<double>()},
        };

        private Dictionary<string, List<double>> TotalPvcLength { get; set; } = new Dictionary<string, List<double>>() 
        {
            {conduitSizes[0], new List<double>()},
            {conduitSizes[1], new List<double>()},
            {conduitSizes[2], new List<double>()},
            {conduitSizes[3], new List<double>()},
            {conduitSizes[4], new List<double>()},
            {conduitSizes[5], new List<double>()},
        };

        private Dictionary<string, List<double>> TotalMcCableLength { get; set; } = new Dictionary<string, List<double>>() 
        {
            {conduitSizes[0], new List<double>()},
            {conduitSizes[1], new List<double>()},
            {conduitSizes[2], new List<double>()},
            {conduitSizes[3], new List<double>()},
            {conduitSizes[4], new List<double>()},
            {conduitSizes[5], new List<double>()},
        };

        public override string ToString()
        {
            var o = "\nConduit Package\n";
            foreach(var s in conduitSizes) o += "emt length - " + s + ": " + GetTotalEmtLengthFeetIn(s) + "\n";
            foreach(var s in conduitSizes) o += "pvc length - " + s + ": " + GetTotalPvcLengthFeetIn(s) + "\n";
            foreach(var s in conduitSizes) o += "mc length - " + s + ": " + GetTotalMcCableLengthFeetIn(s) + "\n";

            foreach(var s in conduitSizes) o += "emt couplings - " + s + ": " + GetTotalEmtCouplings(s) + "\n";
            foreach(var s in conduitSizes) o += "pvc couplings - " + s + ": " + GetTotalPvcCouplings(s) + "\n";
            foreach(var s in conduitSizes) o += "mc couplings - " + s + ": " + GetTotalMcCableCouplings(s) + "\n";
            o += "\n";
            return o;
        }

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

        public int GetTotalEmtLengthRounded(string conduit_size) 
        {
            bool s = TotalEmtLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalEmtLength rounded for size " + conduit_size);
            return (int)(Math.Ceiling(doubles.Sum() / 10.0d) * 10);
        }

        public int GetTotalPvcLengthRounded(string conduit_size)
        {
            bool s = TotalPvcLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalPvcLength rounded for size " + conduit_size);
            return (int)(Math.Ceiling(doubles.Sum() / 10.0d) * 10);
        }

        public int GetTotalMcCableLengthRounded(string conduit_size)
        {
            bool s = TotalMcCableLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalMcCableLength rounded for size " + conduit_size);
            return (int)(Math.Ceiling(doubles.Sum() / 10.0d) * 10);
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
            "emt - conduit run - supported",
            "pvc - conduit run - supported",
            "mc - conduit run - supported"
         }; 

        private BluebeamConduitPackage(IEnumerable<PdfAnnotation> annotations)
        {
            foreach(var a in annotations)
            {
                if(a == null) throw new Exception("The annotation provided is null");
                StoreConduitSizeAndLength(GetSubject(a), GetRcContents(a));
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
            bool has_sub_type = elements.TryGetString("/Subtype", out string sub_type);
            if(!has_sub_type || !sub_type.Equals("/PolyLine")) return false;
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

        private static string GetRcContents(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/RC", out string rc);

            if(!has_subj_name)
                throw new Exception("The provided annotation does not have any /RC data");

            return rc;
        }

        private static string GetSubject(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Subj", out string subject);
            
            if(!has_subj_name)
                 throw new Exception("The provided annotation has no subject");

            return subject;
        }

        private void StoreConduitSizeAndLength(string subject, string rc_contents)
        {
            var conduit_size = subject.Split(" - ").First();
            conduit_size = conduit_size.Trim(' ');

            if(!conduitSizes.Any(x => x.Equals(conduit_size)))
                throw new Exception("The provided poly line is not the corrent type");
                
            // parse rc contents to get the length
            var length_match = Regex.Matches(rc_contents, "<p>(.*?)</p>")
                .Select(x => x.Value).Last().Replace("<p>", "")
                .Replace("</p>", "").Replace("-", " ");
                
            var length = Measure.LengthDbl(length_match);

            //@TODO: CHECK THAT MEASUREMENTS ARE WORKING BECAUSE THEY ARE NOT
            
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
    /// 
    /// Goals:
    /// 1. Take in a set of grouped annotations
    /// </summary>

    public class BluebeamFireAlarmBox
    {
        public string BoxConfig { get; private set; }
        public string BoxSize { get; private set; }

        private static string[] connectorSizes = new string[] { "1/2\"", "3/4\"", "1\"", "1 1/4\"", "1 1/2\"", "2\"" };

        private static string[] connectorIdentifierTags = new string[] {
            "emt - fire alarm connector",
            "pvc - fire alarm connector",
            "mc - fire alarm connector",
         };

        private Dictionary<string, int> EmtConnectors { get; set; } = new Dictionary<string, int>()
        {
            {connectorSizes[0], 0},
            {connectorSizes[1], 0},
            {connectorSizes[2], 0},
            {connectorSizes[3], 0},
            {connectorSizes[4], 0},
            {connectorSizes[5], 0},
        };

        private Dictionary<string, int> PvcConnectors { get; set; } = new Dictionary<string, int>()
        {
            {connectorSizes[0], 0},
            {connectorSizes[1], 0},
            {connectorSizes[2], 0},
            {connectorSizes[3], 0},
            {connectorSizes[4], 0},
            {connectorSizes[5], 0},
        };

        private Dictionary<string, int> McConnectors { get; set; } = new Dictionary<string, int>()
        {
            {connectorSizes[0], 0},
            {connectorSizes[1], 0},
            {connectorSizes[2], 0},
            {connectorSizes[3], 0},
            {connectorSizes[4], 0},
            {connectorSizes[5], 0},
        };

        public int GetEmtConnectorCount(string size) 
        {
            bool s = EmtConnectors.TryGetValue(size, out var count);
            if(!s) throw new Exception("Could not retrieve emt connector count size " + size);
            return count;
        }

        public int GetPvcConnectorCount(string size) 
        {
            bool s = PvcConnectors.TryGetValue(size, out var count);
            if(!s) throw new Exception("Could not retrieve pvc connector count size " + size);
            return count;
        }
        
        public int GetMcCableConnectorCount(string size) 
        {
            bool s = McConnectors.TryGetValue(size, out var count);
            if(!s) throw new Exception("Could not retrieve mc connector count size " + size);
            return count;
        }

        public BluebeamFireAlarmBox(
            string box_size, string box_config_char, 
            IEnumerable<PdfAnnotation> connector_annotations)
        {
            BoxConfig = box_config_char;
            BoxSize = box_size;

            foreach(var a in connector_annotations)
            {
                bool s = ProcessSubject(a, out var result);
                if(!s) continue;
                var size = result[0];
                var type = result[1];

                if(type.Equals("emt"))
                    EmtConnectors[size] += 1;
                else if(type.Equals("pvc"))
                    PvcConnectors[size] += 1;
                else if(type.Equals("mc"))
                    McConnectors[size] += 1;
            }
        }

        private static bool ProcessSubject(PdfAnnotation a, out string[] result)
        {
            result = new string[2] { "", "" };
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subj", out string subject);
            if(!s || !connectorIdentifierTags.Any(x => subject.ToLower().Contains(x))) return false;
            result = new string[2] { subject.Split(" - ").First().Trim(), subject.ToLower().Split(" - ")[1].Trim() };
            return true;
        }
    }

    public class BlubeamFireAlarmBoxPackage
    {
        /// <summary>
        /// Incriment this to match fire alarm markup subject line version
        /// </summary>
        private static string fireAlarmBoxMarkupVersion = "1.0.0";

        public List<BluebeamFireAlarmBox> boxes { get; set; } = new List<BluebeamFireAlarmBox>();

        private static string[] fireAlarmBoxIdentifierTags = new string[] {
            "4\" fire alarm box - jm" + fireAlarmBoxMarkupVersion,
            "4 11/16\" fire alarm box - jm" + fireAlarmBoxMarkupVersion,
        };

        public override string ToString()
        {
            var o = "\n Fire Alarm Box Package\n";
            foreach(var b in boxes)
            {

            }
            return o;
        }

        private BlubeamFireAlarmBoxPackage()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="annotations">the annotations to process</param>
        /// <returns> a package of fire alarm boxes</returns>
        public static BlubeamFireAlarmBoxPackage BoxPackageFromAnnotations(IEnumerable<PdfAnnotation> annotations)
        {
            var filtered_annots = new List<PdfAnnotation>();
            BlubeamFireAlarmBoxPackage package = new BlubeamFireAlarmBoxPackage();

            foreach (var a in annotations)
            {
                bool is_rect = IsRectangle(a);
                bool has_fabs = HasFireAlarmBoxSubject(a, out string subject);
                bool has_group = HasGroupNesting(a, out var group_codes);

                if (!is_rect || !has_fabs || !has_group) continue;

                List<PdfAnnotation> connector_annots = new List<PdfAnnotation>();
                foreach(var g in group_codes)
                {
                    var found_annots = FindAnnotationByGroupCode(g, annotations);
                    connector_annots.AddRange(found_annots);
                }

                var box_size = subject.Split(" - ").First().Trim();
                var configs = GetRcContents(a).ToList();

                var idx = configs.FindIndex(x => x.StartsWith("Box Configuration"));
                if(idx == -1) continue;
                var box_config = configs[idx].Split(":").Last().Trim().ToUpper();
                package.AddBox(new BluebeamFireAlarmBox(box_size, box_config, connector_annots));
            }

            return null;
        }

        /// <summary>
        /// Add a fire alarm box to the package
        /// </summary>
        public void AddBox(BluebeamFireAlarmBox box)
        {
            boxes.Add(box);
        }

        private static IEnumerable<string> GetRcContents(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/RC", out string rc);

            if(!has_subj_name)
                throw new Exception("The provided annotation does not have any /RC data");

            var contents = Regex.Matches(rc, "<p>(.*?)</p>")
                .Select(x => x.Value.Replace("<p>", "")
                .Replace("</p>", "")).Where(x => !x.ToLower().StartsWith("#")).ToList();
            return contents;
        }

        /// <summary>
        /// Tell whether the provided pdf annotation is of subtype /Rectangle
        /// </summary>
        private static bool IsRectangle(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subtype", out string sub_type);
            return s && sub_type.Equals("/Square");
        }

        /// <summary>
        /// Does the provided Pdf annotation have the correct subject for a fire alarm box
        /// </summary>
        private static bool HasFireAlarmBoxSubject(PdfAnnotation a, out string ret_subject)
        {
            ret_subject = string.Empty;
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subj", out string subject);
            if (s) ret_subject = subject;
            return s && fireAlarmBoxIdentifierTags.Any(x => x.Equals(subject.Trim().ToLower()));
        }

        private static bool HasGroupNesting(PdfAnnotation a, out List<string> ret_group_codes)
        {
            ret_group_codes = new List<string>();
            var elements = a.Elements;
            bool s = elements.TryGetValue("/GroupNesting", out PdfItem item);

            if (!s) return false;

            var groups = item.ToString()
                .Split(" ").Where(x => x.StartsWith('/'))
                .Select(x => x.Substring(1)).ToList();

            ret_group_codes = groups;
            return groups.Any();
        }

        private static IEnumerable<PdfAnnotation> FindAnnotationByGroupCode(string group_code, IEnumerable<PdfAnnotation> search)
        {
            List<PdfAnnotation> ret = new List<PdfAnnotation>();

            foreach(var a in search)
            {
                var elements = a.Elements;
                bool s = elements.TryGetString("/NM", out string cmp_code);
                if(!s || string.IsNullOrWhiteSpace(cmp_code)) continue;
                if(cmp_code.Equals(group_code)) ret.Add(a);
            }

            return ret;
        }
    }
}