using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using JPMorrow.Measurements;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;

namespace JPMorrow.Pdf.Bluebeam.FireAlarm
{
    public class BluebeamSingleHanger
    {
        public double ThreadedRodLength { get; private set; }
        public string ThreadedRodSize { get; private set; }
        public string ThreadedRodLengthFtIn { get => Measure.LengthFromDbl(ThreadedRodLength); }
        public string BatwingAttachment { get; private set; }

        private string[] Hardware { get; set; }
        public string Anchor { get => Hardware[0]; }
        public string HexNut { get => Hardware[1]; }
        public string Washer { get => Hardware[2]; }
        
        private static string[] batwingNames = new string[] { "Batwing - K-12", "Batwing - K-16" };
        private static string[] threadedRodSizes = new string[] { "1/4\"", "3/8\"", "1/2\"" };

        public BluebeamSingleHanger(double threaded_rod_length, string conduit_size, string inividual_hanger_threaded_rod_size)
        {
            ThreadedRodLength = threaded_rod_length;
            if (!BluebeamPdfUtil.FireAlarmConduitSizes.Any(x => x.Equals(conduit_size)))
                throw new Exception("The provided conduit size is not valid. It should be in the format -> 1 1/2\"");

            var dbl_length = Measure.LengthDbl(conduit_size);
            var cmp_length = Measure.LengthDbl("3/4\"");
            var bw = dbl_length <= cmp_length ? batwingNames[0] : batwingNames[1];
            BatwingAttachment = bw + " - " + conduit_size;

            var rod_dia_test = Measure.LengthDbl(inividual_hanger_threaded_rod_size);
            if (rod_dia_test < 0) throw new Exception("The provided threaded rod size of '" + inividual_hanger_threaded_rod_size + "' is not valid");
            ThreadedRodSize = inividual_hanger_threaded_rod_size;

            Hardware = new string[]
            {
                "Hilti Concrete Anchor - " + inividual_hanger_threaded_rod_size,
                "Washer - " + inividual_hanger_threaded_rod_size,
                "Hex Nut - " + inividual_hanger_threaded_rod_size,
            };
        }

        public static IEnumerable<BluebeamSingleHanger> SingleHangersFromConduitPackage(BluebeamConduitPackage p, string threaded_rod_size)
        {
            var ret = new List<BluebeamSingleHanger>();

            foreach(var c in p.Conduit)
            {
                ret.Add(new BluebeamSingleHanger(c.HangerElevation, c.Size, threaded_rod_size));
            }

            return ret;
        }
    }

    public class BluebeamConduit
    {
        public string Size { get; private set; }
        public double Length { get; private set; }
        public string Type { get; private set; }
        public double HangerElevation { get; private set; }

        public BluebeamConduit(string size, double length, string type, double hanger_elevation)
        {
            Size = size;
            Length = length;
            Type = type;
            HangerElevation = hanger_elevation;
        }
    }

    /// <summary>
    /// takes poly lines representing conduit in bluebeam and extracts the total length of them
    /// </summary>
    public class BluebeamConduitPackage
    {
        private static string[] conduitTypes = new string[] { "EMT", "PVC", "MC" };

        public List<BluebeamConduit> Conduit { get; set; } = new List<BluebeamConduit>();

        public override string ToString()
        {
            var o = "\nConduit Package\n";
            foreach(var s in BluebeamPdfUtil.FireAlarmConduitSizes) o += "emt length - " + s + ": " + GetTotalEmtLengthFeetIn(s) + "\n";
            foreach(var s in BluebeamPdfUtil.FireAlarmConduitSizes) o += "pvc length - " + s + ": " + GetTotalPvcLengthFeetIn(s) + "\n";
            foreach(var s in BluebeamPdfUtil.FireAlarmConduitSizes) o += "mc length - " + s + ": " + GetTotalMcCableLengthFeetIn(s) + "\n";

            foreach(var s in BluebeamPdfUtil.FireAlarmConduitSizes) o += "emt couplings - " + s + ": " + GetTotalEmtCouplings(s) + "\n";
            foreach(var s in BluebeamPdfUtil.FireAlarmConduitSizes) o += "pvc couplings - " + s + ": " + GetTotalPvcCouplings(s) + "\n";
            foreach(var s in BluebeamPdfUtil.FireAlarmConduitSizes) o += "mc couplings - " + s + ": " + GetTotalMcCableCouplings(s) + "\n";
            o += "\n";
            return o;
        }

        /// <summary>
        /// Conduit Length Totals
        /// </summary>
        
        public string GetTotalEmtLengthFeetIn(string conduit_size) 
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("EMT"))
                .Select(x => x.Length).Sum();
            return Measure.LengthFromDbl(length);
        }

        public string GetTotalPvcLengthFeetIn(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("PVC"))
                .Select(x => x.Length).Sum();
            return Measure.LengthFromDbl(length);
        }

        public string GetTotalMcCableLengthFeetIn(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("MC"))
                .Select(x => x.Length).Sum();
            return Measure.LengthFromDbl(length);
        }

        public int GetTotalEmtLengthRounded(string conduit_size) 
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("EMT"))
                .Select(x => x.Length).Sum();
            return (int)(Math.Ceiling(length / 10.0d) * 10);
        }

        public int GetTotalPvcLengthRounded(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("PVC"))
                .Select(x => x.Length).Sum();
            return (int)(Math.Ceiling(length / 10.0d) * 10);
        }

        public int GetTotalMcCableLengthRounded(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("MC"))
                .Select(x => x.Length).Sum();
            return (int)(Math.Ceiling(length / 10.0d) * 10);
        }

        /// <summary>
        /// Coupling totals
        /// </summary>
        
        public int GetTotalEmtCouplings(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("EMT"))
                .Select(x => x.Length).Sum();
            return (int)(Math.Ceiling(length) / 10);
        }

        public int GetTotalPvcCouplings(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("PVC"))
                .Select(x => x.Length).Sum();
            return (int)(Math.Ceiling(length) / 10);
        }

        public int GetTotalMcCableCouplings(string conduit_size)
        {
            var length = Conduit.Where(x => x.Size.Equals(conduit_size) && x.Type.Equals("MC"))
                .Select(x => x.Length).Sum();
            return (int)(Math.Ceiling(length) / 10);
        }

        /// <summary>
        /// Single Hanger Totals
        /// </summary>
        
        

        private BluebeamConduitPackage(IEnumerable<PdfAnnotation> annotations)
        {
            foreach(var a in annotations)
            {
                if(a == null) throw new Exception("The annotation provided is null");
                bool has_label = GetAnnotationLabel(a, out string label);
                if(!has_label) throw new Exception("No valid label");
                StoreConduitSizeAndLength(BluebeamPdfUtil.GetSubject(a), BluebeamPdfUtil.GetRcContents(a), label);
            }
        }

        public static BluebeamConduitPackage PackageFromPolyLines(IEnumerable<PdfAnnotation> annotations)
        {
            var conduit_poly_line_annots = new List<PdfAnnotation>();
            foreach(var a in annotations)
            {
                if(BluebeamPdfUtil.IsPolyLine(a) && HasFireAlarmConduitSubject(a))
                    conduit_poly_line_annots.Add(a);
            }
            return new BluebeamConduitPackage(conduit_poly_line_annots);
        }

        private static bool HasFireAlarmConduitSubject(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Subj", out string subject);
            if(!has_subj_name) return false;
            if(!BluebeamPdfUtil.FireAlarmConduitSubjectTags.Any(x => subject.ToLower().Contains(x))) return false;
            return true;
        }

        private bool GetAnnotationLabel(PdfAnnotation a, out string label)
        {
            label = null;
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Label", out string l);
            if(!has_subj_name) return false;
            label = l;
            return true;
        }

        private void StoreConduitSizeAndLength(string subject, string rc_contents, string label)
        {
            var conduit_size = subject.Split(" - ").First();
            conduit_size = conduit_size.Trim(' ');

            if(!BluebeamPdfUtil.FireAlarmConduitSizes.Any(x => x.Equals(conduit_size)))
                throw new Exception("The provided poly line is not the correct type");

            // parse rc contents to get the length
            var matches = Regex.Matches(rc_contents, "<p>(.*?)</p>");
            var length_match = matches
                .Select(x => x.Value).Last().Replace("<p>", "")
                .Replace("</p>", "").Replace("-", " ");

            var length = Measure.LengthDbl(length_match);

            //@TODO: CHECK THAT MEASUREMENTS ARE WORKING BECAUSE THEY ARE NOT
            
            if(length < 0)
                throw new Exception("Conduit length resolved to a negative value");

            string s = subject.ToLower();
            var tags = BluebeamPdfUtil.FireAlarmConduitSubjectTags;

            string type = null;
            if(s.Contains(tags[0])) type = "EMT";
            else if(s.Contains(tags[1])) type = "PVC";
            else if(s.Contains(tags[2])) type = "MC";

            if(type == null) throw new Exception("There is no conduit type matching the provided subject '" + subject + "'");

            var hanger_label_len = label.Split(" - ").Last().Trim();
            var hanger_support_len = Measure.LengthDbl(hanger_label_len);
            if(hanger_support_len < 0) throw new Exception("Hanger support length is invalid: " + label);

            Conduit.Add(new BluebeamConduit(conduit_size, length, type, hanger_support_len));
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

        public List<BluebeamFireAlarmBox> Boxes { get; set; } = new List<BluebeamFireAlarmBox>();

        private static string[] fireAlarmBoxIdentifierTags = new string[] {
            "4\" - fire alarm box - jm" + fireAlarmBoxMarkupVersion,
            "4 11/16\" - fire alarm box - jm" + fireAlarmBoxMarkupVersion,
        };

        public override string ToString()
        {
            var o = "\n Fire Alarm Box Package\n";
            foreach(var b in Boxes)
            {

            }
            return o;
        }

        private BlubeamFireAlarmBoxPackage() {}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="annotations">the annotations to process</param>
        /// <returns> a package of fire alarm boxes</returns>
        public static BlubeamFireAlarmBoxPackage BoxPackageFromAnnotations(IEnumerable<PdfAnnotation> annotations)
        {
            BlubeamFireAlarmBoxPackage package = new BlubeamFireAlarmBoxPackage();

            foreach (var a in annotations)
            {
                bool is_rect = BluebeamPdfUtil.IsRectangle(a);
                bool has_fabs = HasFireAlarmBoxSubject(a, out string subject);
                bool has_group = BluebeamPdfUtil.HasGroupNesting(a, out var group_codes);

                if (!is_rect || !has_fabs || !has_group) continue;

                List<PdfAnnotation> connector_annots = new List<PdfAnnotation>();
                foreach(var g in group_codes)
                {
                    var found_annots = BluebeamPdfUtil.FindAnnotationsByGroupCode(g, annotations);
                    connector_annots.AddRange(found_annots);
                }


                var box_size = subject.Contains("4 11/16\"") ? "4 11/16\"" : "4\"";

                var configs = Regex.Matches(BluebeamPdfUtil.GetRcContents(a), "<p>(.*?)</p>")
                .Select(x => x.Value.Replace("<p>", "")
                .Replace("</p>", "")).Where(x => !x.ToLower()
                .StartsWith("#")).ToList();

                var idx = configs.FindIndex(x => x.StartsWith("Box Configuration"));
                if(idx == -1) continue;
                var box_config = configs[idx].Split(":").Last().Trim().ToUpper();
                package.AddBox(new BluebeamFireAlarmBox(box_size, box_config, connector_annots));
            }

            return package;
        }

        /// <summary>
        /// Add a fire alarm box to the package
        /// </summary>
        public void AddBox(BluebeamFireAlarmBox box) => Boxes.Add(box);

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
    }
}