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
            if (!s) throw new Exception("Could not retrieve emt connector count size " + size);
            return count;
        }

        public int GetPvcConnectorCount(string size)
        {
            bool s = PvcConnectors.TryGetValue(size, out var count);
            if (!s) throw new Exception("Could not retrieve pvc connector count size " + size);
            return count;
        }

        public int GetMcCableConnectorCount(string size)
        {
            bool s = McConnectors.TryGetValue(size, out var count);
            if (!s) throw new Exception("Could not retrieve mc connector count size " + size);
            return count;
        }

        public BluebeamFireAlarmBox(
            string box_size, string box_config_char,
            IEnumerable<PdfAnnotation> connector_annotations)
        {
            BoxConfig = box_config_char;
            BoxSize = box_size;

            foreach (var a in connector_annotations)
            {
                bool s = ProcessSubject(a, out var result);
                if (!s) continue;
                var size = result[0];
                var type = result[1];

                if (type.Equals("emt"))
                    EmtConnectors[size] += 1;
                else if (type.Equals("pvc"))
                    PvcConnectors[size] += 1;
                else if (type.Equals("mc"))
                    McConnectors[size] += 1;
            }
        }

        private static bool ProcessSubject(PdfAnnotation a, out string[] result)
        {
            result = new string[2] { "", "" };
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subj", out string subject);
            if (!s || !connectorIdentifierTags.Any(x => subject.ToLower().Contains(x))) return false;
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
            foreach (var b in Boxes)
            {

            }
            return o;
        }

        private BlubeamFireAlarmBoxPackage() { }

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
                foreach (var g in group_codes)
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
                if (idx == -1) continue;
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