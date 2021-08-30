

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JPMorrow.PDF;
using MoreLinq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

namespace JPMorrow.Pdf.Bluebeam.P3
{
    /// <summary>
    /// A class to handle the conversion from a long device code to a short hand device code 
    /// </summary>
    public class BluebeamP3ShorhandDeviceCodeResolver
    {
        private class ShorthandDeviceCodePair
        {
            public int shorhandCodeFirst { get; private set; }
            public int shorhandCodeSecond { get; private set; }
            public int ShorthandCodeNumber
            {
                get
                {
                    string str_cvt = shorhandCodeFirst.ToString() + shorhandCodeSecond.ToString();
                    bool s = int.TryParse(str_cvt, out int re_cvt);
                    if (!s) throw new Exception("Shorthand code integer failed to parse: " + str_cvt);
                    return re_cvt;
                }
            }

            public string ShorthandCodeString { get => ((char)shorhandCodeFirst).ToString() + shorhandCodeSecond.ToString(); }
            public string LongDeviceCode { get; private set; }

            public ShorthandDeviceCodePair(int short_first_part, int short_second_part, string long_code)
            {
                shorhandCodeFirst = short_first_part;
                shorhandCodeSecond = short_second_part;
                LongDeviceCode = long_code;
            }

            /// <summary>
            /// Use the current shorhand device code in this 
            /// structure to generate an increment of that code
            /// </summary>
            /// <param name="prev_shorhand_int"></param>
            public static ShorthandDeviceCodePair GenerateIncrementedShorhandDeviceCode(
                ShorthandDeviceCodePair previous, string long_device_code)
            {
                if (previous == null) return new ShorthandDeviceCodePair(97, 1, long_device_code);

                var first_part = previous.shorhandCodeFirst;
                var second_part = previous.shorhandCodeSecond;

                if (second_part == 5)
                {
                    if (first_part == 122) first_part = 97;
                    else first_part++;

                    second_part = 1;
                }
                else second_part++;

                return new ShorthandDeviceCodePair(first_part, second_part, long_device_code);
            }
        }

        private List<ShorthandDeviceCodePair> Pairs { get; set; } = new List<ShorthandDeviceCodePair>();

        public BluebeamP3ShorhandDeviceCodeResolver() { }

        /// <summary>
        /// Generate and store a shorhand device code that can be retrieved 
        /// from this structure by passing it a long hand device code
        /// </summary>
        /// <param name="long_device">long device code names</param>
        /// <returns>the shorhand device code, string.empty if no device code matches</returns>
        public string StoreShorthandCode(string long_device_code)
        {
            if (Pairs.Any(x => x.LongDeviceCode.Equals(long_device_code)))
                return Pairs.Where(x => x.LongDeviceCode.Equals(long_device_code)).First().ShorthandCodeString;

            var last_code = GetMostRecentShorthandDeviceCode();
            var new_code = ShorthandDeviceCodePair.GenerateIncrementedShorhandDeviceCode(last_code, long_device_code);

            Pairs.Add(new_code);
            return new_code.ShorthandCodeString;
        }

        public string RetrieveShorthandCode(string long_device_code)
        {
            var idx = Pairs.FindIndex(x => x.LongDeviceCode.Equals(long_device_code));
            if (idx == -1) return string.Empty;
            return Pairs[idx].ShorthandCodeString;
        }

        /// <summary>
        /// Get the most recently generated shorhand device code structure
        /// </summary>
        /// <returns>DeviceCode Structure if there is one, otherwise null</returns>
        private ShorthandDeviceCodePair GetMostRecentShorthandDeviceCode()
        {
            if (!Pairs.Any()) return null;
            return Pairs.MaxBy(x => x.ShorthandCodeNumber).First();
        }

        public override string ToString()
        {
            return string.Join("\n", Pairs.Select(x => string.Format("{0} : {1}", x.ShorthandCodeString, x.LongDeviceCode)));
        }
    }

    public class BluebeamP3BoxConfig
    {
        public string Subject { get; set; }
        public string BoxType { get; set; }
        public string BoxSize { get; set; }
        public string Gang { get; set; }
        public string PlasterRingDepth { get; set; }
        public string EntryConnectorType { get; set; }
        public string EntryConnectorSize { get; set; }
        public int TopEmtConnectors { get; set; }
        public int BottomEmtConnectors { get; set; }
        public int TopPvcConnectors { get; set; }
        public int BottomPvcConnectors { get; set; }
        public int TopMcConnectors { get; set; }
        public int BottomMcConnectors { get; set; }
        public string BundleName { get; set; }
        public string BoxElevation { get; set; }
        public PdfAnnotation Annotation { get; set; }
    }

    public class BluebeamP3Box
    {
        public BluebeamP3BoxConfig Config { get; private set; }
        public string DeviceCode { get; private set; } = "";

        public BluebeamP3Box(BluebeamP3BoxConfig config)
        {
            Config = config;
            ProcessFields();
        }

        private void ProcessFields()
        {
            if (Config.BoxSize.Equals("small"))
                DeviceCode += "S";
            else if (Config.BoxSize.Equals("extended"))
                DeviceCode += "X";

            if (Config.Subject.Contains("Fire Alarm"))
                DeviceCode += "F";
            else if (Config.BoxType.Equals("powered"))
                DeviceCode += "P";
            else if (Config.BoxType.Equals("non-powered"))
                DeviceCode += "N";

            // gang
            if (Config.Gang == "1") DeviceCode += "1";
            else if (Config.Gang.Equals("2")) DeviceCode += "2";
            else if (Config.Gang.Equals("3")) DeviceCode += "3";
            else if (Config.Gang.Equals("4")) DeviceCode += "4";
            else if (Config.Gang.Equals("round")) DeviceCode += "R";
            else if (Config.Gang.Equals("E")) DeviceCode += "E";

            var connector_size_swap = new Dictionary<string, string>()
            {
                { "1/2",    "2"},
                { "3/4",    "3"},
                { "1",      "4"},
                { "1 1/4",  "5"},
                { "1 1/2",  "6"},
                { "2",      "8"}
            };

            // connector size
            if (Config.EntryConnectorType.Equals("MC")) DeviceCode += "M";
            else if (Config.EntryConnectorType.Equals("EMT") || Config.EntryConnectorType.Equals("PVC"))
            {
                var cc = Config.EntryConnectorSize.Remove(Config.EntryConnectorSize.Length - 1);

                bool s = connector_size_swap.TryGetValue(cc, out var add_cc);
                if (!s) throw new System.Exception("Connector Size could not be resolved");
                DeviceCode += add_cc + "C";
            }

            // plaster ring
            if (Config.PlasterRingDepth.Equals("adjustable")) DeviceCode += "-A";
            else
            {
                var pr = Config.PlasterRingDepth.Remove(Config.PlasterRingDepth.Length - 1);
                DeviceCode += "-" + pr;
            }

            for (var i = 0; i < Config.TopEmtConnectors; i++)
                DeviceCode += "-CT";
            for (var i = 0; i < Config.BottomEmtConnectors; i++)
                DeviceCode += "-CB";

            for (var i = 0; i < Config.TopPvcConnectors; i++)
                DeviceCode += "-PT";
            for (var i = 0; i < Config.BottomPvcConnectors; i++)
                DeviceCode += "-PB";

            for (var i = 0; i < Config.TopMcConnectors; i++)
                DeviceCode += "-MT";
            for (var i = 0; i < Config.BottomMcConnectors; i++)
                DeviceCode += "-MB";
        }
    }

    public class BluebeamP3BoxCollection
    {
        public static string[] ExpectedBluebeamColumns { get; } = new string[] {
            "Box Type",
            "Box Size",
            "Gang",
            "Plaster Ring Depth",
            "Entry Connector Type",
            "Connector Size",
            "Top Connectors - EMT",
            "Bottom Connectors - EMT",
            "Top Connectors - PVC",
            "Bottom Connectors - PVC",
            "Top Connectors - MC",
            "Bottom Connectors - MC",
            "Bundle Name",
            "Box Elevation"
        };

        public BluebeamP3ShorhandDeviceCodeResolver BSHD_Resolver { get; private set; } = new BluebeamP3ShorhandDeviceCodeResolver();
        public List<BluebeamP3Box> Boxes { get; private set; } = new List<BluebeamP3Box>();
        public BluebeamP3BoxCollection()
        {

        }

        public static BluebeamP3BoxCollection BoxPackageFromAnnotations(
            IEnumerable<PdfAnnotation> box_annotations, PdfCustomColumnCollection columns)
        {
            BluebeamP3BoxCollection collection = new BluebeamP3BoxCollection();
            foreach (var annot in box_annotations)
            {
                var a = annot;
                bool has_box_subject = HasP3BoxSubject(a, out string subject);
                bool is_rect = BluebeamPdfUtil.IsRectangle(a);
                if (!has_box_subject || !is_rect) continue;

                var column_data = columns.MatchColumnHeaderToBISData(a);
                BluebeamP3BoxConfig config = new BluebeamP3BoxConfig();
                config.BoxType = column_data["Box Type"];
                config.BoxSize = column_data["Box Size"];
                config.Gang = column_data["Gang"];
                config.PlasterRingDepth = column_data["Plaster Ring Depth"];
                config.EntryConnectorType = column_data["Entry Connector Type"];
                config.EntryConnectorSize = column_data["Connector Size"];
                config.TopEmtConnectors = int.Parse(column_data["Top Connectors - EMT"]);
                config.BottomEmtConnectors = int.Parse(column_data["Bottom Connectors - EMT"]);
                config.TopPvcConnectors = int.Parse(column_data["Top Connectors - PVC"]);
                config.BottomPvcConnectors = int.Parse(column_data["Bottom Connectors - PVC"]);
                config.TopMcConnectors = int.Parse(column_data["Top Connectors - MC"]);
                config.BottomMcConnectors = int.Parse(column_data["Bottom Connectors - MC"]);
                config.BundleName = column_data["Bundle Name"];
                config.BoxElevation = column_data["Box Elevation"];
                config.Subject = subject;
                config.Annotation = annot;

                var box = new BluebeamP3Box(config);
                collection.Boxes.Add(box);
                collection.BSHD_Resolver.StoreShorthandCode(box.DeviceCode);
            }

            return collection;
        }

        private static string[] P3BoxSubjectNames = new string[] {
            "Power Box",
            "Data Box",
            "Fire Alarm Box",
            "Lighting Box"
         };

        private static bool HasP3BoxSubject(PdfAnnotation a, out string subject)
        {
            subject = "";
            var els = a.Elements;
            bool s = els.TryGetString("/Subj", out string subj);
            bool has_subj = P3BoxSubjectNames.Any(x => subj.Contains(x));
            if (has_subj) subject = subj;
            return s && has_subj;
        }

        public void SaveMarkupPdf(string input_pdf_filepath, string pdf_output_path, Pdforge f)
        {
            PdfDocument copy = PdfReader.Open(input_pdf_filepath, PdfDocumentOpenMode.Import);

            for (int Pg = 0; Pg < f.InputDocument.Pages.Count; Pg++)
            {
                var page = copy.Pages[Pg];

                foreach (PdfAnnotation a in page.Annotations)
                {
                    //if (Boxes.Any(x => x.Config.Annotation.Equals(aa)))
                    a.Elements.SetString("/Subj", "POOPOOCACA");

                }
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var enc1252 = Encoding.GetEncoding(1252);

            copy.Save(pdf_output_path);
        }
    }
}