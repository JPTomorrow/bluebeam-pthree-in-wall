

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JPMorrow.PDF;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf.Advanced;
using MB = MoreLinq.Extensions.MaxByExtension;

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
            var ret = MB.MaxBy(Pairs, x => x.ShorthandCodeNumber).First();
            return ret;
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

                // Console.WriteLine("test: " + column_data["Top Connectors - EMT"]);
                config.TopEmtConnectors = int.Parse(column_data["Top Connectors - EMT"]);
                config.BottomEmtConnectors = int.Parse(column_data["Bottom Connectors - EMT"]);
                config.TopPvcConnectors = int.Parse(column_data["Top Connectors - PVC"]);
                config.BottomPvcConnectors = int.Parse(column_data["Bottom Connectors - PVC"]);
                config.TopMcConnectors = int.Parse(column_data["Top Connectors - MC"]);
                config.BottomMcConnectors = int.Parse(column_data["Bottom Connectors - MC"]);
                Console.WriteLine("end");

                config.BundleName = column_data["Bundle Name"];
                config.BoxElevation = "";
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

        public void SaveMarkupPdf(
            string input_pdf_filepath, string pdf_output_path,
            Pdforge f, PdfCustomColumnCollection columns)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var enc1252 = Encoding.GetEncoding(1252);

            PdfDocument doc = PdfReader.Open(input_pdf_filepath, PdfDocumentOpenMode.Modify);
            /*             copy = PdfReader.Open(input_pdf_filepath, PdfDocumentOpenMode.Import);
                        copy.Save(pdf_output_path);
                        PdfDocument doc = PdfReader.Open(pdf_output_path, PdfDocumentOpenMode.Modify); */

            for (int Pg = 0; Pg < doc.Pages.Count; Pg++)
            {
                var page = doc.Pages[Pg];

                /* // ungoup all the annotations in the document in order to process them better
                foreach (PdfAnnotation a in page.Annotations)
                {
                    bool s = a.Elements.TryGetValue("/GroupNesting", out var nestings);
                    bool ss = a.Elements.TryGetString("/RT", out var rt);
                    bool sss = a.Elements.TryGetValue("/IRT", out var irt);

                    if (s)
                    {
                        // check child annotations for p3 box


                        foreach (PdfAnnotation a in children)
                        {
                            bool has_box_subject = HasP3BoxSubject(a, out string subject);
                            bool is_rect = BluebeamPdfUtil.IsRectangle(a);
                        }

                        a.Elements.Remove("/GroupNesting");
                    }
                    if (ss && rt.Equals("/Group")) a.Elements.Remove("/RT");
                    if (sss) a.Elements.Remove("/IRT");
                } */

                // make annotation modifications here
                foreach (PdfAnnotation a in page.Annotations)
                {
                    foreach (var b in Boxes)
                    {
                        var annot = b.Config.Annotation;
                        bool has_id1 = annot.Elements.TryGetString("/NM", out string id);
                        bool has_id2 = a.Elements.TryGetString("/NM", out string id2);

                        if (has_id1 && has_id2 && id.Equals(id2))
                        {
                            var short_device_code = BSHD_Resolver.RetrieveShorthandCode(b.DeviceCode);
                            var new_data1 = columns.ModifyBSIData(doc, a, "Short Device Code", short_device_code);
                            a.Elements.SetValue("/BSIColumnData", new_data1);
                            var new_data2 = columns.ModifyBSIData(doc, a, "Long Device Code", b.DeviceCode);
                            a.Elements.SetValue("/BSIColumnData", new_data2);
                            PlaceTextAnnotationTagForAnnotation(page, a, short_device_code);
                        }
                    }
                }
            }

            doc.Save(pdf_output_path);
        }

        private class PdfShorthandDeviceCodeAnnotation : PdfAnnotation
        {
            public PdfShorthandDeviceCodeAnnotation(PdfPage page, string short_device_code, PdfAnnotation box_annot)
            {
                var rot = GetRotation(box_annot);
                PdfArray tranformed_arr = TransformTextRectangle(box_annot);
                if (tranformed_arr == null) throw new Exception("failed to resolve pdf annotation");

                PdfDictionary border_props = new PdfDictionary();
                border_props.Elements.Add("/S", new PdfName("/S"));
                border_props.Elements.Add("/Type", new PdfName("/Border"));
                border_props.Elements.Add("/W", new PdfInteger(0));

                /* foreach (var r in (PdfDictionary))
                {
                    Console.WriteLine(string.Format("{0} : {1}", r.Key, r.Value));
                } */


                /* PdfDictionary rot_props = new PdfDictionary();

                rot_props.Elements.Add("/N", new PdfInteger(15822));
                rot_props.Elements.Add();

                << / N 15822 0 R >> */

                string font_size = "4";

                var text_annot_element_fields = new Dictionary<string, object>()
                 {
                    { "/Subj", new PdfString("Short Hand Device Code") },
                    { "/Subtype", new PdfName("/FreeText") },
                    // { "/P", new PdfReference(page) },
                    { "/DA", new PdfString("0 0 0 rg / Helv " + font_size + " Tf") },
                    { "/Rect", tranformed_arr },
                    { "/Contents", new PdfString(short_device_code) },
                    { "/DS", new PdfString("font: Helvetica " + font_size + "pt; text - align:left; margin: 3pt; line - height:13.8pt; color:#FF0000") },
                    { "/BS", border_props },
                    { "/Rotation", new PdfInteger(rot) },
                    /* { "/F", new PdfInteger(4) },
                    { "/RC", "<?xml version=\"1.0\"?><body xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/" +
                        "1.0/\" xfa:contentType=\"text/html\" xfa:APIVersion=\"BluebeamPDFRevu:2018\" xfa:spec" +
                        "=\"2.2.0\" style=\"font:Helvetica " + font_size + "pt; text-align:left; margin:3pt; line-height:" +
                        "13.8pt; color:#000000\" xmlns=\"http://www.w3.org/1999/xhtml\"><p>" + short_device_code + "</p></body>"} */
                 };


                foreach (var f in text_annot_element_fields)
                {
                    bool added = Elements.TryAdd<string, PdfItem>(f.Key, f.Value as PdfItem);
                    if (!added) throw new Exception("Failed to add field to text annotation: " + f.Key);
                }
            }

            private static PdfArray TransformTextRectangle(PdfAnnotation box_annot)
            {
                bool s = box_annot.Elements.TryGetValue("/Rect", out var item);
                if (!s) return null;
                PdfArray rect_arr = item as PdfArray;

                List<double> coords = rect_arr.ToList().Select(x => (x as PdfReal).Value).ToList();
                var ret = new PdfArray();
                ret.Elements.Add(new PdfReal(coords[0])); // bottom left x
                ret.Elements.Add(new PdfReal(coords[1])); // bootom left y
                ret.Elements.Add(new PdfReal(coords[2])); // top right x
                ret.Elements.Add(new PdfReal(coords[3])); // top right y
                return ret;
            }

            private static int GetRotation(PdfAnnotation a)
            {
                bool s = a.Elements.TryGetValue("/Rotation", out var val);
                if (!s) throw new Exception("Failed to get rotation of markup");
                return (val as PdfInteger).Value;
            }
        }

        private static void PlaceTextAnnotationTagForAnnotation(
            PdfPage page, PdfAnnotation a, string short_device_code)
        {
            var txt = new PdfShorthandDeviceCodeAnnotation(page, short_device_code, a);
            page.Annotations.Add(txt);

            foreach (var el in txt.Elements)
            {
                Console.WriteLine(string.Format("{0} : {1}", el.Key, el.Value.ToString()));
            }

            /* bool s = a.Elements.TryGetValue("/Rect", out var item);
            if (!s) return;
            PdfArray arr = item as PdfArray;

            List<double> coords = arr.ToList().Select(x => (x as PdfReal).Value).ToList();
            var gfx = XGraphics.FromPdfPage(page);
            XFont font = new XFont("Helvetica", 4, XFontStyle.Bold);
            gfx.DrawString(short_device_code, font, XBrushes.Black,
                new XRect(coords[0], coords[1], coords[2] - coords[0], coords[3] - coords[1]),
                XStringFormats.CenterLeft);
            Console.WriteLine("drawing text graphic"); */
        }
    }
}