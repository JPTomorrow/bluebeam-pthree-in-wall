using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JPMorrow.Bluebeam.Markup;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

namespace JPMorrow.PDF
{
    public class BundleRegion 
    {
        public string Subject { get; private set; }
        public string GroupingText { get; private set; }
        public string BundleName { get; private set; }
        public string RectString { get; private set; }

        public BundleRegion(PdfAnnotation annot) {
            var props = annot.Elements;

            string trim(PdfItem x) => x.ToString().Trim('(', ')', ' ');

            Subject = trim(props["/Subj"]);
            BundleName = PdfManager.GetBISColumnData(annot).Last();
            GroupingText = props["/GroupNesting"].ToString();
            RectString = props["/Rect"].ToString();
        }

        private class AnnotationXYBoundingBox
        {
            public class Point 
            {
                public double X { get; private set; }
                public double Y { get; private set; }

                public Point(double x , double y) 
                {
                    X = x;
                    Y = y;
                }

                public override string ToString()
                {
                    return string.Format("({0}, {1})", X, Y);
                }
            }

            public Point TopLeft { get; private set; }
            public Point BottomRight { get; private set; }

            public AnnotationXYBoundingBox(string rectangle_str) {
                string[] rect_coord_strs = rectangle_str.Trim('[', ']', ' ').Split(' ');
                double[] rect_coords = rect_coord_strs.Select(x => double.Parse(x)).ToArray();
                TopLeft = new Point(rect_coords[0], rect_coords[1]);
                BottomRight = new Point(rect_coords[2], rect_coords[3]);
            }

            private bool Overlap(Point l1, Point r1, Point l2, Point r2)
            {
                // To check if either rectangle is actually a line
                // For example :  l1 = {-1,0}  r1 = {1,1}  l2 = {0,-1}
                // r2 = {0,1}

                if (l1.X == r1.X || l1.Y == r1.Y || l2.X == r2.X
                    || l2.Y == r2.Y) {
                    // the line cannot have positive overlap
                    return false;
                }
            
                // If one rectangle is on left side of other
                if (l1.X >= r2.X || l2.X >= r1.X)
                    return false;
            
                // If one rectangle is above other
                if (r1.Y >= l2.Y || r2.Y >= l1.Y)
                    return false;
            
                return true;
            }

            public bool DoOverlap(AnnotationXYBoundingBox other) {
                Console.WriteLine("topmost: " + TopLeft.ToString());
                Console.WriteLine("bottomright: " + BottomRight.ToString());
                Console.WriteLine("other.topmost: " + other.TopLeft.ToString());
                Console.WriteLine("other.bottomright: " + other.BottomRight.ToString());
                Console.WriteLine("");
                return Overlap(TopLeft, BottomRight, other.TopLeft, other.BottomRight);
            }
        }

        public bool isInRegion(PdfAnnotation text_annot, IEnumerable<PdfAnnotation> all_annots) 
        {

            var annots = all_annots.Where(x => x != null).ToList();
            var txt_grouping = text_annot.Elements["/GroupNesting"].ToString();

            var idx = annots.FindIndex(x => txt_grouping.Contains(x.Elements["/NM"].ToString().Trim('(', ')', ' ')) && x.Elements["/Subtype"].Equals("/Square"));

            if(idx == -1) return false;
            else // found box now get rect
            {
                var box_rect_str = annots[idx].Elements["/Rect"].ToString();
                AnnotationXYBoundingBox box_bb = new AnnotationXYBoundingBox(box_rect_str);
                AnnotationXYBoundingBox bundle_bb = new AnnotationXYBoundingBox(RectString);
                Console.WriteLine("");
                Console.WriteLine(annots[idx].Elements["/Subj"].ToString() + "\n" + annots[idx].Elements["/Subtype"].ToString());
                if(bundle_bb.DoOverlap(box_bb)) return true;
            }
            
            return false; 
        }

        public bool isBundleRegionText(PdfAnnotation annot)
        {
            var group_text = annot.Elements["/NM"].ToString().Trim('(', ')', ' ');
            if(GroupingText.Contains(group_text)) return true;
            return false;
        }
    }

    public static class PdfManager 
    {
        public static IEnumerable<P3BluebeamFDFMarkup> EditBluebeamMarkups(string pdf_path)
        {
            List<P3BluebeamFDFMarkup> markups = new List<P3BluebeamFDFMarkup>();
            PdfDocument doc = PdfReader.Open(pdf_path, PdfDocumentOpenMode.Import);
            PdfSharp.Pdf.PdfPage page = new PdfSharp.Pdf.PdfPage();
            string[] subject_box_names = new[] { "power", "data", "lighting", "fire alarm" };

            // collect all bundle zones in the document
            var regions = GetBundleRegions(doc).ToList();

            // get page annoations and parse them
            for (int i = 0; i < doc.PageCount; i++)
            {
                page = doc.Pages[i];

                for (int p = 0; p < page.Annotations.Elements.Count; p++)
                {
                    PdfAnnotation text_annot = page.Annotations[p];

                    if (text_annot != null && subject_box_names.Any(x => text_annot.Subject.ToLower().Contains(x + " box")))
                    {
                        if(!text_annot.Elements["/Subtype"].Equals("/FreeText")) continue;

                        var props = GetAnnotationPropertyDict(text_annot);
                        P3BluebeamFDFMarkup markup = null;

                        var idx = regions.FindIndex(x => x.isInRegion(text_annot, page.Annotations.ToList().Select(x => x as PdfAnnotation)));
                        if(idx == -1)
                        {
                            markup = P3BluebeamFDFMarkup.ParseMarkup(props);
                        } 
                        else
                        {
                            markup = P3BluebeamFDFMarkup.ParseMarkup(props, regions[idx]);
                        }

                        markups.Add(markup);
                    }
                }
            }

            return markups;
        }

        private static IEnumerable<BundleRegion> GetBundleRegions(PdfDocument doc) 
        {
            List<BundleRegion> regions = new List<BundleRegion>();
            PdfSharp.Pdf.PdfPage page = new PdfSharp.Pdf.PdfPage();

            for (int i = 0; i < doc.PageCount; i++)
            {
                page = doc.Pages[i];

                for (int p = 0; p < page.Annotations.Elements.Count; p++)
                {
                    PdfAnnotation annot = page.Annotations[p];

                    if (annot != null && annot.Subject.ToLower().Contains("bundle designation zone"))
                    {
                        if(!annot.Elements["/Subtype"].Equals("/Square")) continue;
                        regions.Add(new BundleRegion(annot));
                    }
                }
            }

            return regions;
        }

        public static void CreateCopyDocument(string origin_pdf_path, string copy_pdf_path, IEnumerable<P3BluebeamFDFMarkup> markups)
        {
            PdfDocument origin = PdfReader.Open(origin_pdf_path, PdfDocumentOpenMode.Import);
            PdfDocument copy = new PdfDocument();

            // collect all bundle zones in the document
            var regions = GetBundleRegions(origin).ToList();

            for(int Pg = 0; Pg < origin.Pages.Count; Pg++)
            {
                var origin_page = origin.Pages[Pg];
                PushDeviceCodes(origin_page);
                PushBundleNames(origin_page, regions);
                copy.AddPage(origin_page);
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var enc1252 = Encoding.GetEncoding(1252);

            copy.Save(copy_pdf_path);
        }

        public static void PushBundleNames(PdfPage page, IEnumerable<BundleRegion> regions)
        {
            var r = regions.ToList();
            for (int p = 0; p < page.Annotations.Elements.Count; p++)
            {
                PdfAnnotation annot = page.Annotations[p];

                if (annot != null && annot.Subject.ToLower().Contains("bundle designation zone"))
                {
                    if (!annot.Elements["/Subtype"].Equals("/FreeText")) continue;
                    var idx = r.FindIndex(x => x.isBundleRegionText(annot));

                    if (idx == -1) continue;
                    else
                    {
                        var region = r[idx];
                        annot.Elements.Remove("/AP");
                        annot.Contents = region.BundleName;
                        PdfTextAnnotation tannot = annot as PdfTextAnnotation;
                    }
                }
            }
        }

        /// <summary>
        /// should not be here but needed to get the P3 in wall running
        /// </summary>
        /// <returns></returns>
        public static void PushDeviceCodes(PdfPage page)
        {
            for (int p = 0; p < page.Annotations.Elements.Count; p++)
            {
                PdfAnnotation annot = page.Annotations[p];

                // is text annotation
                if (annot != null && annot.Subject.ToLower().Contains("box"))
                {
                    string[] subject_box_names = new[] { "power", "data", "lighting", "fire alarm" };
                    if(!annot.Elements["/Subtype"].Equals("/FreeText") || !subject_box_names.Any(x => annot.Subject.ToLower().Contains(x + " box"))) continue;
                    annot.Elements.Remove("/AP");

                    var props = GetAnnotationPropertyDict(annot);
                    var markup = P3BluebeamFDFMarkup.ParseMarkup(props);
                    annot.Contents = markup.DeviceCode;
                    PdfTextAnnotation tannot = annot as PdfTextAnnotation;
                }
            }
        }

        public static IEnumerable<string> GetBISColumnData(PdfAnnotation annot)
        {
            var column_str = "";

            var els = annot.Elements;
            foreach(var el in els)
            {
                if(el.Key.Equals("/BSIColumnData")) {
                    column_str = el.Value.ToString();
                }
            }
            
            if(string.IsNullOrWhiteSpace(column_str)) return null;

            var split_props = column_str.Split(") (").ToList();

            split_props = split_props
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim('[', '(', ')', ']', ' ')).ToList();

            return split_props;
        }

        // The order that the parse string come in is as follows:
        // Box Type - Box Size - Connector Size - CT - CB - PT - PB - MT- MB - GANG - Plaster Ring Size - Entry Connector Type
        private static Dictionary<string, string> GetAnnotationPropertyDict(PdfAnnotation annot)
        {
            var split_props = GetBISColumnData(annot).ToList();

            var dict = new Dictionary<string, string>() {
                {"Box Type"                  , split_props[0] },
                {"Box Size"                  , split_props[1] },
                {"Gang"                      , split_props[9] },
                {"Plaster Ring"              , split_props[10] },
                {"Conduit Entry Connector"   , split_props[11] },
                {"Connector Size"            , split_props[2] },
                {"Top EMT Connectors"        , split_props[3] },
                {"Bottom EMT Connectors"     , split_props[4] },
                {"Top PVC Connectors"        , split_props[5] },
                {"Bottom PVC Connectors"     , split_props[6] },
                {"Top MC Connectors"         , split_props[7] },
                {"Bottom MC Connectors"      , split_props[8] },
            };

            // get subject
            dict.Add("Subject", annot.Subject);

            return dict;
        }

        private static string GetBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }
    }
}