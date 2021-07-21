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
    public static class PdfManager 
    {
        public static IEnumerable<P3BluebeamFDFMarkup> EditBluebeamMarkups(string pdf_path)
        {
            List<P3BluebeamFDFMarkup> markups = new List<P3BluebeamFDFMarkup>();
            PdfDocument doc = PdfReader.Open(pdf_path, PdfDocumentOpenMode.Import);
            PdfSharp.Pdf.PdfPage page = new PdfSharp.Pdf.PdfPage();

            for (int i = 0; i < doc.PageCount; i++)
            {
                page = doc.Pages[i];

                for (int p = 0; p < page.Annotations.Elements.Count; p++)
                {
                    PdfAnnotation text_annot = page.Annotations[p];
                    if (text_annot != null && text_annot.Subject.ToLower().Contains("box"))
                    {
                        if(!text_annot.Elements["/Subtype"].Equals("/FreeText")) continue;

                        var props = GetAnnotationPropertyDict(text_annot);
                        var markup = P3BluebeamFDFMarkup.ParseMarkup(props);
                        markups.Add(markup);
                    }
                }
            }

            return markups;
        }

        public static void CreateCopyDocument(string origin_pdf_path, string copy_pdf_path, IEnumerable<P3BluebeamFDFMarkup> markups)
        {
            PdfDocument origin = PdfReader.Open(origin_pdf_path, PdfDocumentOpenMode.Import);
            PdfDocument copy = new PdfDocument();
            
            for(int Pg = 0; Pg < origin.Pages.Count; Pg++)
            {
                var origin_page = origin.Pages[Pg];
                PushDeviceCodes(origin_page);
                copy.AddPage(origin_page);
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var enc1252 = Encoding.GetEncoding(1252);

            copy.Save(copy_pdf_path);
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
                    /* foreach(var el in annot.Elements)
                    {
                        Console.WriteLine(string.Format("{0} : {1}", el.Key, el.Value));
                    }
                    Console.ReadKey(); */

                    annot.Elements.Remove("/AP");

                    if(!annot.Elements["/Subtype"].Equals("/FreeText")) continue;

                    var props = GetAnnotationPropertyDict(annot);
                    var markup = P3BluebeamFDFMarkup.ParseMarkup(props);
                    annot.Contents = markup.DeviceCode;
                    PdfTextAnnotation tannot = annot as PdfTextAnnotation;

                }
            }
        }

        // The order that the parse string come in is as follows:
        // Box Type - Box Size - Connector Size - CT - CB - PT - PB - MT- MB - GANG - Plaster Ring Size - Entry Connector Type
        private static Dictionary<string, string> GetAnnotationPropertyDict(PdfAnnotation annot)
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