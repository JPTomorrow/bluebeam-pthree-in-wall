using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;
using static PdfSharp.Pdf.PdfArray;

namespace JPMorrow.PDF
{
    public class Pdforge
    {
        public string InputFilepath { get; private set; } = "";
        public string OutputFilepath { get; private set; } = "";
        public string InputFilename { get => Path.GetFileName(InputFilepath); }
        public string OutputFilename { get => Path.GetFileName(OutputFilepath); }

        private PdfDocument _inputDocument { get; set; }
        public PdfDocument InputDocument
        {
            get
            {
                if (_inputDocument == null)
                    throw new FileNotFoundException("The Input PDF file is not valid.");
                return _inputDocument;
            }
        }

        public List<PdfAnnotation> AllAnnotations { get => GetPage(0).Annotations.Select(x => x as PdfAnnotation).ToList(); }

        public Pdforge(string input_filepath, string output_filepath)
        {
            InputFilepath = input_filepath;
            OutputFilepath = output_filepath;

            if (!File.Exists(InputFilepath))
                throw new FileNotFoundException("Input file path does not resolve to a valid PDF file", InputFilename);

            var out_dir = Path.GetDirectoryName(OutputFilepath);

            if (!Directory.Exists(out_dir))
                throw new FileNotFoundException("Output file path does not resolve to a valid location", OutputFilename);

            _inputDocument = PdfReader.Open(InputFilepath, PdfDocumentOpenMode.Import);
        }

        public PdfPage GetPage(int page_idx)
        {
            return InputDocument.Pages[page_idx];
        }

        public IEnumerable<PdfAnnotation> GetAnnotationsBySubject(PdfPage page, string subject, bool contains)
        {
            var annots = GetAnnotationsByElementProperty(page, "Subj", subject, contains);
            return annots;
        }

        public IEnumerable<PdfAnnotation> GetAnnotationsByContents(PdfPage page, string contents_name, bool contains_value)
        {
            return null; // GetAnnotationsByElementProperty(page, "/Contents");
        }

        public IEnumerable<PdfAnnotation> GetAnnotationsBySubType(PdfPage page, string sub_type)
        {
            var st = sub_type.StartsWith("/") ? sub_type : "/" + sub_type;
            var annots = GetAnnotationsByElementProperty(page, "Subtype", st, false);
            return annots;
        }

        private IEnumerable<PdfAnnotation> GetAnnotationsByElementProperty(
            PdfPage page, string annotation_element_name,
            string annotation_element_value, bool contains_value)
        {
            var el_name = annotation_element_name.StartsWith('/') ? annotation_element_name : "/" + annotation_element_name;
            List<PdfAnnotation> results = new List<PdfAnnotation>();

            for (var i = 0; i < page.Annotations.Count; i++)
            {
                var annot = page.Annotations[i];
                if (annot == null) continue;

                bool s = annot.Elements.TryGetValue(el_name, out var item);

                if (!s) continue;
                var str = item.ToString().Trim(' ', '[', ']', '(', ')');

                bool has_value = contains_value ? str.Contains(annotation_element_value) : str.Equals(annotation_element_value);
                if (has_value) results.Add(annot);
                else results.Add(annot);
            }

            return results;
        }

        public void SetAnnotationStringProperty(PdfAnnotation annotation, string property_name, string property_value)
        {
            var pname = property_name.StartsWith("/") ? property_name : "/" + property_name;
            bool s = annotation.Elements.TryGetValue(pname, out PdfItem item);
            if (!s) return;
            annotation.Elements.SetValue(pname, new PdfString(property_value));
        }

        public void SetAnnotationStringPropertys(IEnumerable<PdfAnnotation> annotations, string property_name, string property_value)
        {
            foreach (var annot in annotations)
            {
                SetAnnotationStringProperty(annot, property_name, property_value);
            }
        }

        public void PrintAllFirstPageData()
        {

            // var bsi_columns = InputDocument.
            var data = InputDocument.Internals.GetAllObjects();
            PdfCatalog pdf_catalog = InputDocument.Internals.Catalog;

            Console.WriteLine("Catalog Data:\n");
            foreach (var v in pdf_catalog.Elements)
            {
                Console.WriteLine(string.Format("{0} : {1}", v.Key, v.Value));
            }

            Console.WriteLine("data count: ", data.Count().ToString());
            foreach (var v in data)
            {
                Console.WriteLine(string.Format("{0}", v.ToString()));
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Copy the input document and save the copy to the output file path
        /// </summary>
        public void SaveToOutputPdfLocation()
        {
            SavePdfToLocation(OutputFilepath);
        }

        public void SavePdfToLocation(string output_path)
        {
            PdfDocument copy = new PdfDocument();

            for (int Pg = 0; Pg < InputDocument.Pages.Count; Pg++)
            {
                var origin_page = InputDocument.Pages[Pg];
                copy.AddPage(origin_page);
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var enc1252 = Encoding.GetEncoding(1252);

            copy.Save(output_path);
        }

        /// <summary>
        /// Delete the input pdf document file on disk
        /// </summary>
        public bool DeleteInputFile()
        {
            if (!File.Exists(InputFilepath)) return false;

            while (FileIsLocked(InputFilepath, FileAccess.ReadWrite))
            {
                Console.WriteLine("Waiting for PDF input file lock.");
                Thread.Sleep(1000);
            }

            File.Delete(InputFilepath);
            return true;
        }

        /// <summary>
        /// Delete the PDF output file on disk
        /// </summary>
        public bool DeleteOutputFile(string output_path = null)
        {
            if (output_path == null) output_path = OutputFilepath;

            if (!File.Exists(output_path)) return false;

            while (FileIsLocked(output_path, FileAccess.ReadWrite))
            {
                Console.WriteLine("Waiting for PDF output file lock.");
                Thread.Sleep(1000);
            }

            File.Delete(output_path);
            return true;
        }

        public void PrintAllPdfProperies(string txt_file_path)
        {
            var o = "";

            var data = InputDocument.Internals.GetAllObjects();
            var annots = GetPage(0).Annotations;

            o += "data count: " + data.Count().ToString() + "\n\n";
            foreach (var v in data)
            {
                var s = PrintFormatInternalObject(v);
                foreach (var arr in s)
                {
                    foreach (var kvp in arr)
                    {
                        o += string.Format("{0} : {1}\n", kvp.Key, kvp.Value);
                    }
                    if (s.Any()) o += "\n";
                }
            }

            foreach (var a in annots)
            {
                var els = (a as PdfAnnotation).Elements;
                foreach (var e in els)
                    o += string.Format("{0} : {1} -> {2}\n", e.Key, e.Value.ToString(), e.Value.GetType().ToString());
                o += "\n";
            }

            File.WriteAllText(txt_file_path, o);
        }

        /// <summary>
        /// Check if qfile is locked
        /// </summary>
        private static bool FileIsLocked(string filename, FileAccess file_access)
        {
            // Try to open the file with the indicated access.
            try
            {
                FileStream fs =
                    new FileStream(filename, FileMode.Open, file_access);
                fs.Close();
                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public IEnumerable<KeyValuePair<string, string>[]> PrintFormatInternalObject(PdfObject o)
        {
            // handle each type individually
            var type = o.GetType();
            var ret = new List<KeyValuePair<string, string>[]>();

            if (type == typeof(PdfArray))
            {
                // ret.Add(type.ToString());
                var arr = (o as PdfArray).Elements;

                for (int i = 0; i < arr.Count(); i++)
                {
                    var el = arr[i];
                    var tt = el.GetType();
                    if (tt != typeof(PdfDictionary)) continue;
                    List<KeyValuePair<string, string>> f = new List<KeyValuePair<string, string>>();
                    foreach (var p in (el as PdfDictionary).ToList())
                        f.Add(new KeyValuePair<string, string>(p.Key, p.Value.ToString()));
                    ret.Add(f.ToArray());
                }
            }
            else if (type == typeof(PdfAnnotation))
            {
                var els = (o as PdfAnnotation).Elements;
                var arr = els.Select(x => new KeyValuePair<string, string>(x.Key, x.Value.ToString()));
                ret.Add(arr.ToArray());
            }

            return ret;
        }

        public static bool TransferBSIColumnData(PdfAnnotation a, PdfAnnotation b)
        {
            bool has_BSI = a.Elements.TryGetValue("/BSIColumnData", out var d1);
            if (!has_BSI) return false;
            bool has_BSI2 = b.Elements.TryGetValue("/BSIColumnData", out var d2);
            if (has_BSI2) return false;
            b.Elements.SetValue("/BSIColumnData", d1);
            return true;
        }

        public static IEnumerable<PdfAnnotation> GetChildAnnotations(PdfPage page, PdfAnnotation a)
        {
            var ret = new List<PdfAnnotation>();
            bool has_nesting = a.Elements.TryGetValue("/GroupNesting", out var item);
            if (!has_nesting) return ret;

            PdfArray arr = item as PdfArray;
            List<string> group_ids = new List<string>();
            foreach (var aa in arr.Elements)
            {
                if (aa.GetType() == typeof(PdfString))
                    group_ids.Add((aa as PdfString).Value);
            }

            var annots = page.Annotations.Select(x => x as PdfAnnotation).ToList();
            foreach (var pa in annots)
            {
                bool has_group_number = pa.Elements.TryGetString("/NM", out string gn);
                if (!has_group_number) continue;
                if (group_ids.Any(x => x.Equals(gn))) ret.Add(pa);
            }

            return ret;
        }

        /// <summary>
        /// Stupid hack to get PdfSharop to recognize the pages in a sheet you have imported.
        /// </summary>
        public static void RefreshSheetPagesHack(PdfDocument doc)
        {
            int pages = doc.PageCount; // updates the page resgistry for some reason
        }
    }

    /// <summary>
    /// Custom Column header data for matching up BISColumnData
    /// </summary>
    public class PdfCustomColumnCollection
    {
        private class PdfCustomColumn
        {
            public int DisplayOrder { get; private set; } = -1;
            public string SubType { get; private set; } = string.Empty;
            public string HeaderName { get; private set; } = string.Empty;
            public string DefaultValue { get; private set; } = string.Empty;

            public PdfCustomColumn(KeyValuePair<string, string>[] custom_column_data)
            {
                foreach (var p in custom_column_data)
                {
                    if (p.Key.Equals(customColumnStrings[0])) DisplayOrder = int.Parse(p.Value);
                    else if (p.Key.Equals(customColumnStrings[1])) HeaderName = p.Value.Trim('(', ')');
                    else if (p.Key.Equals(customColumnStrings[2])) DefaultValue = p.Value.Trim('(', ')');
                    else if (p.Key.Equals(customColumnStrings[3])) SubType = p.Value.TrimStart('/');
                }
            }

            public override string ToString()
            {
                return string.Format("Column: [ {0}, {1}, {2}, {3} ]", HeaderName, DisplayOrder, SubType, DefaultValue);
            }
        }

        private List<PdfCustomColumn> columns { get; set; } = new List<PdfCustomColumn>();

        private static string[] customColumnStrings = new string[] {
            "/DisplayOrder",
            "/Name",
            "/DefaultValue",
            "/Subtype"
        };

        public PdfCustomColumnCollection(Pdforge forge)
        {
            var all_obj = forge.InputDocument.Internals.GetAllObjects();


            foreach (var obj in all_obj)
            {
                var type = obj.GetType();
                if (type != typeof(PdfArray)) continue;

                // match the keys
                var custom_columns = forge.PrintFormatInternalObject(obj);

                foreach (var arr in custom_columns)
                {
                    if (customColumnStrings.Any(x => arr.Select(x => x.Key).ToList().Contains(x)))
                        columns.Add(new PdfCustomColumn(arr));
                }
            }
        }

        public override string ToString()
        {
            string o = string.Join("\n", columns.Select(x => x.ToString()));
            return o;
        }

        public bool HasExpectedColumnHeaders(string[] column_header_names, out List<string> missing_names)
        {
            missing_names = new List<string>();

            foreach (var c in column_header_names)
            {
                if (columns.Any(x => x.HeaderName.Equals(c))) continue;
                missing_names.Add(c);
            }

            if (missing_names.Any()) return false;
            else return true;
        }

        public Dictionary<string, string> MatchColumnHeaderToBISData(PdfAnnotation a)
        {
            var dict = new Dictionary<string, string>();

            bool s = a.Elements.TryGetValue("/BSIColumnData", out var item);
            if (!s) return dict;
            var bis_data = (item as PdfArray).ToArray()
            .Select(x => x.ToString().Trim('(', ')')).ToArray();

            foreach (var c in columns)
            {
                if (c.DisplayOrder == -1) continue;
                var val = bis_data[c.DisplayOrder];
                dict.Add(c.HeaderName, val);
            }

            return dict;
        }

        public PdfArray ModifyBSIData(PdfDocument doc, PdfAnnotation a, string column_header, string value)
        {
            var bis_data = (a.Elements["/BSIColumnData"] as PdfArray).ToArray();
            var new_data = new List<PdfItem>();

            int ch_idx = columns.FindIndex(x => x.HeaderName.Equals(column_header));
            if (ch_idx == -1) throw new Exception("Provided column header does not exist");
            var col = columns[ch_idx];

            foreach (var d in bis_data)
            {
                var idx = bis_data.ToList().IndexOf(d);
                if (idx == col.DisplayOrder)
                {
                    new_data.Add(new PdfString(value) as PdfItem);
                }
                else
                {
                    new_data.Add(d);
                }
            }

            return new PdfArray(doc, new_data.ToArray());
        }
    }
}