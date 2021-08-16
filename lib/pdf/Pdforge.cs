using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

namespace JPMorrow.PDF {
    public class Pdforge
    {
        public string InputFilepath { get; private set; } = "";
        public string OutputFilepath { get; private set; } = "";
        public string InputFilename { get => Path.GetFileName(InputFilepath); }
        public string OutputFilename { get => Path.GetFileName(OutputFilepath); }

        private PdfDocument _inputDocument { get; set; }
        public PdfDocument InputDocument { get {
            if(_inputDocument == null)
                throw new FileNotFoundException("The Input PDF file is not valid.");
            return _inputDocument;
        } }

        public Pdforge(string input_filepath, string output_filepath) 
        {
            InputFilepath = input_filepath;
            OutputFilepath = output_filepath;

            if(!File.Exists(InputFilepath))
                throw new FileNotFoundException("Input file path does not resolve to a valid PDF file", InputFilename);

            var out_dir = Path.GetDirectoryName(OutputFilepath);

            if(!Directory.Exists(out_dir))
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
                
                if(!s) continue;
                var str = item.ToString().Trim(' ', '[', ']', '(', ')');
                
                bool has_value = contains_value ? str.Contains(annotation_element_value) : str.Equals(annotation_element_value);
                if(has_value) results.Add(annot);
                else results.Add(annot);
            }

            return results;
        }

        public void SetAnnotationStringProperty(PdfAnnotation annotation, string property_name, string property_value)
        {
            var pname = property_name.StartsWith("/") ? property_name : "/" + property_name;
            bool s = annotation.Elements.TryGetValue(pname, out PdfItem item);
            if(!s) return;
            annotation.Elements.SetValue(pname, new PdfString(property_value));
        }

        public void SetAnnotationStringPropertys(IEnumerable<PdfAnnotation> annotations, string property_name, string property_value)
        {
            foreach(var annot in annotations)
            {
                SetAnnotationStringProperty(annot, property_name, property_value);
            }
        }

        

        public void PrintAllFirstPageData()
        {
            
            // var bsi_columns = InputDocument.
            var data = InputDocument.Internals.GetAllObjects();
            Console.WriteLine("data count: ", data.Count().ToString());
            foreach(var v in data)
            {
                Console.WriteLine(string.Format("{0}", v.ToString()));
                Console.WriteLine();
                // Console.WriteLine(string.Format("{0} : {1}", v.Key, v.Value));
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

            for(int Pg = 0; Pg < InputDocument.Pages.Count; Pg++)
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
            if(!File.Exists(InputFilepath)) return false;

            while(FileIsLocked(InputFilepath, FileAccess.ReadWrite))
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
            if(output_path == null) output_path = OutputFilepath;

            if(!File.Exists(output_path)) return false;

            while(FileIsLocked(output_path, FileAccess.ReadWrite))
            {
                Console.WriteLine("Waiting for PDF output file lock.");
                Thread.Sleep(1000);
            }

            File.Delete(output_path);
            return true;
        }

        public void PrintAllElementProperies(PdfPage page, string txt_file_path)
        {
            var o = "";

            foreach(PdfAnnotation annot in page.Annotations)
            {
                if(annot == null) continue;
                var elements = annot.Elements;

                foreach( var p in elements)
                {
                    o += string.Format("{0} : {1}", p.Key, p.Value) + "\n";
                }

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
    }

    
}