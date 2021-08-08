namespace JPMorrow.Test.Console
{
    using JPMorrow.PDF;
    using System;
    using System.IO;
    using System.Linq;

    public static partial class TestBed
    {
        public static void TestPdforgeConstructor(string exe_path, TestAssert a)
        {
            var input_path = exe_path + "Test Document 1.pdf";
            var output_path = exe_path + "Test Document 1 (OUTPUT).pdf";

            Pdforge f = new Pdforge(input_path, output_path);
            a.Assert("Check if input file exists", File.Exists(f.InputFilepath));
            a.Assert("Check if output file exists", Directory.Exists(Path.GetDirectoryName(f.OutputFilepath)));
        }

        public static void TestPdforgeDocumentCleanup(string exe_path, TestAssert a)
        {
            var input_path = exe_path + "Test Document 1.pdf";
            var copy_input_path = exe_path + Path.GetFileNameWithoutExtension(input_path) + " test copy.pdf";
            var output_path = exe_path + "Test Document 1 (OUTPUT).pdf";

            Pdforge f = new Pdforge(input_path, output_path);
            a.Assert("Check if input file exists", File.Exists(f.InputFilepath));
            a.Assert("Check if output file exists", Directory.Exists(Path.GetDirectoryName(f.OutputFilepath)));
            
            f.DeleteInputFile();
            a.Assert("Delete input file", !File.Exists(f.InputFilepath));

            f.SavePdfToLocation(input_path);
            a.Assert("Save copy of input document to input path location", File.Exists(f.InputFilepath));

            f.SavePdfToLocation(copy_input_path);
            a.Assert("Save copy of input document to copy input path location", File.Exists(copy_input_path));

            f.DeleteOutputFile(copy_input_path);
            a.Assert("Delete copy input file", !File.Exists(copy_input_path));

            f.SaveToOutputPdfLocation();
            a.Assert("Save copy of input document to Pdfforge stored input file location", File.Exists(f.OutputFilepath));

            f.DeleteOutputFile();
            a.Assert("Delete output file", !File.Exists(f.OutputFilepath));
        }

        public static void TestPdforgeGetAnnotationsBySubject(string exe_path, TestAssert a)
        {
            var input_path = exe_path + "Test Document 1.pdf";
            var copy_input_path = exe_path + Path.GetFileNameWithoutExtension(input_path) + " test copy.pdf";
            var output_path = exe_path + "Test Document 1 (OUTPUT).pdf";

            Pdforge f = new Pdforge(input_path, output_path);
            
            a.Assert("Check if input file exists", File.Exists(f.InputFilepath));
            a.Assert("Check if output file exists", Directory.Exists(Path.GetDirectoryName(f.OutputFilepath)));

            var page = f.GetPage(0);
            var page_annots_count = page.Annotations.Count();
            a.Assert("Initial page annotations count", page_annots_count > 0);

            var found_annots = f.GetAnnotationsBySubject(page, "Power Box", true).ToList();
            a.Assert("Found annotations count", found_annots.Count() > 0);
        }
    }
}