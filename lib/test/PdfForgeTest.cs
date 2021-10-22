namespace JPMorrow.Test.Console
{
    using JPMorrow.Measurements;
    using JPMorrow.PDF;
    using JPMorrow.Revit.Labor;
    using System;
    using System.Diagnostics;
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

        public static void TestPdforgeSetAnnotationProperty(string exe_path, TestAssert a)
        {
            var input_path = exe_path + "Test Document 1.pdf";
            var output_path = exe_path + "Test Document 1 (OUTPUT).pdf";
            var annotation_test_value = "TESTTESTTESTTEST123abcxyzqwe";

            Pdforge f = new Pdforge(input_path, output_path);

            a.Assert("Check if input file exists", File.Exists(f.InputFilepath));
            a.Assert("Check if output file exists", Directory.Exists(Path.GetDirectoryName(f.OutputFilepath)));

            var page = f.GetPage(0);
            var page_annots_count = page.Annotations.Count();
            a.Assert("Initial page annotations count", page_annots_count > 0);

            var found_annots = f.GetAnnotationsBySubject(page, "Power Box", true).ToList();
            f.SetAnnotationStringPropertys(found_annots, "Subj", annotation_test_value);

            f.SaveToOutputPdfLocation();
            a.Assert("Save copy of input document to Pdfforge stored input file location", File.Exists(f.OutputFilepath));

            Pdforge f2 = new Pdforge(output_path, output_path);
            var test_annots = f2.GetAnnotationsBySubject(page, annotation_test_value, true).ToList();
            a.Assert("Found annotations count", test_annots.Count() > 0);

            f2.DeleteOutputFile();
            a.Assert("Delete output file", !File.Exists(f.OutputFilepath));
        }

        public static void TestMeasurements(string exe_path, TestAssert a)
        {
            var test_sizes = new string[] { "1/2\"", "1 1/2\"", "1'", "1' 1\"", "1' 1/2\"", "1' 1 1/2\"" };

            foreach (var size in test_sizes)
            {
                var d1 = Measure.LengthDbl(size);
                var s1 = Measure.LengthFromDbl(d1);
                var d2 = Measure.LengthDbl(s1);
                var s2 = Measure.LengthFromDbl(d2);
                var assert_str = "Convert back and forth from double to string: " +
                size + " { " + d1.ToString() + ", " + d2.ToString() + ", " + s1 + ", " + s2 + " }";
                a.Assert(assert_str, d1 == d2 && s1.Equals(s2) && s2.Equals(size));
            }

        }

        public static void TestLaborHourEntriesFromResourceFile(string exe_path, TestAssert a)
        {
            var entries = LaborExchange.LoadLaborFromInternalRescource();
            a.Assert("Labor entries are empty or null -> entries: " + entries.Count(), entries.Any());
            LaborExchange lex = new LaborExchange(entries);
            a.Assert("Problem with labor exchange", lex == null || !lex.Items.Any());
        }
    }
}