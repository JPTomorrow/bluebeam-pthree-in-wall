using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using JPMorrow.Bluebeam.Markup;
using JPMorrow.Excel;
using JPMorrow.Measurements;
using JPMorrow.P3;
using JPMorrow.Pdf.Bluebeam;
using JPMorrow.Pdf.Bluebeam.FireAlarm;
using JPMorrow.Pdf.Bluebeam.P3;
using JPMorrow.PDF;
using JPMorrow.Test.Console;
using OfficeOpenXml;
using PdfSharp.Pdf.Annotations;

namespace BluebeamP3InWall
{
    class Program
    {
        static void Main(string[] args)
        {
            string exe_path = GetThisExecutablePath();

#if DEBUG
            Console.WriteLine(exe_path + "\n");
            bool passed = TestBedConsole.TestAll(exe_path);
            if (!passed)
            {
                Console.ReadKey();
                return;
            }
#endif

            try
            {
#if FIRE_ALARM
                Console.WriteLine("Fire Alarm");
                RunFireAlarm(exe_path);
#else
                Console.WriteLine("P3 in wall");
                RunP3InWall(exe_path);
#endif
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.ToString());
#else
                Console.WriteLine(ex.Message.ToString());
#endif
                Console.ReadKey();
            }

            /* 
            return; */
        }

        /// <summary>
        /// Get executable path 
        /// </summary>
        public static string GetThisExecutablePath()
        {
            var split_path = System.Diagnostics.Process
                .GetCurrentProcess().MainModule.FileName
                .Split("\\").ToList();

            split_path.Remove(split_path.Last());
            return string.Join("\\", split_path) + "\\";
        }

        public static void PrintAllAnnotationPropertiesToFile(string exe_path, string pdf_doc_path)
        {
#if DEBUG
            var input_path = exe_path + Path.GetFileName(pdf_doc_path);
            var properties_txt_path = exe_path + "annotation_properties.txt";

            Pdforge f = new Pdforge(pdf_doc_path, Path.GetFullPath(input_path));
            f.PrintAllPdfProperies(properties_txt_path);

            Process.Start("notepad.exe", properties_txt_path);
#endif
        }

        /// <summary>
        /// Process the fire alarm from a bluebeam document
        /// </summary>
        /// <param name="exe_path"></param>
        private static void RunFireAlarm(string exe_path)
        {
            string pdf_input_path = GetPdfInputFileName(exe_path, "");

            if (string.IsNullOrWhiteSpace(pdf_input_path) || !File.Exists(pdf_input_path))
            {
                Console.WriteLine("Failed to retrieve pdf file name");
                Console.ReadKey();
                return;
            }
            else
            {
                Console.WriteLine("Using pdf file: " + pdf_input_path);
            }

            PrintAllAnnotationPropertiesToFile(exe_path, pdf_input_path);

            double hanger_spacing = -1;

            do
            {
                Console.Write("\nPlease enter hanger spacing in feet and inches (ex. 0' 0\"): ");
                var spacing_str = Console.ReadLine();
                var s = Measure.LengthDbl(spacing_str);
                if (s < 0) Console.WriteLine("That is an invalid hanger spacing. Hanger spacing should be provided in feet and inches");
                else hanger_spacing = s;
            } while (hanger_spacing < 0);

            // get threaded rod selection
            string[] rod_size_sels = new string[] {
                "1. 1/4\"",
                "2. 3/8\"",
                "3. 1/2\"",
            };

            string threaded_rod_size = null;
            Console.WriteLine();
            do
            {
                Console.WriteLine("Threaded rod sizes:");
                rod_size_sels.ToList().ForEach(x => Console.WriteLine(x));
                Console.Write("\nPlease enter a number (1-3) coresponding to the size of threaded rod you would like to use: ");
                var tr_in = Console.ReadLine();
                var idx = rod_size_sels.ToList().FindIndex(x => x.StartsWith(tr_in));
                if (idx == -1)
                    Console.WriteLine("Selection not recognized. Please provide one of the choice numbers (1-3)");
                else
                    threaded_rod_size = rod_size_sels[idx].Split(' ').Last();
            } while (threaded_rod_size == null);


            var pdf_output_path = exe_path +
                 Path.GetFileNameWithoutExtension(pdf_input_path) +
                 "_fire_alarm_processed.pdf";

            // parse poly lines into conduit package
            Pdforge f = new Pdforge(pdf_input_path, pdf_output_path);
            var poly_lines = f.GetAnnotationsBySubType(f.GetPage(0), "PolyLine");
            BluebeamConduitPackage conduit_pkg = BluebeamConduitPackage.PackageFromPolyLines(poly_lines);
            var hangers = BluebeamSingleHanger.SingleHangersFromConduitPackage(conduit_pkg, threaded_rod_size);

            // parse all groups into fire alarm boxes with connectors package
            BlubeamFireAlarmBoxPackage box_package = BlubeamFireAlarmBoxPackage.BoxPackageFromAnnotations(f.AllAnnotations);

            // Output to Excel
            var excel_out_path = exe_path + Path.GetFileNameWithoutExtension(pdf_input_path) + ".xlsx";

            if (FileIsLocked(excel_out_path, FileAccess.Read))
            {
                Process[] workers = Process.GetProcessesByName("excel");
                foreach (Process worker in workers)
                {
                    if (worker.MainWindowTitle.Contains(excel_out_path))
                    {
                        Console.WriteLine("Killing open excel process -> " + worker.MainWindowTitle);
                        worker.Kill();
                        worker.WaitForExit();
                        worker.Dispose();
                    }
                }
            }

            if (File.Exists(excel_out_path))
                File.Delete(excel_out_path);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelEngine exporter = null;

            try
            {
                exporter = new ExcelEngine(excel_out_path);

                if (!ExcelEngine.PrepExportFile(excel_out_path))
                {
                    Console.WriteLine("Failed to generate Excel output.");
                    Console.ReadKey();
                    return;
                }

                ExcelOutputSheet s1 = new ExcelOutputSheet(ExportStyle.FireAlarm);
                Console.WriteLine("\nGenerating Bill of Material:");
                Console.WriteLine("\tCreated sheet");
                exporter.RegisterSheets("Fire Alarm", s1);
                Console.WriteLine("\tRegistered sheet");

                var labor_path = exe_path + "labor\\";
                labor_path = labor_path + @"labor_entries.json";

                s1.GenerateFireAlarmSheet(labor_path, "<Project Title Goes Here>", conduit_pkg, box_package, hangers, hanger_spacing);
                Console.WriteLine("\tFilled Sheet");
                exporter.Close();
                Console.WriteLine("\tBill of Material Generated!");
                exporter.OpenExcel();
                // exporter.OpenPDF(PdfCopyFilename);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel Engine failure: " + ex.Message.ToString());
                Console.ReadKey();
            }

            Console.WriteLine("\nBill of Material: ");
            Console.WriteLine("\tFilename - " + excel_out_path.Split("\\").Last());
            Console.WriteLine("\tFull Path - " + excel_out_path);

            Console.ReadKey();
        }

        /// <summary>
        /// Process teh P3 in wall from a bluebeam document
        /// </summary>
        private static void RunP3InWall(string exe_path)
        {
            // FDFImportPath = exe_path + @"p3export.fdf";
            // FDFOutputPath = exe_path + @"p3export_with_markups.fdf";

            var labor_path = exe_path + "labor\\labor_entries.json";
            var excel_output_path = exe_path;

            var pdf_output_suffix = "_p3inwall_device_codes.pdf";
            var pdf_input_path = GetPdfInputFileName(exe_path, pdf_output_suffix);

            // check the input file path
            if (string.IsNullOrWhiteSpace(pdf_input_path) || !File.Exists(pdf_input_path))
            {
                Console.WriteLine("Failed to retrieve pdf file name");
                Console.ReadKey();
                return;
            }
            else
            {
                Console.WriteLine("Using pdf file: " + pdf_input_path);
            }

            // PrintAllAnnotationPropertiesToFile(exe_path, Path.GetFullPath(pdf_input_path));

            var pdf_output_path = exe_path + Path.GetFileNameWithoutExtension(pdf_input_path) + pdf_output_suffix;

            Console.WriteLine("Parsing file: " + Path.GetFileNameWithoutExtension(pdf_input_path));
            CleanupPdfOutputFiles(pdf_output_path);

            Pdforge f = new Pdforge(pdf_input_path, pdf_output_path);
            PdfCustomColumnCollection columns = new PdfCustomColumnCollection(f);

            // check the columns for expected
            if (!columns.HasExpectedColumnHeaders(BluebeamP3BoxCollection.ExpectedBluebeamColumns, out var missing_columns))
            {
                Console.WriteLine("\nThe following custom columns are not loaded in the PDF in Bluebeam:");
                missing_columns.ToList().ForEach(x => Console.WriteLine(x));
                Console.WriteLine("\nPlease load the columns into the selected PDF document using Bluebeam and then rerun this program");
                Console.ReadKey();
                return;
            }
            else
            {
                Console.WriteLine("All expected custom columns are loaded correctly");
            }

            var box_annots = f.GetPage(0).Annotations.Select(x => x as PdfAnnotation);
            BluebeamP3BoxCollection boxes = BluebeamP3BoxCollection.BoxPackageFromAnnotations(box_annots, columns);

            Console.WriteLine("\nShorthand Device Code Pairs:");
            Console.WriteLine(boxes.BSHD_Resolver.ToString());

            // Generate Pdf with device codes
            while (File.Exists(pdf_output_path) && FileIsLocked(pdf_output_path, FileAccess.ReadWrite))
            {
                Console.WriteLine("The pdf output file " + Path.GetFileName(pdf_output_path) + " is in use. Please close it in Bluebeam");
                Thread.Sleep(500);
            }

            if (File.Exists(pdf_output_path)) File.Delete(pdf_output_path);

            boxes.SaveMarkupPdf(pdf_input_path, pdf_output_path, f, columns);

            // @TODO: Open Pdf file after save
            //Process.Start(pdf_output_path);

            // Output to Excel
            var excel_out_path = exe_path + Path.GetFileNameWithoutExtension(pdf_input_path) + ".xlsx";

            if (FileIsLocked(excel_out_path, FileAccess.Read))
            {
                Process[] workers = Process.GetProcessesByName("excel");
                foreach (Process worker in workers)
                {
                    if (worker.MainWindowTitle.Contains(excel_out_path))
                    {
                        Console.WriteLine("Killing open excel process -> " + worker.MainWindowTitle);
                        worker.Kill();
                        worker.WaitForExit();
                        worker.Dispose();
                    }
                }
            }

            if (File.Exists(excel_out_path))
                File.Delete(excel_out_path);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelEngine exporter = null;

            try
            {
                exporter = new ExcelEngine(excel_out_path);

                if (!ExcelEngine.PrepExportFile(excel_out_path))
                {
                    Console.WriteLine("Failed to generate Excel output.");
                    Console.ReadKey();
                    return;
                }

                var parts = P3InWall.GetLegacyDevices(boxes.Boxes);

                // print bundle names
                Console.WriteLine("\nBundle Names:");
                var b_name_print = parts.Select(x => x.BundleName).Distinct().ToList();
                b_name_print.Sort();
                Console.WriteLine(string.Join("\n", b_name_print));

                ExcelOutputSheet s1 = new ExcelOutputSheet(ExportStyle.LegacyP3InWall);
                Console.WriteLine("\nGenerating Bill of Material:");
                Console.WriteLine("\tCreated sheet");
                exporter.RegisterSheets("P3 In Wall", s1);
                Console.WriteLine("\tRegistered sheet");

                s1.GenerateLegacyP3InWallSheet(labor_path, "<Project Title Goes Here>", parts);
                Console.WriteLine("\tFilled Sheet");
                exporter.Close();
                Console.WriteLine("\tBill of Material Generated!");
                exporter.OpenExcel();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel Engine failure: " + ex.Message.ToString());
                Console.ReadKey();
            }

            Console.WriteLine("\nBill of Material: ");
            Console.WriteLine("\tFilename - " + excel_out_path.Split("\\").Last());
            Console.WriteLine("\tFull Path - " + excel_out_path);

            Console.ReadKey();
        }

        public static void CleanupExcelFiles(params string[] excel_file_paths)
        {
            foreach (var p in excel_file_paths)
                if (File.Exists(p)) File.Delete(p);
        }

        public static void CleanupPdfOutputFiles(params string[] pdf_output_file_paths)
        {
            foreach (var p in pdf_output_file_paths)
                if (File.Exists(p)) File.Delete(p);
        }

        public static void KillExcelProcess(string excel_output_path)
        {
            if (FileIsLocked(excel_output_path, FileAccess.Read))
            {
                Process[] workers = Process.GetProcessesByName("excel");
                foreach (Process worker in workers)
                {
                    if (worker.MainWindowTitle.Contains(Path.GetFileNameWithoutExtension(excel_output_path)))
                    {
                        Console.WriteLine("Killing open excel process -> " + worker.MainWindowTitle);
                        worker.Kill();
                        worker.WaitForExit();
                        worker.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Get user input to select a pdf file name if multiple are present
        /// </summary>
        public static string GetPdfInputFileName(string exe_path, string pdf_suffix)
        {
            var raw_fns = Directory
                .GetFiles(exe_path)
                .Where(x =>
                    x.ToLower().EndsWith(".pdf") &&
                    !x.Contains(string.IsNullOrWhiteSpace(pdf_suffix) ? "STUPIDBYPASS" : pdf_suffix));

            if (!raw_fns.Any()) return String.Empty;

            string get_fn(string raw_fn) => raw_fn.Split('\\').Last();

            if (raw_fns.Count() == 1)
            {
                return get_fn(raw_fns.First());
            }
            else
            {
                List<string> fns = new List<string>();
                fns.AddRange(raw_fns.Select(x => get_fn(x)));

                for (var i = 0; i < fns.Count(); i++)
                {
                    Console.WriteLine(string.Format("{0}. {1}", i + 1, fns[i]));
                }

                Console.WriteLine();

                int response = -1;
                bool running = true;

                do
                {
                    Console.Write("Select the PDF file you would like to parse (by number): ");
                    var str_resp = Console.ReadLine();
                    bool s = int.TryParse(str_resp, out int result);

                    if (!s || ((result - 1) > (fns.Count() - 1) || (result - 1) < 0))
                    {
                        Console.WriteLine(
                            "Please provide a number in the range 1 - " +
                            fns.Count());
                    }
                    else
                    {
                        response = result - 1;
                        running = false;
                    }
                }
                while (running);


                return fns[response];
            }
        }

        public static bool FileIsLocked(string filename, FileAccess file_access)
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

        /* private static void ParsePdfMarkups(
            IEnumerable<P3BluebeamFDFMarkup> markups,
            string pdf_output_path, string excel_output_path,
            string labor_path)
        {
            KillExcelProcess(excel_output_path);
            CleanupExcelFiles(excel_output_path);

            ExcelEngine exporter = null;

            try
            {
                exporter = new ExcelEngine(excel_output_path);

                if (!ExcelEngine.PrepExportFile(excel_output_path))
                {
                    Console.WriteLine("Failed to generate Excel output.");
                    Console.ReadKey();
                    return;
                }

                var parts = P3InWall.GetLegacyDevices(markups);
                // Console.WriteLine(string.Join("\n", parts.Select(x => x.ToString())));
                ExcelOutputSheet s1 = new ExcelOutputSheet(ExportStyle.LegacyP3InWall);
                Console.WriteLine("Generating Bill of Material:");
                Console.WriteLine("\tCreated sheet");
                exporter.RegisterSheets("P3 in wall", s1);
                Console.WriteLine("\tRegistered sheet");
                s1.GenerateLegacyP3InWallSheet(labor_path, "<Project Title Goes Here>", parts);
                Console.WriteLine("\tFilled Sheet");
                exporter.Close();
                Console.WriteLine("\tBill of Material Generated!");
                exporter.OpenExcel();
                exporter.OpenPDF(pdf_output_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }

            Console.WriteLine("\nBill of Material: ");
            Console.WriteLine("\tFilename - " + excel_output_path.Split("\\").Last());
            Console.WriteLine("\tFull Path - " + excel_output_path);

            Console.WriteLine("\nFloorplan w/ Device Codes: ");
            Console.WriteLine("\tFilename - " + Path.GetFileName(pdf_output_path));
            Console.WriteLine("\tFull Path - " + pdf_output_path);
        } */

        /* public static string FDFImportPath { get; set; } = "";
        public static string FDFOutputPath { get; set; } = ""; */
        // public static string CSVImportPath { get; set; } = "";

        /* /// <summary>
        /// Error handler for watcher class
        /// </summary>
        private static void OnError(object sender, ErrorEventArgs e)
        {
            if (e.GetException().GetType() == typeof(InternalBufferOverflowException))
            {
                Console.WriteLine("Error: File System Watcher internal buffer overflow at " + DateTime.Now);
            }
            else
            {
                Console.WriteLine("Error: Watched directory not accessible at " + DateTime.Now);
            }
            Console.ReadKey();
        } */

        /* /// <summary>
        /// file system watcher changed handler that looks for fdf file in directory
        /// </summary>
        /// <param name="watch_path">the path the watch</param>
        private static void WatchForFdf(string watch_path)
        {
            //CSVImportPath = one_up_path + "\\p3export.csv";
            var input_gobble_path = watch_path;

            // delete any csv's already present
            CleanupFDFInputFiles();

            // watch for csv summary
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = input_gobble_path;
            watcher.EnableRaisingEvents = true;
            watcher.NotifyFilter = NotifyFilters.FileName;
            watcher.Filter = "p3export.fdf";
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.Error += new ErrorEventHandler(OnError);

            do
            {
                Console.WriteLine("Listening for .fdf input file at: " + FDFImportPath);
                watcher.WaitForChanged(WatcherChangeTypes.Created);
            } while (true);
        } */

        /// <summary>
        /// Parse .fdf bluebeam markup export file when detected in directory
        /// </summary>
        /* private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            CleanupFDFOutputFiles();

            Console.WriteLine("Found .fdf input! Killing open Excel and waiting for file lock");
            CleanupExcelFiles();
            KillExcelProcess();

            while (FileIsLocked(FDFImportPath, FileAccess.Read))
                Thread.Sleep(100);


            Console.WriteLine("File is ready to be parsed at: " + ExcelOutputFilePath + "\n");

            ExcelEngine exporter = null;
            BluebeamP3MarkupExport export = null;

            // parse bluebeam markup
            try
            {
                export = new BluebeamP3MarkupExport(FDFImportPath, FDFOutputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            try
            {
                exporter = new ExcelEngine(ExcelOutputFilePath);

                if (!ExcelEngine.PrepExportFile(ExcelOutputFilePath))
                {
                    Console.WriteLine("Failed to generate Excel output.");
                    Console.ReadKey();
                    return;
                }

                var parts = P3InWall.GetLegacyDevices(export);
                // Console.WriteLine(string.Join("\n", parts.Select(x => x.ToString())));
                ExcelOutputSheet s1 = new ExcelOutputSheet(ExportStyle.LegacyP3InWall);
                Console.WriteLine("Created sheet");
                exporter.RegisterSheets("P3 in wall", s1);
                Console.WriteLine("Registered sheet");
                s1.GenerateLegacyP3InWallSheet(LaborFilePath, "<Project Title Goes Here>", parts);
                Console.WriteLine("Generated Sheet");
                exporter.Close();
                exporter.OpenExcel();
                export.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                CleanupFDFInputFiles();
                Console.ReadKey();
            }

            CleanupFDFInputFiles();
            Console.WriteLine("Finished");
            Console.WriteLine("----------------------------------------------\n");
            Console.WriteLine("A file named 'p3export_with_markups.fdf' has been generated in the folder that you just exported markups too. Go to [Markups List -> Markups -> Import] and import 'p3export_with_markups.fdf' to place Device Codes on the markups in your file");
        } */

        /* public static void CleanupCSVFiles() {
            if(File.Exists(CSVImportPath))
                File.Delete(CSVImportPath);
        } */
    }
}

