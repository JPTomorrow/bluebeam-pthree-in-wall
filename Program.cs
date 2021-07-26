using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using JPMorrow.Bluebeam.Markup;
using JPMorrow.Excel;
using JPMorrow.P3;
using JPMorrow.PDF;

namespace BluebeamP3InWall
{
    class Program
    {
        public static string ExcelOutputPath { get; set; } = "";
        public static string ExcelOutputFileExt { get; set; } =  "_P3_In_Wall_BOM.xlsx";
        public static string LaborFilePath { get; set; } = "";

        public static string ExcelOutputFilePath { get {
                return ExcelOutputPath + PdfImportFilename.Split('.').First() + ExcelOutputFileExt;
            }}

        static void Main(string[] args)
        {
            string exe_path = GetThisExecutablePath();
            bool got_pdf_fn = ResolvePdfInputFilePath(exe_path);

            if(!got_pdf_fn) 
            {
                Console.WriteLine("Failed to retrieve pdf file name");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Parsing file: " + PdfImportFilename);
            GenerateFilePaths(exe_path);
            CleanupPdfOutputFiles();
            ParsePdf();
            Console.ReadKey();
            return;
        }

        public static string GetThisExecutablePath()
        {
            var split_path = System.Diagnostics.Process
                .GetCurrentProcess().MainModule.FileName
                .Split("\\").ToList();

            split_path.Remove(split_path.Last()); 
            return string.Join("\\", split_path) + "\\";
        }

        /// <summary>
        /// Generate all of the paths that the program will use
        /// </summary>
        public static void GenerateFilePaths(string exe_path)
        {
            var labor_path = exe_path + "labor\\";
            LaborFilePath = labor_path + @"labor_entries.json";
            ExcelOutputPath = exe_path;
            FDFImportPath = exe_path + @"p3export.fdf";
            FDFOutputPath = exe_path + @"p3export_with_markups.fdf";
            PdfImportPath = exe_path; 
            PdfCopyPath = exe_path;
        }

        public static void CleanupFDFInputFiles() 
        {
            if(File.Exists(FDFImportPath))
                File.Delete(FDFImportPath);
        }

        public static void CleanupFDFOutputFiles ()
        {
            if(File.Exists(FDFOutputPath))
                File.Delete(FDFOutputPath);
        }

        public static void CleanupExcelFiles() 
        {
            if(File.Exists(ExcelOutputFilePath))
                File.Delete(ExcelOutputFilePath);
        }

        public static void CleanupPdfOutputFiles()
        {
            if(File.Exists(PdfCopyPath + PdfCopyFilename))
                File.Delete(PdfCopyPath + PdfCopyFilename);
        }

        /// <summary>
        /// 
        /// </summary>
        public static void KillExcelProcess()
        {
            if(FileIsLocked(ExcelOutputFilePath, FileAccess.Read))
            {
                Process[] workers = Process.GetProcessesByName("excel");
                foreach (Process worker in workers)
                {
                    if(worker.MainWindowTitle.Contains(ExcelOutputFileExt))
                    {
                        Console.WriteLine("Killing open excel process -> " + worker.MainWindowTitle);
                        worker.Kill();
                        worker.WaitForExit();
                        worker.Dispose();
                    }
                }
            }
        }

        /// 
        /// Pdf METHOD
        /// 

        public static string PdfImportPath { get; set; } = "";
        public static string PdfCopyPath { get; set; } = "";

        public static string PdfImportFilename { get; set; } = "";
        public static string PdfCopyFilenameExt { get; set; } = "_Floorplan_With_Device_Codes";

        public static string PdfCopyFilename { get {
                var split = PdfImportFilename.Split('.');
                if(split.Length != 2) throw new Exception("PdfImportFilename is incorrect value");
                return split[0] + PdfCopyFilenameExt + "." + split[1];
            } }

        /// <summary>
        /// Get user input to select a pdf file name if multiple are present
        /// </summary>
        /// <returns>true if path gathered successfully, else false</returns>
        public static bool ResolvePdfInputFilePath(string exe_path)
        {
            var raw_fns = Directory
                .GetFiles(exe_path)
                .Where(x => x.ToLower().EndsWith(".pdf") && !x.Contains(PdfCopyFilenameExt));

            if(!raw_fns.Any()) return false;

            string get_fn(string raw_fn) => raw_fn.Split('\\').Last();

            if (raw_fns.Count() == 1)
            {
                PdfImportFilename = get_fn(raw_fns.First());
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

                    if(!s || ((result - 1) > (fns.Count() - 1) || (result - 1) < 0))
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


                PdfImportFilename = fns[response];
            }

            return true;
        }

        /// <summary>
        /// Parse a pdf bluebeam document for p3 in wall
        /// </summary>
        private static void ParsePdf()
        {
            try
            {
                var markups = PdfManager.EditBluebeamMarkups(PdfImportPath + PdfImportFilename);
                PdfManager.CreateCopyDocument(PdfImportPath + PdfImportFilename, PdfCopyPath + PdfCopyFilename, markups);
                ParsePdfMarkups(markups);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }
        }

        private static void ParsePdfMarkups(IEnumerable<P3BluebeamFDFMarkup> markups) 
        {
            KillExcelProcess();
            CleanupExcelFiles();
            
            ExcelEngine exporter = null;

            try
            {
                exporter = new ExcelEngine(ExcelOutputFilePath);

                /* var rows = P3CSVRow.ParseCSV(CSVImportPath);
                Console.WriteLine(P3CSVRow.PrintRaw(rows)); */

                if(!ExcelEngine.PrepExportFile(ExcelOutputFilePath)) 
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
                s1.GenerateLegacyP3InWallSheet(LaborFilePath, "<Project Title Goes Here>", parts);
                Console.WriteLine("\tFilled Sheet");
                exporter.Close();
                Console.WriteLine("\tBill of Material Generated!");
                exporter.OpenExcel();
                exporter.OpenPDF(PdfCopyFilename);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }

            Console.WriteLine("\nBill of Material: ");
            Console.WriteLine("\tFilename - " + ExcelOutputFilePath.Split("\\").Last());
            Console.WriteLine("\tFull Path - " + ExcelOutputFilePath);

            Console.WriteLine("\nFloorplan w/ Device Codes: ");
            Console.WriteLine("\tFilename - " + PdfCopyFilename);
            Console.WriteLine("\tFull Path - " + PdfCopyPath + PdfCopyFilename);
        }
        
        //
        // FDF METHOD
        //

        public static string FDFImportPath { get; set; } = "";
        public static string FDFOutputPath { get; set; } = "";
        // public static string CSVImportPath { get; set; } = "";

        /// <summary>
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
        }

        /// <summary>
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
        }

        /// <summary>
        /// Parse .fdf bluebeam markup export file when detected in directory
        /// </summary>
        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            CleanupFDFOutputFiles();

            Console.WriteLine("Found .fdf input! Killing open Excel and waiting for file lock");
            CleanupExcelFiles();
            KillExcelProcess();

            while(FileIsLocked(FDFImportPath, FileAccess.Read)) 
                Thread.Sleep(100);

            
            Console.WriteLine("File is ready to be parsed at: " + ExcelOutputFilePath + "\n");

            ExcelEngine exporter = null;
            BluebeamP3MarkupExport export = null;

            // parse bluebeam markup
            try
            {
                export = new BluebeamP3MarkupExport(FDFImportPath, FDFOutputPath);
                // Console.WriteLine(string.Join("\n", export.Markups.Select(x => x.DeviceCode)));
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            try
            {
                exporter = new ExcelEngine(ExcelOutputFilePath);

                /* var rows = P3CSVRow.ParseCSV(CSVImportPath);
                Console.WriteLine(P3CSVRow.PrintRaw(rows)); */

                if(!ExcelEngine.PrepExportFile(ExcelOutputFilePath)) 
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
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                CleanupFDFInputFiles();
                Console.ReadKey();
            }
            
            CleanupFDFInputFiles();
            Console.WriteLine("Finished");
            Console.WriteLine("----------------------------------------------\n");
            Console.WriteLine("A file named 'p3export_with_markups.fdf' has been generated in the folder that you just exported markups too. Go to [Markups List -> Markups -> Import] and import 'p3export_with_markups.fdf' to place Device Codes on the markups in your file");
        }

        /* public static void CleanupCSVFiles() {
            if(File.Exists(CSVImportPath))
                File.Delete(CSVImportPath);
        } */

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
    }
}

