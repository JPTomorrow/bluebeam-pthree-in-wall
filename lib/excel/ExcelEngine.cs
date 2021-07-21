using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace JPMorrow.Excel
{
    public class ExcelEngine {
        public ExcelPackage ExcelInstance { get; private set; }
        public string FileName { get; set; }

        private string CleanFileName { get {
            string clean_file_name = System.IO.Path.GetFileNameWithoutExtension(FileName);
			clean_file_name = clean_file_name.Replace("_", " ");
			clean_file_name = clean_file_name.Split('.').First();
			return clean_file_name;
        }}

        public ExcelEngine(string file_name) 
		{
            FileName = file_name;
            ExcelInstance = new ExcelPackage(new FileInfo(file_name));
        }

        // add a sheet to the sheets for this ExcelInstance
        public void RegisterSheets(string sheet_name_prefix, params ExcelOutputSheet[] sheets) 
		{
			try
			{
				foreach(var s in sheets) 
				{
					s.SetSheet(this, sheet_name_prefix);
				}
			}
			catch
			{
                throw new Exception("sheet is corrupted");
            }
        }

		public void Close() 
		{
			ExcelInstance.Save();
			ExcelInstance.Dispose();
		}

		/// <summary>
        /// Opens an excel file of the current export
        /// </summary>
        /// <returns>true if the file opened, false if it failed, or was already open</returns>
        public void OpenExcel()
		{
            var fn = FileName;
            try 
            {
                var p = new Process();
                ProcessStartInfo info = new ProcessStartInfo(fn);
                info.UseShellExecute = true;

                p.StartInfo = info;
                p.Start();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Opens an pdf file of the current export
        /// </summary>
        /// <returns>true if the file opened, false if it failed, or was already open</returns>
        public bool OpenPDF(string pdf_filename)
		{
            try 
            {
                Process.Start(pdf_filename);
                return true;
            }
            catch
            {
                return false;
            }
        }

		#region file prep
        // perform all necessary checks on an export file.
        public static bool PrepExportFile(string filename) 
		{
            if(!File.Exists(filename)) 
			{
                var file = File.Create(filename);
                file.Close();
            }

            if(IsFileLocked(new FileInfo(filename))) 
			{
                Console.WriteLine("Excel file must be closed in order to export.");
                return false;
            }

            if(File.Exists(filename)) File.Delete(filename);
            return true;
        }

		// check to see if the file is locked
        private static bool IsFileLocked(FileInfo file) 
		{
            FileStream stream = null;

            try 
			{
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException) 
			{
                return true;
            }
            finally 
			{
                if (stream != null) stream.Close();
            }

            //file is not locked
            return false;
        }
        #endregion
    }
}

/* namespace JPMorrow.Test
{
    using JPMorrow.Excel;

    public static partial class TestBed
    {
        public static TestResult TestExcelEngine(string settings_path, Document doc, UIDocument uidoc)
        {
            var test_path = settings_path + "test/" + "TestExcelWorkbook.xlsx";
            Console.WriteLine(settings_path);
            ExcelEngine e = new ExcelEngine(test_path);

            ExcelOutputSheet s1 = new ExcelOutputSheet(ExportStyle.WirePull);
            e.RegisterSheets("Test", s1);
            

            return new TestResult("Excel Engine", true);
        }
    }
} */