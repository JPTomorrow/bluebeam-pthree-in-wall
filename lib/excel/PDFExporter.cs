using System;
using ExcelInst = Microsoft.Office.Interop.Excel;

namespace JPMorrow.Tools.PDF
{
    public static class PDFExport
	{
		/// <summary>
        /// Print Excel file to PDF
        /// </summary>
        /// <param name="excelLocation">The location of the Excel file</param>
        /// <param name="outputLocation">The output location of the PDF file</param>
		public static bool ExceltoPdf(string excelLocation, string outputLocation)
		{
			ExcelInst.Application app = new ExcelInst.Application();
			app.Visible = false;
			ExcelInst.Workbook wkb = app.Workbooks.Open(excelLocation);
            bool created = false;
            try
			{
				wkb.ExportAsFixedFormat(ExcelInst.XlFixedFormatType.xlTypePDF, outputLocation);
                created = true;
            }
			catch
			{
                Console.WriteLine(
					"The PDF export file is currently open and cannot be exported." + 
					" Please close the file and re-export");
                created = false;
            }
			finally
			{
				wkb.Close();
				app.Quit();
			}

            return created;
        }
	}
}