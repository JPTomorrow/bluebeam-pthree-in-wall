using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Draw = System.Drawing;

namespace JPMorrow.Excel
{
    // represents an excel worksheet that can be injected into an BOMExporter ExcelInstance
    public partial class ExcelOutputSheet 
	{
		private static bool ShowRowCrawlGraph = true;
        private List<string> RowCrawlGraph = new List<string>();

        public ExcelWorksheet Sheet { get; private set; }
        public ExcelSheetStyle SheetStyle { get; private set; }
        public bool HasData { get; private set; } = false;

        public int Row { get; set; } = 0;
		public int R { get => Row; }
        

        // Increment the current row
        public int NextRow(int cnt)
		{
			if(ShowRowCrawlGraph) RowCrawlGraph.Add(Row.ToString());
            Row += cnt;
            return Row;
        } 

		public string PrintRowCrawlGraph()
		{
			if(!RowCrawlGraph.Any())
                return "No Row Crawl Graph Present.";

            return string.Join("\n", RowCrawlGraph);
        }

        public ExcelOutputSheet(ExcelSheetStyle style) 
		{
            SheetStyle = style;
        }

		public void SetSheet(ExcelEngine exporter, string title_prefix) 
		{
			Sheet = exporter.ExcelInstance.Workbook.Worksheets.Add(title_prefix + SheetStyle.SheetTitleSuffix);
		}

        // format the excel worksheet for print
        public void FormatExcelSheet(decimal margins) 
		{
            ApplyBorderToSheet();
            Sheet.PrinterSettings.Orientation = eOrientation.Portrait;
            Sheet.PrinterSettings.HorizontalCentered = true;
            Sheet.PrinterSettings.FitToPage = true;
            Sheet.PrinterSettings.FitToWidth = 1;
            Sheet.PrinterSettings.FitToHeight = 0;
            SetRepeatRow();
            SetMargins(margins);
            Sheet.Cells.AutoFitColumns();
        }

        /// <summary>
		/// Insert a full row of content into the excel sheet
		/// </summary>
		public void InsertIntoRow(params object[] values) 
		{
			Stack val_stack = new Stack(values.Reverse().ToArray());

			for(char c = SheetStyle.ColumnExtents[0]; c <= SheetStyle.ColumnExtents[1]; c++)
			{
				if(val_stack.Count > 0)
					this[c, Row].Value = val_stack.Pop();
			}
		}

        /// <summary>
		/// Insert a divider of a certain color into the excel file
		/// </summary>
		public void InsertSingleDivider(
			Draw.Color bg_color, Draw.Color font_color,
			string header_text = null, double height = 20) 
		{
			ExcelRange rng = this[SheetStyle.ColumnExtents[0], SheetStyle.ColumnExtents[1], Row, Row];
			rng.Merge = true;
			rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
			rng.Style.Fill.BackgroundColor.SetColor(bg_color);
			rng.Style.Font.Color.SetColor(font_color);
			rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
			rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			Sheet.Row(Row).Height = height;

			if (string.IsNullOrWhiteSpace(header_text))
			{
				NextRow(1);
				return;
			}

			rng.Value = header_text;
			NextRow(1);
		}

        /// <summary>
		/// Insert a header title on the specified Excel Worksheet
		/// </summary>
		public void InsertHeader(string title, string subtitle, string postfix) 
		{
			Row = 1;

			//make main headers
			this[1,Row].Merge = true;
			this[1,Row].Style.WrapText = true;
			var rich_txt = this[1,Row].RichText.Add(title);
			rich_txt.Size = 24f;

			Row = 2;

			this[2, Row].Merge = true;
			this[2, Row].Value = subtitle;

			this[1, Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			this[1, Row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			this[1, Row].Style.Fill.PatternType = ExcelFillStyle.Solid;
			this[1, Row].Style.Fill.BackgroundColor.SetColor(Draw.Color.Maroon);
			rich_txt.Color = Draw.Color.White;
			this[2, Row].Style.Font.Color.SetColor(Draw.Color.White);

			this[2, Row].Style.Font.Size = 14f;

			Sheet.Row(1).Height = 60;
			Sheet.Row(2).Height = 40;

			//make column headers
			Row = 3;
			InsertIntoRow(SheetStyle.ColumnHeaders);

			this[3, Row].AutoFitColumns();
			this[3, Row].Style.Fill.PatternType = ExcelFillStyle.Solid;
			this[3, Row].Style.Border.BorderAround(ExcelBorderStyle.Medium);
			this[3, Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			this[3, Row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			this[3, Row].Style.Fill.BackgroundColor.SetColor(Draw.Color.Orange);
			this[3, Row].Style.Font.Color.SetColor(Draw.Color.Black);
			Sheet.Row(Row).Height = 15;
			HideColumnsToEnd();
			Row = 4;
		}

        /// <summary>
		/// Insert a grand total on the specified Excel Worksheet
		/// </summary>
        public void InsertGrandTotal(
			string gt_header, ref double gt, bool extra_space,
            bool skip_first_row_inc = false, bool reset_total = false )
		{

			if(skip_first_row_inc) NextRow(1);

			this[SheetStyle.ColumnExtents[1] + R.ToString()].Value =  gt_header + ": " + Math.Ceiling(gt);
			this[SheetStyle.ColumnExtents[1] + R.ToString()].Style.Border.BorderAround(ExcelBorderStyle.Thick);
			if(reset_total)
				gt = 0.0;

			NextRow(1);
			if(extra_space) NextRow(1);
		}

		public void MakeFooter() 
		{
			var dt = DateTime.Now;
			Sheet.HeaderFooter.AlignWithMargins = true;
			Sheet.HeaderFooter.OddFooter.RightAlignedText = string.Format(
				"Created {0}/{1}/{2} at {3}:{4}{5}", 
				dt.Month, dt.Day, dt.Year, dt.Hour, dt.Minute, 
				dt.ToString("tt", CultureInfo.InvariantCulture));
		}

        /// <summary>
		/// Hide extranious columns in a worksheet
		/// </summary>
		public void HideColumnsToEnd() 
		{
			var start = ColLetterToInt((char)(SheetStyle.ColumnExtents[1] + 2));
			var end = Sheet.Cells.End.Column + 1;

			for (var i = start; i < end; i++)
				Sheet.Column(i).Hidden = true;
		}

        /// <summary>
		/// Merge cells in the current sheet and assign a value to them
		/// </summary>
        public void MergeCells(
            char start_col, char end_col, int start_row, 
            int end_row, string cell_value, 
            string fill_txt = null, int font_size = 18) 
		{
            if(start_row >= end_row) return;

			if(fill_txt == null)
				fill_txt = "---";

			var rng = this[start_col, end_col, start_row, end_row];
			rng.Merge = true;
			rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			rng.Style.Font.Size = font_size;

            this[start_col, end_col, start_row, end_row].Value = "---";
		}

        /// <summary>
		/// Merge cells in the current sheet and assign a value to them
		/// </summary> 
        public void ApplyBorderToSheet() 
		{
			int last_row = Sheet.Dimension.End.Row;
			foreach(var  cell in FullRngClamped(3, last_row))
				cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);
		}

        /// <summary>
		/// Return a clampled range between the two rows specified.
		/// </summary>
		public ExcelRange FullRngClamped(params int[] rows) 
		{
			if(rows.Count() != 2)
				throw new ArgumentException("Need to provide exactly 2 row numbers.");

			return	Sheet.Cells[(SheetStyle.ColumnExtents[0] + ":" + SheetStyle.ColumnExtents[1]).Insert(3, rows[1].ToString()).Insert(1, rows[0].ToString())];
		}

        /// <summary>
		/// Change the width of a column
		/// </summary>
        public void ChangeColumnWidth(char col, double width) 
		{

			int col_int = ColLetterToInt(col);
			Sheet.Column(col_int).Width = width;
		}

        /// <summary>
		/// Change the alignment of a column
		/// </summary>
        public void ChangeColumnAlignment(int start_row, char[] col_range, ExcelHorizontalAlignment align)
		{
			if(col_range.Length != 2)
				throw new ArgumentException("Requires two values to be entered for columns.");

			Sheet.Cells[col_range[0] + start_row.ToString() + ":" +  col_range[1] + LastRow().ToString()]
				.Style.HorizontalAlignment = align;
		}

        /// <summary>
		/// Convert a character to an integer version of itself
		/// </summary>
		private int ColLetterToInt(char letter) 
		{
			char parseLetter = Char.ToUpper(letter);
			string key = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			if (!key.Contains(parseLetter)) return -1; //failure
			return key.IndexOf(parseLetter) + 1;
		}

        /// <summary>
		/// Apply color formatting to a cell
		/// </summary>
		public void ApplyColorToColumn(char col, string color) 
		{
			col = char.ToUpper(col);

			SystemColorInfo color_info;
			try 
			{
				color_info = system_color_swap[color];
			}
			catch 
			{
				color_info = system_color_swap["Black"];
			}

			var rng = this[col + Row.ToString()];

			rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
			rng.Style.Fill.BackgroundColor.SetColor(color_info.Background_Color);
			rng.Style.Border.BorderAround(ExcelBorderStyle.Thin);

			rng.Style.Border.Top.Color.SetColor(color_info.Border_Color);
			rng.Style.Border.Bottom.Color.SetColor(color_info.Border_Color);
			rng.Style.Border.Left.Color.SetColor(color_info.Border_Color);
			rng.Style.Border.Right.Color.SetColor(color_info.Border_Color);

			rng.Style.Font.Color.SetColor(color_info.Font_Color);
		}

        public void SetMargins(decimal margin) 
		{
			Sheet.PrinterSettings.TopMargin = margin;
			Sheet.PrinterSettings.BottomMargin = 0.5m;
			Sheet.PrinterSettings.LeftMargin = margin;
			Sheet.PrinterSettings.RightMargin = margin;
		}

        public void SetRepeatRow() => Sheet.PrinterSettings.RepeatRows = new ExcelAddress(String.Format("'{1}'!${0}:${0}", 3, Sheet.Name));
        private int LastRow() => Sheet.Cells.Where(cell => !string.IsNullOrEmpty(cell.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Row;

        #region indexers
		public ExcelRange this[int row_1, int row_2] {
			get {
				string rng  = SheetStyle.ColumnExtents[0] + row_1.ToString() + ":" + SheetStyle.ColumnExtents[1] + row_2.ToString();
				return Sheet.Cells[rng];
			}
		}

		public ExcelRange this[char col_1, char col_2, int row_1, int row_2] {
			get {
				string rng  = col_1 + row_1.ToString() + ":" + col_2 + row_2.ToString();
				return Sheet.Cells[rng];
			}
		}

		public ExcelRange this[string rng] {
			get {
				if(rng.StartsWith("row:")) {
					string fixed_rng = rng.Remove(0, 4);
					fixed_rng = Regex.Replace(fixed_rng, @"\s+", "");
					fixed_rng = SheetStyle.ColumnExtents[0] + fixed_rng + ":" + SheetStyle.ColumnExtents[1] + fixed_rng;
					return Sheet.Cells[fixed_rng];
				}

				return Sheet.Cells[rng];
			}
		}

		public ExcelRange this[char col] {
			get => Sheet.Cells[col.ToString() + ":" + col.ToString()];
		}

		public ExcelRange this[char col1, char col2] {
			get => Sheet.Cells[col1.ToString() + ":" + col2.ToString()];
		}

		public ExcelRange this[char col, int row] {
			get => Sheet.Cells[col.ToString() + row.ToString() + ":" + col.ToString() + row.ToString()];
		}
		#endregion
    }
}