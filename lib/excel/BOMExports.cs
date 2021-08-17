///
/// This file is a list of all the preconfigured exports as extention functions that can be called by a BOMOutputSheet
/// Author: Justin Morrow
///


using System;
using System.Collections.Generic;
using System.Linq;
using JPMorrow.P3;
using JPMorrow.Pdf.Bluebeam.FireAlarm;
using JPMorrow.Revit.Labor;
using OfficeOpenXml.Style;
using Draw = System.Drawing;

namespace JPMorrow.Excel
{
    public partial class ExcelOutputSheet {

        // dictionary for apply color lookup
		private static readonly Dictionary<string, SystemColorInfo> system_color_swap = new Dictionary<string, SystemColorInfo>() {
			{"None", new SystemColorInfo() 						{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.DimGray } },
			{"Black", new SystemColorInfo() 					{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.Black, Border_Color=Draw.Color.Black } },
			{"Red", new SystemColorInfo() 						{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.Red, Border_Color=Draw.Color.Red } },
			{"Blue", new SystemColorInfo() 						{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.Blue, Border_Color=Draw.Color.Blue } },

			{"White w/ Black Stripe", new SystemColorInfo() 	{ Font_Color=Draw.Color.DimGray, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Black } },
			{"White w/ Red Stripe", new SystemColorInfo() 		{ Font_Color=Draw.Color.Red, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Red } },
			{"White w/ Blue Stripe", new SystemColorInfo() 		{ Font_Color=Draw.Color.Blue, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Blue } },
			{"White w/ Orange Stripe", new SystemColorInfo() 	{ Font_Color=Draw.Color.Orange, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Orange } },

			{"Brown", new SystemColorInfo() 					{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.Brown, Border_Color=Draw.Color.Brown } },
			{"Orange", new SystemColorInfo() 					{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.Orange, Border_Color=Draw.Color.Orange } },
			{"Yellow", new SystemColorInfo() 					{ Font_Color=Draw.Color.DimGray, Background_Color=Draw.Color.Yellow, Border_Color=Draw.Color.Yellow } },

			{"Gray w/ Brown Stripe", new SystemColorInfo() 		{ Font_Color=Draw.Color.SandyBrown, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.SaddleBrown } },
			{"Gray w/ Orange Stripe", new SystemColorInfo() 	{ Font_Color=Draw.Color.Orange, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.Orange } },
			{"Gray w/ Yellow Stripe", new SystemColorInfo() 	{ Font_Color=Draw.Color.Yellow, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.Yellow } },

			{"White", new SystemColorInfo() 					{ Font_Color=Draw.Color.DimGray, Background_Color=Draw.Color.White, Border_Color=Draw.Color.White } },
			{"Gray", new SystemColorInfo() 						{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.DimGray } },
			{"Green", new SystemColorInfo() 					{ Font_Color=Draw.Color.White, Background_Color=Draw.Color.Green, Border_Color=Draw.Color.Green } },

			{"Green w/ Yellow Stripe", new SystemColorInfo() 	{ Font_Color=Draw.Color.Yellow, Background_Color=Draw.Color.Green, Border_Color=Draw.Color.Yellow } },
		};

		//Struct to hold info about a system color
		internal struct SystemColorInfo {
			public Draw.Color Font_Color {get; set;}
			public Draw.Color Background_Color {get; set;}
			public Draw.Color Border_Color {get; set;}
		}

        /// <summary>
		/// Export a Legacy P3 In Wall sheet
		/// </summary>
        public void GenerateLegacyP3InWallSheet(string labor_import_path, string project_title, IEnumerable<P3PartCollection> colls) 
		{
            if(HasData) throw new Exception("The sheet already has data");
            string title = "M.P.A.C.T. - P3 In Wall";
            InsertHeader(title, project_title, "");

            var code_one_gt = 0.0;
			var code_one_sub = 0.0;
            static double shave_labor(double labor) => labor * 0.82;

            //colls = colls.OrderBy(x => x.DeviceCode).ToList();
            var field_hardware = P3PartTotal.GetPartTotals(colls, P3PartCategory.Hardware, P3PartCategory.Clip);
            var per_box_items = P3PartCollection.GetPartTotalsByBundleThenCategory(colls, P3PartCategory.Box, P3PartCategory.Bracket, P3PartCategory.Plaster_Ring, P3PartCategory.Stinger, P3PartCategory.Connector);
            var item_total = P3PartTotal.GetPartTotals(colls, P3PartCategory.Box, P3PartCategory.Bracket, P3PartCategory.Plaster_Ring, P3PartCategory.Stinger, P3PartCategory.Connector);

            var entries = LaborExchange.LoadLaborFromFile(labor_import_path);
            var l = new LaborExchange(entries);

			foreach(var kvp in per_box_items)
			{
                var bundle_name = kvp.Key.Equals(string.Empty) ? "UNSET" : kvp.Key;
                InsertSingleDivider(Draw.Color.OrangeRed, Draw.Color.White, "Bundle Name: " + bundle_name);

				foreach(var t in kvp.Value)
				{
					var code = t.DeviceCode;
					InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, code);

					foreach(var p in t.Parts)
					{
						var has_item = l.GetItem(out var li, (double)p.Qty, p.Name);
						if(!has_item) throw new Exception("No Labor item for: " + p.Name);
						InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
						code_one_sub += li.TotalLaborValue; NextRow(1);
					}
					
					code_one_sub = Math.Ceiling(code_one_sub);
					code_one_gt += code_one_sub;
					InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
					code_one_sub = 0.0;
				}
			}

            

			InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Fixture Item Totals");

            foreach(var t in item_total.Parts)
            {
                var has_item = l.GetItem(out var li, (double)t.Qty, t.Name);
                if(!has_item) throw new Exception("No Labor item for: " + t.Name);
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                NextRow(1);
            }

            NextRow(1);
			InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Field Labor Hardware Items");

			foreach(var part in field_hardware.Parts)
			{
                var has_item = l.GetItem(out var li, (double)part.Qty, part.Name);
                if(!has_item) throw new Exception("No Labor item for: " + part.Name);
				InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
				code_one_sub += li.TotalLaborValue; NextRow(1);
			}

			code_one_gt += code_one_sub;
			code_one_sub = Math.Ceiling(code_one_sub);
			InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
			code_one_sub = 0.0;

			InsertGrandTotal("Code 01 | Empty Raceway | Grand Total", ref code_one_gt, false, false, false);
			code_one_gt = shave_labor(code_one_gt);
			InsertGrandTotal("Code 01 w/ 0.82 Labor Factor", ref code_one_gt, true, false, true);

            FormatExcelSheet(0.1M);
            MakeFooter();

            // debugger.show(err:PrintRowCrawlGraph());
            HasData = true;
        }

		public void GenerateFireAlarmSheet(
			string labor_import_path, string project_title, 
			BluebeamConduitPackage conduit_pkg, BlubeamFireAlarmBoxPackage box_pkg) 
		{
			if(HasData) throw new Exception("The sheet already has data");
            string title = "M.P.A.C.T. - Fire Alarm";
            InsertHeader(title, project_title, "");

            var code_one_gt = 0.0;
			var code_one_sub = 0.0;
            static double shave_labor(double labor) => labor * 0.82;

			var entries = LaborExchange.LoadLaborFromFile(labor_import_path);
            var l = new LaborExchange(entries);

			// boxes

            // conduit
            var emt_total_1 = conduit_pkg.GetTotalEmtLengthRounded("1/2\"");
            var emt_total_2 = conduit_pkg.GetTotalEmtLengthRounded("3/4\"");
            var emt_total_3 = conduit_pkg.GetTotalEmtLengthRounded("1\"");
            var emt_total_4 = conduit_pkg.GetTotalEmtLengthRounded("1 1/4\"");
            var emt_total_5 = conduit_pkg.GetTotalEmtLengthRounded("1 1/2\"");
            var emt_total_6 = conduit_pkg.GetTotalEmtLengthRounded("2\"");

			var pvc_total_1 = conduit_pkg.GetTotalPvcLengthRounded("1/2\"");
            var pvc_total_2 = conduit_pkg.GetTotalPvcLengthRounded("3/4\"");
            var pvc_total_3 = conduit_pkg.GetTotalPvcLengthRounded("1\"");
            var pvc_total_4 = conduit_pkg.GetTotalPvcLengthRounded("1 1/4\"");
            var pvc_total_5 = conduit_pkg.GetTotalPvcLengthRounded("1 1/2\"");
            var pvc_total_6 = conduit_pkg.GetTotalPvcLengthRounded("2\"");

			var mc_total_1 = conduit_pkg.GetTotalMcCableLengthRounded("1/2\"");
            var mc_total_2 = conduit_pkg.GetTotalMcCableLengthRounded("3/4\"");
            var mc_total_3 = conduit_pkg.GetTotalMcCableLengthRounded("1\"");
            var mc_total_4 = conduit_pkg.GetTotalMcCableLengthRounded("1 1/4\"");
            var mc_total_5 = conduit_pkg.GetTotalMcCableLengthRounded("1 1/2\"");
            var mc_total_6 = conduit_pkg.GetTotalMcCableLengthRounded("2\"");

			void print_conduit(int qty, string labor_str)
			{
				if(qty.Equals("0' 0\"")) return;
				var has_item = l.GetItem(out var li, qty, labor_str);
				if(!has_item) throw new Exception("No Labor item for conduit");
				InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
				code_one_sub += li.TotalLaborValue; NextRow(1);
			}

			InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Conduit");

            print_conduit(emt_total_1, "Conduit - EMT - 1/2\"");
            print_conduit(emt_total_2, "Conduit - EMT - 3/4\"");
            print_conduit(emt_total_3, "Conduit - EMT - 1\"");
            print_conduit(emt_total_4, "Conduit - EMT - 1 1/4\"");
            print_conduit(emt_total_5, "Conduit - EMT - 1 1/2\"");
            print_conduit(emt_total_6, "Conduit - EMT - 2\"");
 
			print_conduit(pvc_total_1, "Conduit - PVC - 1/2\"");
            print_conduit(pvc_total_2, "Conduit - PVC - 3/4\"");
            print_conduit(pvc_total_3, "Conduit - PVC - 1\"");
            print_conduit(pvc_total_4, "Conduit - PVC - 1 1/4\"");
            print_conduit(pvc_total_5, "Conduit - PVC - 1 1/2\"");
            print_conduit(pvc_total_6, "Conduit - PVC - 2\"");

			//@TODO:MC Cable
			/* print_conduit(mc_total_1, "1/2\"", 		"Conduit - EMT - 1/2\"");
            print_conduit(mc_total_2, "3/4\"", 		"Conduit - EMT - 3/4\"");
            print_conduit(mc_total_3, "1\"", 		"Conduit - EMT - 1\"");
            print_conduit(mc_total_4, "1 1/4\"", 	"Conduit - EMT - 1 1/4\"");
            print_conduit(mc_total_5, "1 1/2\"", 	"Conduit - EMT - 1 1/2\"");
            print_conduit(mc_total_6, "2\"", 		"Conduit - EMT - 2\""); */

            code_one_sub = Math.Ceiling(code_one_sub);
			code_one_gt += code_one_sub;
			InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
			code_one_sub = 0.0;

			// couplings

			// connectors

			// hangers

			InsertGrandTotal("Code 01 | Empty Raceway | Grand Total", ref code_one_gt, false, false, false);
			code_one_gt = shave_labor(code_one_gt);
			InsertGrandTotal("Code 01 w/ 0.82 Labor Factor", ref code_one_gt, true, false, true);

            FormatExcelSheet(0.1M);
            MakeFooter();

            // debugger.show(err:PrintRowCrawlGraph());
            HasData = true;
        }
    }
}