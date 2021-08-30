///
/// This file is a list of all the preconfigured exports as extention functions that can be called by a BOMOutputSheet
/// Author: Justin Morrow
///


using System;
using System.Collections.Generic;
using System.Linq;
using JPMorrow.P3;
using JPMorrow.Pdf.Bluebeam;
using JPMorrow.Pdf.Bluebeam.FireAlarm;
using JPMorrow.Revit.Labor;
using OfficeOpenXml.Style;
using Draw = System.Drawing;

namespace JPMorrow.Excel
{
    public partial class ExcelOutputSheet
    {

        // dictionary for apply color lookup
        private static readonly Dictionary<string, SystemColorInfo> system_color_swap = new Dictionary<string, SystemColorInfo>() {
            {"None", new SystemColorInfo()                      { Font_Color=Draw.Color.White, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.DimGray } },
            {"Black", new SystemColorInfo()                     { Font_Color=Draw.Color.White, Background_Color=Draw.Color.Black, Border_Color=Draw.Color.Black } },
            {"Red", new SystemColorInfo()                       { Font_Color=Draw.Color.White, Background_Color=Draw.Color.Red, Border_Color=Draw.Color.Red } },
            {"Blue", new SystemColorInfo()                      { Font_Color=Draw.Color.White, Background_Color=Draw.Color.Blue, Border_Color=Draw.Color.Blue } },

            {"White w/ Black Stripe", new SystemColorInfo()     { Font_Color=Draw.Color.DimGray, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Black } },
            {"White w/ Red Stripe", new SystemColorInfo()       { Font_Color=Draw.Color.Red, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Red } },
            {"White w/ Blue Stripe", new SystemColorInfo()      { Font_Color=Draw.Color.Blue, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Blue } },
            {"White w/ Orange Stripe", new SystemColorInfo()    { Font_Color=Draw.Color.Orange, Background_Color=Draw.Color.White, Border_Color=Draw.Color.Orange } },

            {"Brown", new SystemColorInfo()                     { Font_Color=Draw.Color.White, Background_Color=Draw.Color.Brown, Border_Color=Draw.Color.Brown } },
            {"Orange", new SystemColorInfo()                    { Font_Color=Draw.Color.White, Background_Color=Draw.Color.Orange, Border_Color=Draw.Color.Orange } },
            {"Yellow", new SystemColorInfo()                    { Font_Color=Draw.Color.DimGray, Background_Color=Draw.Color.Yellow, Border_Color=Draw.Color.Yellow } },

            {"Gray w/ Brown Stripe", new SystemColorInfo()      { Font_Color=Draw.Color.SandyBrown, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.SaddleBrown } },
            {"Gray w/ Orange Stripe", new SystemColorInfo()     { Font_Color=Draw.Color.Orange, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.Orange } },
            {"Gray w/ Yellow Stripe", new SystemColorInfo()     { Font_Color=Draw.Color.Yellow, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.Yellow } },

            {"White", new SystemColorInfo()                     { Font_Color=Draw.Color.DimGray, Background_Color=Draw.Color.White, Border_Color=Draw.Color.White } },
            {"Gray", new SystemColorInfo()                      { Font_Color=Draw.Color.White, Background_Color=Draw.Color.DimGray, Border_Color=Draw.Color.DimGray } },
            {"Green", new SystemColorInfo()                     { Font_Color=Draw.Color.White, Background_Color=Draw.Color.Green, Border_Color=Draw.Color.Green } },

            {"Green w/ Yellow Stripe", new SystemColorInfo()    { Font_Color=Draw.Color.Yellow, Background_Color=Draw.Color.Green, Border_Color=Draw.Color.Yellow } },
        };

        //Struct to hold info about a system color
        internal struct SystemColorInfo
        {
            public Draw.Color Font_Color { get; set; }
            public Draw.Color Background_Color { get; set; }
            public Draw.Color Border_Color { get; set; }
        }

        /// <summary>
		/// Export a Legacy P3 In Wall sheet
		/// </summary>
        public void GenerateLegacyP3InWallSheet(string labor_import_path, string project_title, IEnumerable<P3PartCollection> colls)
        {
            if (HasData) throw new Exception("The sheet already has data");
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

            foreach (var kvp in per_box_items)
            {
                var bundle_name = kvp.Key.Equals(string.Empty) ? "UNSET" : kvp.Key;
                InsertSingleDivider(Draw.Color.OrangeRed, Draw.Color.White, "Bundle Name: " + bundle_name);

                foreach (var t in kvp.Value)
                {
                    var code = t.DeviceCode;
                    InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, code);

                    foreach (var p in t.Parts)
                    {
                        var has_item = l.GetItem(out var li, (double)p.Qty, p.Name);
                        if (!has_item) throw new Exception("No Labor item for: " + p.Name);
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

            foreach (var t in item_total.Parts)
            {
                var has_item = l.GetItem(out var li, (double)t.Qty, t.Name);
                if (!has_item) throw new Exception("No Labor item for: " + t.Name);
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                NextRow(1);
            }

            NextRow(1);
            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Field Labor Hardware Items");

            foreach (var part in field_hardware.Parts)
            {
                var has_item = l.GetItem(out var li, (double)part.Qty, part.Name);
                if (!has_item) throw new Exception("No Labor item for: " + part.Name);
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
            BluebeamConduitPackage conduit_pkg, BlubeamFireAlarmBoxPackage box_pkg,
            IEnumerable<BluebeamSingleHanger> hangers, double hanger_spacing)
        {
            if (HasData) throw new Exception("The sheet already has data");
            string title = "M.P.A.C.T. - Fire Alarm";
            InsertHeader(title, project_title, "");

            var code_one_gt = 0.0;
            var code_one_sub = 0.0;
            static double shave_labor(double labor) => labor * 0.82;

            var entries = LaborExchange.LoadLaborFromFile(labor_import_path);
            var l = new LaborExchange(entries);

            // boxes

            var d_boxes_1 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("D") && x.BoxSize.Equals("4\"")).Count();
            var i_boxes_1 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("I") && x.BoxSize.Equals("4\"")).Count();
            var x_boxes_1 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("X") && x.BoxSize.Equals("4\"")).Count();
            var t_boxes_1 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("T") && x.BoxSize.Equals("4\"")).Count();
            var xy_boxes_1 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("XY") && x.BoxSize.Equals("4\"")).Count();
            var y_boxes_1 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("Y") && x.BoxSize.Equals("4\"")).Count();

            var d_boxes_2 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("D") && x.BoxSize.Equals("4 11/16\"")).Count();
            var i_boxes_2 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("I") && x.BoxSize.Equals("4 11/16\"")).Count();
            var x_boxes_2 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("X") && x.BoxSize.Equals("4 11/16\"")).Count();
            var t_boxes_2 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("T") && x.BoxSize.Equals("4 11/16\"")).Count();
            var xy_boxes_2 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("XY") && x.BoxSize.Equals("4 11/16\"")).Count();
            var y_boxes_2 = box_pkg.Boxes.Where(x => x.BoxConfig.Equals("Y") && x.BoxSize.Equals("4 11/16\"")).Count();

            var total_box_cnt = box_pkg.Boxes.Count();

            void print_boxes(int qty, string labor_str)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for fire alarm box");
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Fire Alarm Boxes");

            print_boxes(d_boxes_1, "4\" Square Fire Alarm Box - D Config");
            print_boxes(i_boxes_1, "4\" Square Fire Alarm Box - I Config");
            print_boxes(x_boxes_1, "4\" Square Fire Alarm Box - X Config");
            print_boxes(t_boxes_1, "4\" Square Fire Alarm Box - T Config");
            print_boxes(xy_boxes_1, "4\" Square Fire Alarm Box - XY Config");
            print_boxes(y_boxes_1, "4\" Square Fire Alarm Box - Y Config");

            print_boxes(d_boxes_2, "4 11/16\" Square Fire Alarm Box - D Config");
            print_boxes(i_boxes_2, "4 11/16\" Square Fire Alarm Box - I Config");
            print_boxes(x_boxes_2, "4 11/16\" Square Fire Alarm Box - X Config");
            print_boxes(t_boxes_2, "4 11/16\" Square Fire Alarm Box - T Config");
            print_boxes(xy_boxes_2, "4 11/16\" Square Fire Alarm Box - XY Config");
            print_boxes(y_boxes_2, "4 11/16\" Square Fire Alarm Box - Y Config");

            code_one_sub = Math.Ceiling(code_one_sub);
            code_one_gt += code_one_sub;
            InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
            code_one_sub = 0.0;

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
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for conduit");
                var s = li.EntryName.Split(" - ");
                var red_conduit_fix = s[0] + " - Red " + s[1] + " - " + s[2];
                InsertIntoRow(red_conduit_fix, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
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
            var emt_coup_total_1 = conduit_pkg.GetTotalEmtCouplings("1/2\"");
            var emt_coup_total_2 = conduit_pkg.GetTotalEmtCouplings("3/4\"");
            var emt_coup_total_3 = conduit_pkg.GetTotalEmtCouplings("1\"");
            var emt_coup_total_4 = conduit_pkg.GetTotalEmtCouplings("1 1/4\"");
            var emt_coup_total_5 = conduit_pkg.GetTotalEmtCouplings("1 1/2\"");
            var emt_coup_total_6 = conduit_pkg.GetTotalEmtCouplings("2\"");

            var pvc_coup_total_1 = conduit_pkg.GetTotalPvcCouplings("1/2\"");
            var pvc_coup_total_2 = conduit_pkg.GetTotalPvcCouplings("3/4\"");
            var pvc_coup_total_3 = conduit_pkg.GetTotalPvcCouplings("1\"");
            var pvc_coup_total_4 = conduit_pkg.GetTotalPvcCouplings("1 1/4\"");
            var pvc_coup_total_5 = conduit_pkg.GetTotalPvcCouplings("1 1/2\"");
            var pvc_coup_total_6 = conduit_pkg.GetTotalPvcCouplings("2\"");

            /* var mc_coup_total_1 = conduit_pkg.GetTotalMcCableCouplings("1/2\"");
            var mc_coup_total_2 = conduit_pkg.GetTotalMcCableCouplings("3/4\"");
            var mc_coup_total_3 = conduit_pkg.GetTotalMcCableCouplings("1\"");
            var mc_coup_total_4 = conduit_pkg.GetTotalMcCableCouplings("1 1/4\"");
            var mc_coup_total_5 = conduit_pkg.GetTotalMcCableCouplings("1 1/2\"");
            var mc_coup_total_6 = conduit_pkg.GetTotalMcCableCouplings("2\""); */

            void print_couplings(int qty, string labor_str)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for coupling: " + labor_str);
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Couplings");

            print_couplings(emt_coup_total_1, "Coupling - Set Screw Steel - EMT - 1/2\"");
            print_couplings(emt_coup_total_2, "Coupling - Set Screw Steel - EMT - 3/4\"");
            print_couplings(emt_coup_total_3, "Coupling - Set Screw Steel - EMT - 1\"");
            print_couplings(emt_coup_total_4, "Coupling - Set Screw Steel - EMT - 1 1/4\"");
            print_couplings(emt_coup_total_5, "Coupling - Set Screw Steel - EMT - 1 1/2\"");
            print_couplings(emt_coup_total_6, "Coupling - Set Screw Steel - EMT - 2\"");

            print_couplings(pvc_coup_total_1, "Coupling - Standard - PVC - 1/2\"");
            print_couplings(pvc_coup_total_2, "Coupling - Standard - PVC - 3/4\"");
            print_couplings(pvc_coup_total_3, "Coupling - Standard - PVC - 1\"");
            print_couplings(pvc_coup_total_4, "Coupling - Standard - PVC - 1 1/4\"");
            print_couplings(pvc_coup_total_5, "Coupling - Standard - PVC - 1 1/2\"");
            print_couplings(pvc_coup_total_6, "Coupling - Standard - PVC - 2\"");

            code_one_sub = Math.Ceiling(code_one_sub);
            code_one_gt += code_one_sub;
            InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
            code_one_sub = 0.0;

            // connectors

            var emt_conn_total_1 = 0;
            var emt_conn_total_2 = 0;
            var emt_conn_total_3 = 0;
            var emt_conn_total_4 = 0;
            var emt_conn_total_5 = 0;
            var emt_conn_total_6 = 0;

            box_pkg.Boxes.ForEach(x => emt_conn_total_1 += x.GetEmtConnectorCount("1/2\""));
            box_pkg.Boxes.ForEach(x => emt_conn_total_2 += x.GetEmtConnectorCount("3/4\""));
            box_pkg.Boxes.ForEach(x => emt_conn_total_3 += x.GetEmtConnectorCount("1\""));
            box_pkg.Boxes.ForEach(x => emt_conn_total_4 += x.GetEmtConnectorCount("1 1/4\""));
            box_pkg.Boxes.ForEach(x => emt_conn_total_5 += x.GetEmtConnectorCount("1 1/2\""));
            box_pkg.Boxes.ForEach(x => emt_conn_total_6 += x.GetEmtConnectorCount("2\""));

            var pvc_conn_total_1 = 0;
            var pvc_conn_total_2 = 0;
            var pvc_conn_total_3 = 0;
            var pvc_conn_total_4 = 0;
            var pvc_conn_total_5 = 0;
            var pvc_conn_total_6 = 0;

            box_pkg.Boxes.ForEach(x => pvc_conn_total_1 += x.GetPvcConnectorCount("1/2\""));
            box_pkg.Boxes.ForEach(x => pvc_conn_total_2 += x.GetPvcConnectorCount("3/4\""));
            box_pkg.Boxes.ForEach(x => pvc_conn_total_3 += x.GetPvcConnectorCount("1\""));
            box_pkg.Boxes.ForEach(x => pvc_conn_total_4 += x.GetPvcConnectorCount("1 1/4\""));
            box_pkg.Boxes.ForEach(x => pvc_conn_total_5 += x.GetPvcConnectorCount("1 1/2\""));
            box_pkg.Boxes.ForEach(x => pvc_conn_total_6 += x.GetPvcConnectorCount("2\""));

            void print_connectors(int qty, string labor_str)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for connector: " + labor_str);
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Connectors");

            print_connectors(emt_conn_total_1, "Connector - Set Screw Steel - EMT - 1/2\"");
            print_connectors(emt_conn_total_2, "Connector - Set Screw Steel - EMT - 3/4\"");
            print_connectors(emt_conn_total_3, "Connector - Set Screw Steel - EMT - 1\"");
            print_connectors(emt_conn_total_4, "Connector - Set Screw Steel - EMT - 1 1/4\"");
            print_connectors(emt_conn_total_5, "Connector - Set Screw Steel - EMT - 1 1/2\"");
            print_connectors(emt_conn_total_6, "Connector - Set Screw Steel - EMT - 2\"");

            print_connectors(pvc_conn_total_1, "Connector - Female Adapter - PVC - 1/2\"");
            print_connectors(pvc_conn_total_2, "Connector - Female Adapter - PVC - 3/4\"");
            print_connectors(pvc_conn_total_3, "Connector - Female Adapter - PVC - 1\"");
            print_connectors(pvc_conn_total_4, "Connector - Female Adapter - PVC - 1 1/4\"");
            print_connectors(pvc_conn_total_5, "Connector - Female Adapter - PVC - 1 1/2\"");
            print_connectors(pvc_conn_total_6, "Connector - Female Adapter - PVC - 2\"");

            code_one_sub = Math.Ceiling(code_one_sub);
            code_one_gt += code_one_sub;
            InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
            code_one_sub = 0.0;

            // hangers

            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Hangers");

            var att = hangers.GroupBy(x => x.BatwingAttachment);
            var washers = hangers.GroupBy(x => x.Washer);
            var hex_nuts = hangers.GroupBy(x => x.HexNut);
            var anchors = hangers.GroupBy(x => x.Anchor);
            var threaded_rod = hangers.GroupBy(x => x.ThreadedRodSize);

            void print_hanger_hardware(int qty, string labor_str)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for hanger hardware: " + labor_str);
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            foreach (var a in att) print_hanger_hardware(a.Count(), a.Key);
            foreach (var w in washers) print_hanger_hardware(w.Count(), w.Key);
            foreach (var h in hex_nuts) print_hanger_hardware(h.Count(), h.Key);
            foreach (var a in anchors) print_hanger_hardware(a.Count(), a.Key);
            foreach (var tr in threaded_rod) print_hanger_hardware((int)(tr.Select(x => x.ThreadedRodLength).Sum() + (total_box_cnt * 10.0)), "Threaded Rod - " + tr.Key);

            code_one_sub = Math.Ceiling(code_one_sub);
            code_one_gt += code_one_sub;
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
    }
}