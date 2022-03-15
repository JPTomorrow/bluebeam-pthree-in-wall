using System;
using System.Collections.Generic;
using System.Linq;
using JPMorrow.Bluebeam.FlexConduit;
using JPMorrow.Pdf.Bluebeam;
using JPMorrow.Pdf.Bluebeam.FireAlarm;
using JPMorrow.Revit.Labor;
using Draw = System.Drawing;

namespace JPMorrow.Excel
{
    public partial class ExcelOutputSheet
    {
        public void GenerateFireAlarmSheet(
            string labor_import_path, string project_title,
            BluebeamConduitPackage conduit_pkg, BlubeamFireAlarmBoxPackage box_pkg,
            BluebeamFireAlarmConnectorPackage connector_pkg,
            IEnumerable<BluebeamSingleHanger> hangers, double hanger_spacing, int tbar_cnt)
        {
            if (HasData) throw new Exception("The sheet already has data");
            string title = "M.P.A.C.T. - Fire Alarm";
            InsertHeader(title, project_title, "");

            var code_one_gt = 0.0;
            var code_one_sub = 0.0;
            static double shave_labor(double labor) => labor * 0.82;

            var entries = LaborExchange.LoadLaborFromInternalRescource();
            var l = new LaborExchange(entries);

            var total_box_cnt = box_pkg.Boxes.Count();
            BluebeamFlexConduitTotal fct = new BluebeamFlexConduitTotal(total_box_cnt);

            // boxes
            var boxes_1 = box_pkg.Boxes.Where(x => x.BoxSize.Equals("4\"")).Count();
            var boxes_2 = box_pkg.Boxes.Where(x => x.BoxSize.Equals("4 11/16\"")).Count();
            var octagon_boxes = box_pkg.Boxes.Where(x => x.BoxSize.Equals("4\" Octagon")).Count();

            void print_boxes(int qty, string labor_str)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for fire alarm box");
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Fire Alarm Boxes");

            var box_cnt_1 = boxes_1;
            print_boxes(box_cnt_1, "4\" Square Fire Alarm Box");
            var box_cnt_2 = boxes_2;
            print_boxes(box_cnt_2, "4 11/16\" Square Fire Alarm Box");
            print_boxes(octagon_boxes, "4\" Octagon Fire Alarm Box");

            void print_covers(int qty)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, "Red Blank Box Cover");
                if (!has_item) throw new Exception("No Labor item for fire alarm box cover");
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            print_covers(total_box_cnt);

            void print_brackets(int qty)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, "Caddy 812MB18A Helicopter Bracket");
                if (!has_item) throw new Exception("No Labor item for fire alarm box Caddy 812MB18A Helicopter Bracket");
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            print_brackets(total_box_cnt);

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

            var mc_total = conduit_pkg.GetTotalMcCableLengthRounded("3/4\"");

            void print_conduit(int qty, string labor_str, bool red = false)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for conduit");
                var s = li.EntryName.Split(" - ");
                var red_str = red ? " - Red " : " - ";
                var red_conduit_fix = s[0] + red_str + s[1] + " - " + s[2];
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

            void print_flex_whips(int qty, string labor_str)
            {
                if (qty == 0) return;
                var has_item = l.GetItem(out var li, qty, labor_str);
                if (!has_item) throw new Exception("No Labor item for fire alarm flex whips");
                InsertIntoRow(li.EntryName, li.Quantity, li.PerUnitLabor, li.LaborCodeLetter, li.TotalLaborValue);
                code_one_sub += li.TotalLaborValue; NextRow(1);
            }

            //get flex count
            var flex_boxes = box_pkg.Boxes.Where(x => x.HasMc);
            var oh_flex_boxes = flex_boxes.Where(x => x.McSize.Equals("1/2")).Count();
            var oq_flex_boxes = flex_boxes.Where(x => x.McSize.Equals("3/4")).Count();

            print_flex_whips(oh_flex_boxes, BluebeamFlexConduitTotal.OneHalfFlexPipeName);
            print_flex_whips(oq_flex_boxes, BluebeamFlexConduitTotal.ThreeQuarterFlexPipeName);

            // print_conduit(mc_total, "Conduit - FMC - 3/4\"", true);

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

            emt_conn_total_1 += connector_pkg.GetEmtConnectorCount("1/2\"");
            emt_conn_total_2 += connector_pkg.GetEmtConnectorCount("3/4\"");
            emt_conn_total_3 += connector_pkg.GetEmtConnectorCount("1\"");
            emt_conn_total_4 += connector_pkg.GetEmtConnectorCount("1 1/4\"");
            emt_conn_total_5 += connector_pkg.GetEmtConnectorCount("1 1/2\"");
            emt_conn_total_6 += connector_pkg.GetEmtConnectorCount("2\"");

            var pvc_conn_total_1 = 0;
            var pvc_conn_total_2 = 0;
            var pvc_conn_total_3 = 0;
            var pvc_conn_total_4 = 0;
            var pvc_conn_total_5 = 0;
            var pvc_conn_total_6 = 0;

            pvc_conn_total_1 += connector_pkg.GetPvcConnectorCount("1/2\"");
            pvc_conn_total_2 += connector_pkg.GetPvcConnectorCount("3/4\"");
            pvc_conn_total_3 += connector_pkg.GetPvcConnectorCount("1\"");
            pvc_conn_total_4 += connector_pkg.GetPvcConnectorCount("1 1/4\"");
            pvc_conn_total_5 += connector_pkg.GetPvcConnectorCount("1 1/2\"");
            pvc_conn_total_6 += connector_pkg.GetPvcConnectorCount("2\"");

            // var mc_conn_total = 0;
            // mc_conn_total += connector_pkg.GetMcCableConnectorCount("3/4\"");

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

            print_connectors(oh_flex_boxes, BluebeamFlexConduitTotal.OneHalfFlexConnectorName);
            print_connectors(oq_flex_boxes, BluebeamFlexConduitTotal.ThreeQuarterFlexConnectorName);


            // print_connectors(fct.ConnectorQty, BluebeamFlexConduitTotal.OneHalfFlexPipeName);

            code_one_sub = Math.Ceiling(code_one_sub);
            code_one_gt += code_one_sub;
            InsertGrandTotal("Sub Total", ref code_one_sub, true, false, true);
            code_one_sub = 0.0;

            // hangers

            InsertSingleDivider(Draw.Color.SlateGray, Draw.Color.White, "Hangers");

            var att = hangers.GroupBy(x => x.BatwingAttachment);
            // var washers = hangers.GroupBy(x => x.Washer);
            // var hex_nuts = hangers.GroupBy(x => x.HexNut);
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
            // foreach (var w in washers) print_hanger_hardware(w.Count(), w.Key);
            // foreach (var h in hex_nuts) print_hanger_hardware(h.Count(), h.Key);
            print_hanger_hardware(total_box_cnt, "Caddy 4Z34812M Batwing/Clamp");

            int att_cnt = att.Sum(x => x.Count());
            print_hanger_hardware((int)Math.Round(att_cnt * 0.10), "Unistrut Strap - 3/4\"");

            print_hanger_hardware(total_box_cnt * 2, "Washer - 1/4\"");
            print_hanger_hardware(total_box_cnt * 2, "Hex Nut - 1/4\"");
            print_hanger_hardware(tbar_cnt, "24\" Span T-Bar Hanger - Caddy 512");
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