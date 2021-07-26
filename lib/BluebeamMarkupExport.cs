using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using JPMorrow.PDF;

namespace JPMorrow.Bluebeam.Markup
{
    public class P3BluebeamFDFMarkup
    {
        public string Subject { get; private set; }
        public string BoxType { get; private set; }
        public string BoxSize { get; private set; }
        public string Gang { get; private set; }
        public string PlasterRing { get; private set; }
        public string EntryConnectorType { get; private set; }
        public string ConnectorSize { get; private set; }

        public int TopConnectorsEmt { get; private set; }
        public int TopConnectorsPvc { get; private set; }
        public int TopConnectorsMc { get; private set; }

        public int BottomConnectorsEmt { get; private set; }
        public int BottomConnectorsPvc { get; private set; }
        public int BottomConnectorsMc { get; private set; }

        public BundleRegion BundleRegion { get; private set; } = null;

        public string DeviceCode { get; private set; } = "";

        public P3BluebeamFDFMarkup(
            string subject, string box_type, string box_size,
            string gang, string plaster_ring, string connector_entry_type, string connector_size,
            int top_connectors_emt, int bottom_connectors_emt,
            int top_connectors_pvc, int bottom_connectors_pvc,
            int top_connectors_mc, int bottom_connectors_mc, BundleRegion bundle_region = null)
        {
            Subject = subject;
            BoxType = box_type;
            Gang = gang;
            PlasterRing = plaster_ring;
            BoxSize = box_size;
            EntryConnectorType = connector_entry_type;
            ConnectorSize = connector_size;


            TopConnectorsEmt = top_connectors_emt;
            TopConnectorsPvc = top_connectors_pvc;
            TopConnectorsMc = top_connectors_mc;

            BottomConnectorsEmt = bottom_connectors_emt;
            BottomConnectorsPvc = bottom_connectors_pvc;
            BottomConnectorsMc = bottom_connectors_mc;

            if(bundle_region != null) BundleRegion = bundle_region;

            ProcessFields();
        }

        public static P3BluebeamFDFMarkup ParseMarkup(Dictionary<string, string> fdf_values, BundleRegion bundle_region = null)
        {
            /* ConsoleDebugger.Debug.Show("\nProcessing Markup Properties {", false);
            foreach (KeyValuePair<string, string> kvp in fdf_values)
            {
                ConsoleDebugger.Debug.Show(string.Format("\t{0} : {1}", kvp.Key, kvp.Value));
            }
            ConsoleDebugger.Debug.Show("}"); */

            var subject =   fdf_values["Subject"];
            var bt =        fdf_values["Box Type"];
            var bs =        fdf_values["Box Size"];
            var gang =      fdf_values["Gang"];
            var pr =        fdf_values["Plaster Ring"];
            var cet =       fdf_values["Conduit Entry Connector"];
            var cs =        fdf_values["Connector Size"];
            var tce =       int.Parse(fdf_values["Top EMT Connectors"]);
            var bce =       int.Parse(fdf_values["Bottom EMT Connectors"]);
            var tcp =       int.Parse(fdf_values["Top PVC Connectors"]);
            var bcp =       int.Parse(fdf_values["Bottom PVC Connectors"]);
            var tcm =       int.Parse(fdf_values["Top MC Connectors"]);
            var bcm =       int.Parse(fdf_values["Bottom MC Connectors"]);
            var record = new P3BluebeamFDFMarkup(subject, bt, bs, gang, pr, cet, cs, tce, bce, tcp, bcp, tcm, bcm, bundle_region);
            return record;
        }

        public static string PrintRaw(IEnumerable<P3BluebeamFDFMarkup> rows)
        {
            string o = "";

            foreach(var r in rows) {
                o += string.Format(
                    "{0}-{1}-{2}-{3}-{4}-{5}-{6}-{7}-{8}-{9} | {10}\n",
                    r.Subject, r.BoxType, r.Gang, r.PlasterRing, 
                    r.TopConnectorsEmt, r.TopConnectorsPvc, r.TopConnectorsMc, 
                    r.BottomConnectorsEmt, r.BottomConnectorsPvc, r.BottomConnectorsMc,
                    r.DeviceCode
                );
            }

            return o;
        }

        private void ProcessFields()
        {
            if(BoxSize.Equals("small"))
                DeviceCode += "S";
            else if(BoxSize.Equals("extended"))
                DeviceCode += "X";
                
            if(Subject.Contains("Fire Alarm"))
                DeviceCode += "F";
            else if(BoxType.Equals("powered"))
                DeviceCode += "P";
            else if(BoxType.Equals("non-powered"))
                DeviceCode += "N";

            // gang
            if(Gang == "1") DeviceCode += "1";
            else if(Gang.Equals("2")) DeviceCode += "2";
            else if(Gang.Equals("3")) DeviceCode += "3";
            else if(Gang.Equals("4")) DeviceCode += "4";
            else if(Gang.Equals("round")) DeviceCode += "R";
            else if(Gang.Equals("E")) DeviceCode += "E";
            
            var connector_size_swap = new Dictionary<string, string>() 
            {
                { "1/2",    "2"},
                { "3/4",    "3"},
                { "1",      "4"},
                { "1 1/4",  "5"},
                { "1 1/2",  "6"},
                { "2",      "8"}
            };
            
            // connector size
            if(EntryConnectorType.Equals("MC")) DeviceCode += "M";
            else if(EntryConnectorType.Equals("EMT") || EntryConnectorType.Equals("PVC"))
            {
                var cc = ConnectorSize.Remove(ConnectorSize.Length - 1);
                
                bool s = connector_size_swap.TryGetValue(cc, out var add_cc);
                if(!s) throw new System.Exception("Connector Size could not be resolved");
                DeviceCode += add_cc + "C";
            }
            
            // plaster ring
            if(PlasterRing.Equals("adjustable")) DeviceCode += "-A";
            else 
            {
                var pr = PlasterRing.Remove(PlasterRing.Length - 1);
                DeviceCode += "-" + pr;
            }

            for (var i = 0; i < TopConnectorsEmt; i++) 
                DeviceCode += "-CT";
            for (var i = 0; i < BottomConnectorsEmt; i++) 
                DeviceCode += "-CB";

            for (var i = 0; i < TopConnectorsPvc; i++) 
                DeviceCode += "-PT";
            for (var i = 0; i < BottomConnectorsPvc; i++) 
                DeviceCode += "-PB";

            for (var i = 0; i < TopConnectorsMc; i++) 
                DeviceCode += "-MT";
            for (var i = 0; i < BottomConnectorsMc; i++) 
                DeviceCode += "-MB";
        }
    }

    public class BluebeamP3MarkupExport
    {
        public static string ExportFileExt { get => ".fdf"; }
        public string InputFilePath { get; private set; }
        public string OutputFilePath { get; private set; }
        private List<string> FDF_FileLines { get; set; } = new List<string>();

        private List<P3BluebeamFDFMarkup> markups { get; set; } = new List<P3BluebeamFDFMarkup>();
        public IList<P3BluebeamFDFMarkup> Markups { get => markups; }

        public string TextContent { get => string.Join("\n", FDF_FileLines); }

        public BluebeamP3MarkupExport(string input_file_path, string output_file_path) 
        {
            ConsoleDebugger.Debug.Show("-------------------------------------------------------", false);
            InputFilePath = input_file_path;
            OutputFilePath = output_file_path;
            var text = File.ReadAllText(InputFilePath);
            FDF_FileLines = text.Split("\n").ToList();

            while(CurrentFileLineIdx != FDF_FileLines.Count() - 1)
            {
                ParseNextP3BoxFDFObject();
            }

            CurrentFileLineIdx = 0;
            ConsoleDebugger.Debug.Show("Finished Constructing BluebeamP3MarkupExport", false);
            ConsoleDebugger.Debug.Show("-------------------------------------------------------", false);
        }

        public void Save()
        {
            File.WriteAllText(OutputFilePath, TextContent);
        }

        public string PrintRawInputFile()
        {
            var raw_print = string.Join("\n", FDF_FileLines);
            return raw_print;
        }

        public string PrintRawOutputFile()
        {
            if(!File.Exists(OutputFilePath)) return "";
            var txt = File.ReadAllText(OutputFilePath);
            return txt;
        }

        private int CurrentFileLineIdx = 0;

        /* private int FindNextFDFObject()
        {
            var line_chk = "/Subtype/Square";
            var idx = FDF_FileLines.IndexOf(line_chk, CurrentFileLineIdx);
            return idx;
        } */

        private int FindNextFDFObject()
        {
            var line_chk = "/Subtype/FreeText";
            for (var i = CurrentFileLineIdx; i < FDF_FileLines.Count(); i++)
            {
                string current_line = FDF_FileLines[i];
                if(current_line.Contains(line_chk))
                {
                    var next_line_idx = i + 1;
                    CurrentFileLineIdx = next_line_idx;
                    return i;
                }
            }
            
            return -1;
        }
        
        // The order that the parse string come in as is as follows:
        // Box Type - Box Size - Connector Size - CT - CB - PT - PB - MT- MB - GANG - Plaster Ring Size - Entry Connector Type
        private Dictionary<string, string> GetFDFColumnProperties(string object_line)
        {
            var column_raw_data = GetBetween(object_line, "/BSIColumnData", ")]");
            var split_props = column_raw_data.Split(")(").ToList();

            split_props = split_props
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim('[', '(', ')', ']', ' ')).ToList();

            var dict = new Dictionary<string, string>() {
                {"Box Type"                  , split_props[0] },
                {"Box Size"                  , split_props[1] },
                {"Gang"                      , split_props[9] },
                {"Plaster Ring"              , split_props[10] },
                {"Conduit Entry Connector"   , split_props[11] },
                {"Connector Size"            , split_props[2] },
                {"Top EMT Connectors"        , split_props[3] },
                {"Bottom EMT Connectors"     , split_props[4] },
                {"Top PVC Connectors"        , split_props[5] },
                {"Bottom PVC Connectors"     , split_props[6] },
                {"Top MC Connectors"         , split_props[7] },
                {"Bottom MC Connectors"      , split_props[8] },
            };

            return dict;
        }

        private string GetFDFObjectSubject(string obj_line)
        {
            var expr = @"\/Subj\([a-zA-Z\s]+\)";
            var rgx = Regex.Match(obj_line, expr);
            ConsoleDebugger.Debug.Show("\nGetting Subject", false);

            if(rgx.Success)
            {
                var final_str = GetBetween(rgx.Value, "/Subj(", ")");
                ConsoleDebugger.Debug.Show("Subject: " + final_str, false);
                 return final_str;
            }
            else 
            {
                ConsoleDebugger.Debug.Show("No Subject line was gathered", true);
                return "";
            }
        }

        public string UpdateFDFObjectTextContent(string obj_line, string update_txt)
        {
            var new_contents = "/Contents(" + update_txt + ")";
            var expr = @"\/Contents\([a-zA-Z\s]+\)";
            ConsoleDebugger.Debug.Show("\nGetting Text Content", false);

            if(Regex.Match(obj_line, expr).Success)
            {
                
                var replace = Regex.Replace(obj_line, expr, new_contents);
                ConsoleDebugger.Debug.Show("Replaced Text content: " + Regex.Match(obj_line, expr).Value + " with " + new_contents, false);
                return replace;
            }
            else 
            {
                ConsoleDebugger.Debug.Show("Failed to replace text content", false);
                return "";
            }
        }

        private void ParseNextP3BoxFDFObject()
        {
            var obj_idx = FindNextFDFObject();
            ConsoleDebugger.Debug.Show("\nParsing Box Object at index: " + obj_idx.ToString(), false);

            if (obj_idx == -1 || (obj_idx + 1) > FDF_FileLines.Count() - 1)
            {
                ConsoleDebugger.Debug.Show("The object index has overan EOF or is -1", true);
                CurrentFileLineIdx = FDF_FileLines.Count() - 1;
                return;
            }
                

            // parse rect
            var obj_line = FDF_FileLines[obj_idx];
            ConsoleDebugger.Debug.Show("Object line: " + obj_line, false);
            var props = GetFDFColumnProperties(obj_line);
            props.Add("Subject", GetFDFObjectSubject(obj_line));
            ConsoleDebugger.Debug.Show("Prepended Subject to dictionary", false);
            var markup = P3BluebeamFDFMarkup.ParseMarkup(props);
            markups.Add(markup);

            FDF_FileLines[obj_idx] = UpdateFDFObjectTextContent(obj_line, markup.DeviceCode);
            ConsoleDebugger.Debug.Show("\nChanged Obejct Line: " + FDF_FileLines[obj_idx], false);
        }

        private static string GetBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }

    }
}