using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;

namespace JPMorrow.P3
{
    public class P3CSVRow
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

        public string DeviceCode { get; private set; } = "";

        public P3CSVRow(
            string subject, string box_type, string box_size,
            string gang, string plaster_ring, string connector_entry_type, string connector_size,
            int top_connectors_emt, int bottom_connectors_emt,
            int top_connectors_pvc, int bottom_connectors_pvc,
            int top_connectors_mc, int bottom_connectors_mc)
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

            ProcessFields();
        }

        public static IEnumerable<P3CSVRow> ParseCSV(string csv_filepath)
        {
            var records = new List<P3CSVRow>();

            using (var reader = new StreamReader(csv_filepath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {

                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    var subject = csv.GetField<string>("Subject");
                    var bt = csv.GetField<string>("Box Type");
                    var bs = csv.GetField<string>("Box Size");
                    var gang = csv.GetField<string>("Gang");
                    var pr = csv.GetField<string>("Plaster Ring Depth");
                    var cet = csv.GetField<string>("Entry Connector Type");
                    var cs = csv.GetField<string>("Connector Size");
                    var tce = int.Parse(csv.GetField<string>("Top Connectors - EMT"));
                    var bce = int.Parse(csv.GetField<string>("Bottom Connectors - EMT"));
                    var tcp = int.Parse(csv.GetField<string>("Top Connectors - PVC"));
                    var bcp = int.Parse(csv.GetField<string>("Bottom Connectors - PVC"));
                    var tcm = int.Parse(csv.GetField<string>("Top Connectors - MC"));
                    var bcm = int.Parse(csv.GetField<string>("Bottom Connectors - MC"));

                    var record = new P3CSVRow(subject, bt, bs, gang, pr, cet, cs, tce, bce, tcp, bcp, tcm, bcm);

                    records.Add(record);
                }
            }

            return records;
        }

        public static string PrintRaw(IEnumerable<P3CSVRow> rows)
        {
            string o = "";

            foreach (var r in rows)
            {
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
            if (BoxSize.Equals("small"))
                DeviceCode += "S";
            else if (BoxSize.Equals("extended"))
                DeviceCode += "X";

            if (Subject.Contains("Fire Alarm"))
                DeviceCode += "F";
            else if (BoxType.Equals("powered"))
                DeviceCode += "P";
            else if (BoxType.Equals("non-powered"))
                DeviceCode += "N";

            // gang
            if (Gang == "1") DeviceCode += "1";
            else if (Gang.Equals("2")) DeviceCode += "2";
            else if (Gang.Equals("3")) DeviceCode += "3";
            else if (Gang.Equals("4")) DeviceCode += "4";
            else if (Gang.Equals("round")) DeviceCode += "R";
            else if (Gang.Equals("E")) DeviceCode += "E";

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
            if (EntryConnectorType.Equals("MC")) DeviceCode += "M";
            else if (EntryConnectorType.Equals("EMT") || EntryConnectorType.Equals("PVC"))
            {
                var cc = ConnectorSize.Remove(ConnectorSize.Length - 1);

                bool s = connector_size_swap.TryGetValue(cc, out var add_cc);
                if (!s) throw new System.Exception("Connector Size could not be resolved");
                DeviceCode += add_cc + "C";
            }

            // plaster ring
            if (PlasterRing.Equals("adjustable")) DeviceCode += "-A";
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
}