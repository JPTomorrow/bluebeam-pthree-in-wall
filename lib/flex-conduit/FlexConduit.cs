
using System;
using JPMorrow.Revit.Labor;

namespace JPMorrow.Bluebeam.FlexConduit
{
    public class BluebeamFlexConduitTotal
    {

    }

    public static class BluebeamFlexConduit
    {
        public static BluebeamFlexConduitTotal GetTotalFromBoxCount(int box_count)
        {
            BluebeamFlexConduitTotal total = new BluebeamFlexConduitTotal();

            // get 70% of box count
            int adjusted_box_cnt = (int)Math.Round(box_count * 0.70);

            string flex_pipe_name = "Steel Flex - 1/2\"";
            string flex_coupling_name = "Flex - Set Screw Steel - Coupling";

            /* string bc_name = "Blank Box Cover";
            string gs_name = "Ground Stinger";
            string bracket_name = "Helicopter Bracket";
            string strap_name = "Caddy Conduit Straps 812M4I";
            string washer_name = "Washer - 1/4\"";
            string hn_name = "Hex Nut - 1/4\"";

            bool has_box_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(box_name));
            bool has_blank_cover_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(bc_name));
            bool has_stinger_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(gs_name));
            bool has_bracket_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(bracket_name));
            bool has_strap_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(strap_name));
            bool has_washer_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(washer_name));
            bool has_hex_entry = ALS.AppData.LaborHourEntries.Any(x => x.EntryName.Equals(hn_name));

            void add_entry(string name, double labor, LetterCodePair pair)
            {
                var ldata = new LaborData(pair, labor);
                LaborEntry entry = new LaborEntry(name, ldata);
                ALS.AppData.LaborHourEntries.Add(entry);
            }

            // make harware entries for jboxes
            if (!has_box_entry)
                add_entry(box_name, 0.23, LaborExchange.LetterCodes.GetByLetter('E'));

            if (!has_blank_cover_entry)
                add_entry(bc_name, 0.23, LaborExchange.LetterCodes.GetByLetter('E'));

            if (!has_stinger_entry)
                add_entry(gs_name, 0.23, LaborExchange.LetterCodes.GetByLetter('E'));

            if (!has_bracket_entry)
                add_entry(bracket_name, 0.5, LaborExchange.LetterCodes.GetByLetter('E'));

            if (!has_strap_entry)
                add_entry(strap_name, 0.07, LaborExchange.LetterCodes.GetByLetter('C'));

            if (!has_washer_entry)
                add_entry(washer_name, 0.23, LaborExchange.LetterCodes.GetByLetter('E'));

            if (!has_hex_entry)
                add_entry(hn_name, 0.23, LaborExchange.LetterCodes.GetByLetter('E')); */

            return total;
        }
    }
}