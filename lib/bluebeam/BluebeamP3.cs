

using System;
using System.Collections.Generic;
using System.Linq;
using MoreLinq;
using PdfSharp.Pdf.Annotations;

namespace JPMorrow.Pdf.Bluebeam.P3
{
    /// <summary>
    /// A class to handle the conversion from a long device code to a short hand device code 
    /// </summary>
    public class BluebeamP3ShorhandDeviceCodeResolver
    {
        private class ShorthandDeviceCodePair
        {
            public int shorhandCodeFirst { get; private set; }
            public int shorhandCodeSecond { get; private set; }
            public int ShorthandCodeNumber
            {
                get
                {
                    string str_cvt = shorhandCodeFirst.ToString() + shorhandCodeSecond.ToString();
                    bool s = int.TryParse(str_cvt, out int re_cvt);
                    if (!s) throw new Exception("Shorthand code integer failed to parse: " + str_cvt);
                    return re_cvt;
                }
            }

            public string ShorthandCodeString { get => ((char)shorhandCodeFirst).ToString() + shorhandCodeSecond.ToString(); }
            public string LongDeviceCode { get; private set; }

            public ShorthandDeviceCodePair(int short_first_part, int short_second_part, string long_code)
            {
                shorhandCodeFirst = short_first_part;
                shorhandCodeSecond = short_second_part;
                LongDeviceCode = long_code;
            }

            /// <summary>
            /// Use the current shorhand device code in this 
            /// structure to generate an increment of that code
            /// </summary>
            /// <param name="prev_shorhand_int"></param>
            public static ShorthandDeviceCodePair GenerateIncrementedShorhandDeviceCode(
                ShorthandDeviceCodePair previous, string long_device_code)
            {
                var first_part = previous.shorhandCodeFirst;
                var second_part = previous.shorhandCodeSecond;

                if (first_part == 26) first_part = 1;
                else first_part++;

                if (second_part == 5) second_part = 1;
                else second_part++;

                return new ShorthandDeviceCodePair(first_part, second_part, long_device_code);
            }
        }

        private List<ShorthandDeviceCodePair> Pairs { get; set; } = new List<ShorthandDeviceCodePair>();

        public BluebeamP3ShorhandDeviceCodeResolver() { }

        /// <summary>
        /// Generate and store a shorhand device code that can be retrieved 
        /// from this structure by passing it a long hand device code
        /// </summary>
        /// <param name="long_device">long device code names</param>
        /// <returns>the shorhand device code, string.empty if no device code matches</returns>
        public string StoreShorthandCode(string long_device_code)
        {
            if (Pairs.Any(x => x.LongDeviceCode.Equals(long_device_code)))
                return Pairs.Where(x => x.LongDeviceCode.Equals(long_device_code)).First().ShorthandCodeString;

            var last_code = GetMostRecentShorthandDeviceCode();
            var new_code = ShorthandDeviceCodePair.GenerateIncrementedShorhandDeviceCode(last_code, long_device_code);

            Pairs.Add(new_code);
            return new_code.ShorthandCodeString;
        }

        public string RetrieveShorthandCode(string long_device_code)
        {
            var idx = Pairs.FindIndex(x => x.LongDeviceCode.Equals(long_device_code));
            if (idx == -1) return string.Empty;
            return Pairs[idx].ShorthandCodeString;
        }

        /// <summary>
        /// Get the most recently generated shorhand device code structure
        /// </summary>
        /// <returns>DeviceCode Structure if there is one, otherwise null</returns>
        private ShorthandDeviceCodePair GetMostRecentShorthandDeviceCode()
        {
            if (!Pairs.Any()) return null;
            return Pairs.MaxBy(x => x.ShorthandCodeNumber).First();
        }


    }

    public class BluebeamP3BoxConfig
    {

    }

    public class BluebeamP3Box
    {
        public BluebeamP3BoxConfig Config { get; private set; }
        public BluebeamP3Box(BluebeamP3BoxConfig config)
        {
            Config = config;
        }
    }

    public class BluebeamP3BoxCollection
    {
        private BluebeamP3ShorhandDeviceCodeResolver BSHD_Resolver { get; set; } = new BluebeamP3ShorhandDeviceCodeResolver();
        public BluebeamP3BoxCollection()
        {

        }

        /* public static BluebeamP3BoxCollection BoxPackageFromAnnoations(IEnumerable<PdfAnnotation> box_annotations)
        {
            foreach (var annot in box_annotations)
            {
                var a = annot;
                bool has_box_subject = HasP3BoxSubject(a, out string subject);

                if (!has_box_subject) has_box_subject = HasP3BoxSubjectInChildren(a, out subject, out a);
                if (!has_box_subject) continue;

                bool is_rect = BluebeamPdfUtil.IsRectangle(a);
                bool
            }
        } */


    }
}