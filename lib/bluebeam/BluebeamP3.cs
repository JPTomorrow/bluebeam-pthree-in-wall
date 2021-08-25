

using System;
using System.Collections.Generic;
using System.Linq;

namespace JPMorrow.Pdf.Bluebeam.P3
{
    /// <summary>
    /// A class to handle the conversion from a short hand device code 
    /// that can fit on a legend, and a long device codes
    /// </summary>
    public class BluebeamP3ShorhandDeviceCodeResolver
    {
        private class ShorthandDeviceCodePair
        {
            private int shorhandCodeFirst { get; set; }
            private int shorhandCodeSecond { get; set; }
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

            public string ShorthandCodeString { get => ((char)shorhandCodeFirst) + shorhandCodeSecond.ToString(); }
            public string LongDeviceCode { get; private set; }

            public ShorthandDeviceCodePair(int short_first_part, int short_second_part, string long_code)
            {
                shorhandCodeFirst = short_first_part;
                shorhandCodeSecond = short_second_part;
                LongDeviceCode = long_code;
            }
        }

        private List<ShorthandDeviceCodePair> Pairs { get; set; } = new List<ShorthandDeviceCodePair>();

        public BluebeamP3ShorhandDeviceCodeResolver() { }

        /// <summary>
        /// START HERE
        /// </summary>
        /// <param name="long_device">long device code names</param>
        /// <returns>the shorhand device code</returns>
        /* public string ProcessDeviceCode(string long_device_code)
        {
            if (Pairs.Any(x => x.LongDeviceCode.Equals(long_device_code)))



        } */
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
        public BluebeamP3BoxCollection()
        {

        }


    }
}