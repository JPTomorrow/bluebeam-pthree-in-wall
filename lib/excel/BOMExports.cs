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

        // dictionary for applying color to cells
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
    }
}