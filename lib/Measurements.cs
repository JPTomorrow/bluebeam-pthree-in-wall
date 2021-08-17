/*
    Author: Justin Morrow
    Date Created: 4/21/2021
    Description: A Module that helps with Revit measurement conversions
*/

using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace JPMorrow.Measurements
{
    public static class Measure 
    {
        public static double LengthDbl(string cvt_str) 
        {
            return LengthConverter.ToDouble(cvt_str);
        }

        public static string LengthFromDbl(double dbl) 
        {
            return LengthConverter.ToString(dbl);
        } 

        private static class LengthConverter
        {
            public static double ToDouble(string length)
            {
                string exp1 = "^\\d+[']\\s\\d+\\s\\d+[\\/]\\d+[\"]$"; // 11' 1 1/2"
                string exp2 = "^\\d+[']$"; // 11'
                string exp3 = "^\\d+\\s\\d+[\\/]\\d+[\"]$"; // 1 1/2"
                string exp4 = "^\\d+[\\/]\\d+[\"]$"; // 1/2"
                string exp5 = "^\\d+[\"]$"; // 2"
                string exp6 = "^\\d+[']\\s\\d+[\"]$"; // 11' 7"
                string exp7 = "^\\d+[']\\s\\d+\\s\\d+[\\/]\\d+[\"]$"; // 11' 7 1/2"

                if( !Regex.Match(length, exp1).Success &&
                    !Regex.Match(length, exp2).Success &&
                    !Regex.Match(length, exp3).Success &&
                    !Regex.Match(length, exp4).Success &&
                    !Regex.Match(length, exp5).Success &&
                    !Regex.Match(length, exp6).Success &&
                    !Regex.Match(length, exp7).Success)
                {
                    return -1;
                }

                if(Regex.Match(length, exp1).Success)
                {
                    var split = length.Split(" ").ToList();
                    var feet_raw = split.First();
                    feet_raw = feet_raw.Remove(feet_raw.Length - 1);
                    split.Remove(split.First());
                    var inch_raw = string.Join(" ", split);
                    inch_raw = inch_raw.Remove(inch_raw.Length - 1);

                    double final = 0.0;
                    // parse feet
                    var feet = int.Parse(feet_raw);
                    final += feet;

                    // parse inches
                    var whole_in = int.Parse(inch_raw.Split(" ").First());
                    var fractional_top = inch_raw.Split(" ").Last().Split("/").First();
                    var fractional_bottom = inch_raw.Split(" ").Last().Split("/").Last();
                    double fraction = (double.Parse(fractional_top) / double.Parse(fractional_bottom));
                    final += ((double)whole_in / 12.0);
                    final += fraction;

                    return final;
                }
                else if(Regex.Match(length, exp2).Success)
                {
                    var feet_raw = length.ToString();
                    feet_raw = feet_raw.Remove(feet_raw.Length - 1);

                    double final = 0.0;

                    // parse feet
                    var feet = int.Parse(feet_raw);
                    final += feet;

                    return final;
                }
                else if(Regex.Match(length, exp3).Success)
                {
                    var inch_raw = length;
                    inch_raw = inch_raw.Remove(inch_raw.Length - 1);

                    double final = 0.0;

                    // parse inches
                    var whole_in = int.Parse(inch_raw.Split(" ").First());
                    var fractional_top = inch_raw.Split(" ").Last().Split("/").First();
                    var fractional_bottom = inch_raw.Split(" ").Last().Split("/").Last();
                    double fraction = (double.Parse(fractional_top) / double.Parse(fractional_bottom));
                    final += ((double)whole_in / 12.0);
                    final += fraction;

                    return final;
                }
                else if(Regex.Match(length, exp4).Success)
                {
                    var inch_raw = length;
                    inch_raw = inch_raw.Remove(inch_raw.Length - 1);

                    double final = 0.0;

                    var fractional_top = inch_raw.Split("/").First();
                    var fractional_bottom = inch_raw.Split("/").Last();
                    double fraction = (double.Parse(fractional_top) / double.Parse(fractional_bottom));
                    final += fraction;

                    return final;
                }
                else if(Regex.Match(length, exp5).Success)
                {
                    var inch_raw = length;
                    inch_raw = inch_raw.Remove(inch_raw.Length - 1);
                    double final = double.Parse(inch_raw) / 12.0;
                    return final;
                }
                else if(Regex.Match(length, exp6).Success)
                {
                    var split = length.Split(" ").ToList();
                    var feet_raw = split.First();
                    feet_raw = feet_raw.Remove(feet_raw.Length - 1);
                    split.Remove(split.First());
                    var inch_raw = string.Join(" ", split);
                    inch_raw = inch_raw.Remove(inch_raw.Length - 1);

                    double final = 0.0;

                    // parse feet
                    var feet = int.Parse(feet_raw);
                    final += feet;

                    // parse inches
                    final += int.Parse(inch_raw) / 12.0;

                    return final;
                }
                else if(Regex.Match(length, exp7).Success)
                {
                    var split = length.Split(" ").ToList();
                    var feet_raw = split.First();
                    feet_raw = feet_raw.Remove(feet_raw.Length - 1);
                    split.Remove(split.First());
                    var inch_raw = string.Join(" ", split);
                    inch_raw = inch_raw.Remove(inch_raw.Length - 1);

                    double final = 0.0;

                    // parse feet
                    var feet = int.Parse(feet_raw);
                    final += feet;

                    // parse inches
                    var whole_in = int.Parse(inch_raw.Split(" ").First());
                    var fractional_top = inch_raw.Split(" ").Last().Split("/").First();
                    var fractional_bottom = inch_raw.Split(" ").Last().Split("/").Last();
                    double fraction = (double.Parse(fractional_top) / double.Parse(fractional_bottom));
                    final += ((double)whole_in / 12.0);
                    final += fraction;

                    return final;
                }
                else
                {
                    return -1;
                }
            }

            public static string ToString(double length)
            {
                string final = "";
                double whole = Math.Truncate(length);
                var remainder = length - whole;

                if(whole > 0.0)
                {
                    final += ((int)whole).ToString() + "\'";
                    if(remainder > 0.0)
                        final += " ";
                }
                    

                if(remainder > 0.0)
                {
                    var fraction = ToFraction64(remainder);
                    final += fraction + "\"";
                }

                return length <= 0.0 ? "0' 0\"" : final;
            }

            private static string ToFraction64(double value) 
            {
                // denominator is fixed
                int denominator = 64;
                // integer part, can be signed: 1, 0, -3,...
                int integer = (int) value;
                // numerator: always unsigned (the sign belongs to the integer part)
                // + 0.5 - rounding, nearest one: 37.9 / 64 -> 38 / 64; 38.01 / 64 -> 38 / 64
                int numerator = (int) ((Math.Abs(value) - Math.Abs(integer)) * denominator + 0.5);
                
                while ((numerator % 2 == 0) && (denominator % 2 == 0)) {
                    numerator /= 2;
                    denominator /= 2;
                }

                if (denominator > 1)
                {
                    if (integer != 0) // all three: integer + numerator + denominator
                        return string.Format("{0} {1}/{2}", integer, numerator, denominator);
                    else if (value < 0) // negative numerator/denominator, e.g. -1/4
                        return string.Format("-{0}/{1}", numerator, denominator);
                    else // positive numerator/denominator, e.g. 3/8
                        return string.Format("{0}/{1}", numerator, denominator);
                }
                else 
                    return integer.ToString(); // just an integer value, e.g. 0, -3, 12...  
            }
        }
    }
}
