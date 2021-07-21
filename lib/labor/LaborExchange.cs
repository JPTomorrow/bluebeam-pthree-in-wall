
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Newtonsoft.Json.Linq;

namespace JPMorrow.Revit.Labor
{
    /// <summary>
	/// use this class as an override for the labor exchange
	/// when you have a very specific labor setup
	/// </summary>
	public class LaborDefinition
	{
		public string EntryName { get; }
		public double LaborUnit { get; }
		public string LaborCode { get; }
		public double ResolveLabor(int cnt) => cnt * LaborUnit;

		public LaborDefinition(string name, double unit, string code)
		{
			EntryName = name;
			LaborUnit = unit;
			LaborCode = code;
		}
	}

    [DataContract]
    public class LaborItem
    {
        [DataMember]
        public LaborEntry LaborEntry { get; private set; }
        [DataMember]
        public double Quantity { get; private set; }
        
        public string DisplayQuantity { get => ((int)Math.Round(Quantity)).ToString(); }
        public double TotalLaborValue { get =>  LaborEntry.LaborData.PerUnitLabor * Quantity; }
        public double PerUnitLabor { get => LaborEntry.LaborData.PerUnitLabor; }
        public string LaborCodeLetter { get => LaborEntry.LaborData.LaborCodePair.Letter; }
        public string EntryName { get => LaborEntry.EntryName; }

        public LaborItem(LaborEntry entry, double qty)
        {
            LaborEntry = entry;
            Quantity = qty;
        }   

        public string MakeEntryName(string prefix = "", string postfix = "")
        {
            var prespace = prefix == "" ? "" : " ";
            var postspace = postfix == "" ? "" : " ";
            return prefix + prespace + EntryName + postspace + postfix;
        }

        public string PerUnitWithSuffix(string suffix)
        {
            return PerUnitLabor.ToString() + " " + suffix;
        }
    }

    // single named entry for labor
	[DataContract]
	public class LaborEntry
	{
        [DataMember]
        public string EntryName { get; private set; }
        [DataMember]
        public LaborData LaborData { get; private set; }

        public LaborEntry(string entry_name, LaborData data)
        {
            EntryName = entry_name;
            LaborData = data;
        }

        public double GetLaborValue(int qty) => LaborData.PerUnitLabor * qty;

        public static string PrintLaborEntries(IEnumerable<LaborEntry> entries)
        {
            string o = "";

            foreach(var e in entries)
            {
                o += e.EntryName + "\t{ " + 
                    e.LaborData.LaborCodePair.Letter + ", " + 
                    e.LaborData.PerUnitLabor + " },\n";
            }

            return o;
        }
    }

    // information for unit labor
    [DataContract]
    public class LaborData
    {
        [DataMember]
        public LetterCodePair LaborCodePair { get; private set; }
        [DataMember]
        public double PerUnitLabor { get; private set; }

        public LaborData(LetterCodePair pair, double per_unit_labor)
        {
            LaborCodePair = pair;
            PerUnitLabor = per_unit_labor;
        }
    }

    [DataContract]
    public class LetterCodePair
    {
        [DataMember]
        public string Letter { get; private set; }
        [DataMember]
        public LaborCode Code { get; private set; }
        
        public int Qty { get => (int)Code; }

        public LetterCodePair(string letter, LaborCode code)
        {
            Letter = letter;
            Code = code;
        }
    }

    public class LetterCodeCollection
    {
        private List<LetterCodePair> _pairs { get; set; } = new List<LetterCodePair>();
        public IList<LetterCodePair> Pairs { get => _pairs; }

        public LetterCodeCollection()
        {
            _pairs.AddRange(new LetterCodePair[] {
                new LetterCodePair("E", LaborCode.Each),
                new LetterCodePair("C", LaborCode.Per_Hundred),
                new LetterCodePair("M", LaborCode.Per_Thousand),
            });
        }

        public LetterCodePair GetByLaborCode(LaborCode code)
        {
            return _pairs.First(x => x.Code == code);
        }

        public LetterCodePair GetByLetter(char letter)
        {
            var idx = _pairs.FindIndex(x => x.Letter.Contains(letter));
            
            if(idx == -1)
                throw new Exception("The specified letter does not exist in the collection");

            return _pairs[idx];
        }
    }

	public enum LaborCode
	{
		Each = 1,
		Per_Hundred = 100,
		Per_Thousand = 1000
	}

	public class LaborExchange
	{
        public static LetterCodeCollection LetterCodes { get; } = new LetterCodeCollection();
        private List<LaborEntry> CheckEntries { get; set; }
        
        private List<LaborItem> _items { get; set; }
        public IList<LaborItem> Items { get => _items; }

        public LaborExchange(IEnumerable<LaborEntry> entries = null)
        {
            if(entries == null) entries = new List<LaborEntry>();
            CheckEntries = entries.ToList();
        }

        public bool GetEntry(out LaborEntry entry, params string[] part_segment_names)
        {
            entry = null;
            var part_name = string.Join(" - ", part_segment_names);
            int idx = CheckEntries.FindIndex(x => x.EntryName.Equals(part_name));

            if(idx == -1) return false;
            else
            {
                LaborEntry found_entry = CheckEntries[idx];
                entry = found_entry;
            }

            return true;
        }

        public bool GetItem(out LaborItem item, double qty, params string[] part_segment_names)
        {
            item = null;
            bool s = GetEntry(out var entry, part_segment_names);
            if(!s) return false;
            item = new LaborItem(entry, qty);
            return true;
        }

		public static LaborEntry MakeLaborEntry(string name, LaborCode code, double per_unit)
		{
			LaborEntry entry = new LaborEntry(name, new LaborData(LetterCodes.GetByLaborCode(code), per_unit));
			return entry;
		}

        public static IEnumerable<LaborEntry> LoadLaborFromFile(string file_path)
		{
			List<LaborEntry> ret_entries = new List<LaborEntry>();
            using StreamReader r = new StreamReader(file_path);
            string json = r.ReadToEnd();
            var con = (JContainer)JToken.Parse(json);


            var query = (con as JContainer).DescendantsAndSelf().OfType<JArray>();

            foreach(var q in query)
            {
                //if(!(q as JContainer).Path.ToLower().Contains("4 11/16\" square plaster ring")) continue;

                var container = (q as JContainer);
                var path = container.Path;
                char code_char;
                double labor_per_unit;
                //debugger.debug_show(err:path);

                if(q.Count == 1)
                {
                    path = path.Replace("['", "");
                    path = path.Replace("']", "");
                    var val = q.Value<JArray>();
                    code_char = (char)val[0];
                    labor_per_unit = (double)val[1];
                }
                else
                {
                    /* foreach (var i in path.Where(x => x == '[').Select(x => path.IndexOf(x)))
                    {
                        debugger.debug_show(err:i.ToString());
                        var idx = i - 1;
                        if(idx < 0) continue;
                        StringBuilder ss = new StringBuilder(path);

                        if(path[idx] == '.')
                        {
                            ss.Remove(idx, 3);
                            ss.Insert(idx, " - ", 1);
                        }
                        else
                        {
                            ss.Remove(i, 2);
                            ss.Insert(i, " - ", 1);
                        }
                        
                        path = ss.ToString();
                    } */

                    /* foreach (var i in path.Where(x => x == '.').Select(x => path.IndexOf(x)))
                    {
                        //debugger.debug_show(err:i.ToString());
                        var idx1 = i - 1;
                        var idx2 = i + 1;
                        if(idx1 <= 0 || path[idx1] == ' ') continue;
                        if(idx2 > path.Length || path[idx2] == ' ') continue;
                        StringBuilder ss = new StringBuilder(path);
                        ss.Remove(i, 1);
                        ss.Insert(i, " - ", 1);
                        path = ss.ToString();
                    } */

                    //debugger.debug_show(err:"after period: " + path);

                    if(path.StartsWith("['"))
                        path = path.Remove(0, 2);

                    if(path.EndsWith("']"))
                        path = path.Remove(path.Length - 2, 2);

                    //debugger.debug_show(err:"after first brackets: " + path);

                    path = path.Replace("']['", " - ");
                    path = path.Replace("'].", " - ");
                    path = path.Replace(".['", " - ");
                    
                    foreach (var i in path.Where(x => x == '.').Select(x => path.IndexOf(x)))
                    {
                        //debugger.debug_show(err:i.ToString());
                        var idx1 = i - 1;
                        var idx2 = i + 1;
                        if(idx1 <= 0 || path[idx1] == ' ') continue;
                        if(idx2 > path.Length || path[idx2] == ' ') continue;
                        StringBuilder ss = new StringBuilder(path);
                        ss.Remove(i, 1);
                        ss.Insert(i, " - ", 1);
                        path = ss.ToString();
                    }
                    
                    path = path.Replace("']", " - ");
                    path = path.Replace("['", " - ");

                    //debugger.debug_show(err:"after second brackets: " + path);

                    code_char = (char)q[0];
                    labor_per_unit = (double)q[1];
                }
                //debugger.debug_show(err:"finished: " + path);
                
                var entry = new LaborEntry(path, new LaborData(LetterCodes.GetByLetter(code_char), labor_per_unit));
                ret_entries.Add(entry);
            }

            return ret_entries;
		}

        public void GenTwentyTestCases()
		{
            List<object[]> parts = new List<object[]>() { 
                new object[] { "Screw", "Big", "Many", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.001  ) },
                new object[] { "Steve", "John", "Doe", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.05   ) },
                new object[] { "Larry", "Cable", "Guy", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.5    ) },
                new object[] { "Sassy", "Molassy", "Texas", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.1    ) },
                new object[] { "Deputy", "John", "Snew", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.005  ) },

                new object[] { "Hodl", "The", "Door", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Hundred), 0.001) },
                new object[] { "To", "The", "Moon", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Hundred), 0.05 ) },
                new object[] { "Separate", "The", "Masses", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Hundred), 0.5  ) },
                new object[] { "Make", "America", "Great", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Hundred), 0.1  ) },
                new object[] { "Engender", "Raging", "Communism", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Hundred), 0.005) },

                new object[] { "Gary", "And", "Demons", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Thousand), 0.001) },
                new object[] { "Dead", "By", "Daylight", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Thousand), 0.05 ) },
                new object[] { "God", "Is", "Good", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Thousand), 0.5  ) },
                new object[] { "I", "Am", "Broken", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Thousand), 0.1  ) },
                new object[] { "Shun", "The", "Nonbelievers", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Per_Thousand), 0.005) },

                new object[] { "Jim", "Crows", "Ass", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.001) },
                new object[] { "Remember", "September", "Novacaine", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.05 ) },
                new object[] { "Fellow", "Grass", "Test", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.5  ) },
                new object[] { "Shimmy", "Shimmy", "Shimmy", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.1  ) },
                new object[] { "Raise", "Hail", "Dale", new LaborData(LetterCodes.GetByLaborCode(LaborCode.Each), 0.005) }
            };

            foreach(var p in parts)
            {
                var pname = string.Join(" - ", p[0], p[1], p[2]);
                var data = p[3] as LaborData;
                var entry = new LaborEntry(pname, data);
                CheckEntries.Add(entry);
            }
        }
	}
}

/* namespace JPMorrow.Test
{
    using JPMorrow.Revit.Labor;

    public static partial class TestBed
    {
        public static TestResult TestLoadLaborFromFile(string settings_path, Document doc, UIDocument uidoc)
        {
            LaborExchange ex = new LaborExchange(settings_path);
            var entries = LaborExchange.LoadLaborFromFile(LaborExchange.DefaultLaborFilePath);
            // debugger.show(header:"TestBed", err:LaborEntry.PrintLaborEntries(entries), max_len:500);

            return new TestResult("Load Labor From File", entries.Any());
        }

        public static TestResult TestLabor(string settings_path, Document doc, UIDocument uidoc)
        {
            LaborExchange ex = new LaborExchange(settings_path);
            ex.GenTwentyTestCases();

            return new TestResult("Labor Exchange", true);
        }
    }
} */