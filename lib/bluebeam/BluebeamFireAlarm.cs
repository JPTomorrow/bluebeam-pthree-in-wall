using System;
using System.Collections.Generic;
using System.Linq;
using JPMorrow.Measurements;
using PdfSharp.Pdf.Annotations;

namespace JPMorrow.Pdf.Bluebeam.FireAlarm
{   
    /// <summary>
    /// takes poly lines representing conduit in bluebeam and extracts the total length of them
    /// </summary>
    public class ConduitLengthTotal
    {
        private Dictionary<string, List<double>> TotalEmtLength { get; set; } = new Dictionary<string, List<double>>()
        {
            {"1/2\"", new List<double>()},
            {"3/4\"", new List<double>()},
            {"1\"", new List<double>()},
            {"1 1/2\"", new List<double>()},
        };

        private Dictionary<string, List<double>> TotalPvcLength { get; set; } = new Dictionary<string, List<double>>() 
        {
            {"1/2\"", new List<double>()},
            {"3/4\"", new List<double>()},
            {"1\"", new List<double>()},
            {"1 1/2\"", new List<double>()},
        };

        private Dictionary<string, List<double>> TotalMcCableLength { get; set; } = new Dictionary<string, List<double>>() 
        {
            {"1/2\"", new List<double>()},
            {"3/4\"", new List<double>()},
            {"1\"", new List<double>()},
            {"1 1/2\"", new List<double>()},
        };

        public string GetTotalEmtLengthFeetIn(string conduit_size) 
        {
            bool s = TotalEmtLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalEmtLength in feet and inches");
            return Measure.LengthFromDbl(doubles.Sum());
        }

        public string GetTotalPvcLengthFeetIn(string conduit_size)
        {
            bool s = TotalPvcLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalPvcLength in feet and inches");
            return Measure.LengthFromDbl(doubles.Sum());
        }

        public string GetTotalMcCableLengthFeetIn(string conduit_size)
        {
            bool s = TotalMcCableLength.TryGetValue(conduit_size, out var doubles);
            if(!s) throw new Exception("Could not retrieve TotalMcCableLength in feet and inches");
            return Measure.LengthFromDbl(doubles.Sum());
        }

        private static string emtSubject = "conduit - emt - fire alarm";
        private static string pvcSubject = "conduit - pvc - fire alarm";
        private static string mcSubject = "conduit - mc - fire alarm";

        public ConduitLengthTotal(IEnumerable<PdfAnnotation> annotations)
        {
            foreach(var a in annotations)
            {
                if(a == null) throw new Exception("The annotation provided is null");
                var sub_type = GetSubType(a);
                var subject = GetSubject(a).ToLower();
                var content = GetContents(a);

                if(subject.Contains(emtSubject))
                {
                    
                }
                else if(subject.Contains(pvcSubject))
                {

                }
                else if(subject.Contains(mcSubject))
                {

                }
            }
        }

        private static string GetSubType(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_sub_type = elements.TryGetString("/SubType", out string sub_type);

            if(!has_sub_type || !sub_type.ToLower().Equals("/PolyLine"))
                throw new Exception("The provided annotation is not a poly line");

            return sub_type;
        }

        private static string GetSubject(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Subj", out string subject);

            if(!has_subj_name)
                throw new Exception("The provided annotation does not have the correct subject");

            return subject;
        }

        private static string GetContents(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Contents", out string contents);

            if(!has_subj_name)
                throw new Exception("The provided annotation does not have any contents");

            return contents;
        }

        private static string GetConduitSizeFromContent(string contents)
        {

        }
    }
}