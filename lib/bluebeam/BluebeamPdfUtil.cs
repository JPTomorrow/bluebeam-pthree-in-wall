
using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;

namespace JPMorrow.Pdf.Bluebeam
{
    public static class BluebeamPdfUtil
    {
        public static string[] FireAlarmConduitSubjectTags = new string[] 
        {
            "emt - conduit run - supported",
            "pvc - conduit run - supported",
            "mc - conduit run - supported"
        };

        public static string[] FireAlarmConduitSizes = new string[] { "1/2\"", "3/4\"", "1\"", "1 1/4\"", "1 1/2\"", "2\"" };

        public static bool IsPolyLine(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_sub_type = elements.TryGetString("/Subtype", out string sub_type);
            if(!has_sub_type || !sub_type.Equals("/PolyLine")) return false;
            return true;
        }

        /// <summary>
        /// Tell whether the provided pdf annotation is of subtype /Rectangle
        /// </summary>
        public static bool IsRectangle(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool s = elements.TryGetString("/Subtype", out string sub_type);
            return s && sub_type.Equals("/Square");
        }

        public static string GetRcContents(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/RC", out string rc);

            if(!has_subj_name)
                throw new Exception("The provided annotation does not have any /RC data");

            return rc;
        }

        public static string GetSubject(PdfAnnotation a)
        {
            var elements = a.Elements;
            bool has_subj_name = elements.TryGetString("/Subj", out string subject);
            
            if(!has_subj_name)
                 throw new Exception("The provided annotation has no subject");

            return subject;
        }

        public static bool HasGroupNesting(PdfAnnotation a, out List<string> ret_group_codes)
        {
            ret_group_codes = new List<string>();
            var elements = a.Elements;
            bool s = elements.TryGetValue("/GroupNesting", out PdfItem item);

            if (!s) return false;

            var groups = item.ToString()
                .Split(" ").Where(x => x.StartsWith('/'))
                .Select(x => x.Substring(1)).ToList();

            ret_group_codes = groups;
            return groups.Any();
        }

        public static IEnumerable<PdfAnnotation> FindAnnotationsByGroupCode(
            string group_code, IEnumerable<PdfAnnotation> search)
        {
            List<PdfAnnotation> ret = new List<PdfAnnotation>();

            foreach(var a in search)
            {
                var elements = a.Elements;
                bool s = elements.TryGetString("/NM", out string cmp_code);
                if(!s || string.IsNullOrWhiteSpace(cmp_code)) continue;
                if(cmp_code.Equals(group_code)) ret.Add(a);
            }

            return ret;
        }
    }
}