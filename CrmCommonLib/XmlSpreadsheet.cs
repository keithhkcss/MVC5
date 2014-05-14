using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows.Forms;

// DotNetZip
using Ionic.Zip;

namespace CrmCommonLib
{
    public class XmlSpreadsheet
    {
        public string TemplateName { get; set; }
        public string FilePath { get; set; }
        public string Body { get; set; }

        public XmlSpreadsheet(string TemplateName, string FilePath)
        {
            this.TemplateName = TemplateName;
            this.FilePath = FilePath;
        }

        public void load()
        {
            this.Body = System.IO.File.ReadAllText(this.FilePath + "/" + this.TemplateName);

            if(Body==null) Body = "{%DUMMY%}<Row>{%REPEAT-START:DUMMY%}</Row> <Row>{%REPEAT-END:DUMMY%}</Row>";
       
            /*
            bool cont = true;

            while(cont)
            {
                FindRepeatMergeResult r = FindRepeatMerge(Body);
                if(r==null)
                    cont = false;
                else
                {
                    int count = OnRepeatCount(r.repeatName);
                    string s = "";
                    for(int i=0; i < count; i++)
                    {
                        string s1 = r.repeatBody;
                        bool cont1 = true;
                        while(cont1)
                        {
                            FindMergeResult r1 = FindMerge(s1);
                            if(r1==null)
                                cont1 = false;
                            else
                            {
                                string s2 = OnRepeatMergeField(r1.mergeName, r.repeatName, i);
                                if(s2==null) s2 = "";
                                s1 = s1.Substring(0, r1.istart) + s2 + s1.Substring(r1.iend);
                            }
                        }
                        s += s1;
                    }
                    Body = Body.Substring(0, r.istart) + s + Body.Substring(r.iend);
                }
            }
        
            cont = true;
            while(cont)
            {
                FindMergeResult r = FindMerge(Body);
                if(r==null)
                    cont = false;
                else
                {
                    string s = OnMergeField(r.mergeName);
                    if(s==null) s = "";
                    Body = Body.Substring(0, r.istart) + s + Body.Substring(r.iend);
                }
            }
        
            //workaround.  this property may cause an error when opening in Excel, 
            // if the row count specified in template doesn't match the actual row count.
            Body = Body.Replace("ss:ExpandedRowCount=", "ss:ExpandedRowCountO=");
            */
        }

        public FindRepeatMergeResult GetRepeatBody_Neated(string repeatName, string parentRepeatBody)
        {
            string repeatStartTag = "{%REPEAT-START:" + repeatName + "%}";

            int p1 = parentRepeatBody.IndexOf(repeatStartTag);
            int p2 = parentRepeatBody.IndexOf("%}", p1);

            string endTag = "{%REPEAT-END:" + repeatName + "%}";
            int p3 = parentRepeatBody.IndexOf(endTag, p2);

            Range r1 = GetRowRange(parentRepeatBody, p1);
            Range r2 = GetRowRange(parentRepeatBody, p3);

            string repeatBody = parentRepeatBody.Substring(r1.iend, r2.istart - r1.iend);

            return new FindRepeatMergeResult(r1.istart, r2.iend, repeatName, repeatBody);
            
        }

        public FindRepeatMergeResult GetRepeatBody(string repeatName)
        {
            string repeatStartTag = "{%REPEAT-START:" + repeatName + "%}";

            int p1 = Body.IndexOf(repeatStartTag);
            int p2 = Body.IndexOf("%}", p1);

            string endTag = "{%REPEAT-END:" + repeatName + "%}";
            int p3 = Body.IndexOf(endTag, p2);

            Range r1 = GetRowRange(Body, p1);
            Range r2 = GetRowRange(Body, p3);

            string repeatBody = Body.Substring(r1.iend, r2.istart - r1.iend);

            return new FindRepeatMergeResult(r1.istart, r2.iend, repeatName, repeatBody);
        }

        public FindRepeatMergeResult GetRepeatBody_Neated_Col(string repeatName, string parentRepeatBody)
        {
            string repeatStartTag = "{%REPEAT-START:" + repeatName + "%}";

            int p1 = parentRepeatBody.IndexOf(repeatStartTag);
            int p2 = parentRepeatBody.IndexOf("%}", p1);

            string endTag = "{%REPEAT-END:" + repeatName + "%}";
            int p3 = parentRepeatBody.IndexOf(endTag, p2);

            Range r1 = GetColRange(parentRepeatBody, p1);
            Range r2 = GetColRange(parentRepeatBody, p3);

            string repeatBody = parentRepeatBody.Substring(r1.iend, r2.istart - r1.iend);

            return new FindRepeatMergeResult(r1.istart, r2.iend, repeatName, repeatBody);
        }

        public FindRepeatMergeResult GetRepeatBody_Col(string repeatName)
        {
            string repeatStartTag = "{%REPEAT-START:" + repeatName + "%}";

            int p1 = Body.IndexOf(repeatStartTag);
            int p2 = Body.IndexOf("%}", p1);

            string endTag = "{%REPEAT-END:" + repeatName + "%}";
            int p3 = Body.IndexOf(endTag, p2);

            Range r1 = GetColRange(Body, p1);
            Range r2 = GetColRange(Body, p3);

            string repeatBody = Body.Substring(r1.iend, r2.istart - r1.iend);

            return new FindRepeatMergeResult(r1.istart, r2.iend, repeatName, repeatBody);
        }

        public string GetRepeatHeaderCell_Col(string repeatName)
        {
            string repeatStartTag = "{%REPEAT-START:" + repeatName + "%}";
            int p1 = Body.IndexOf(repeatStartTag);

            Range r1 = GetColRange(Body, p1);

            string repeatHeaderCell = Body.Substring(r1.istart, r1.iend - r1.istart);

            return repeatHeaderCell;
        }

        public string GetFirstCell_ColRepeater(string repeatName)
        {
            string repeatStartTag = "{%REPEAT-START:" + repeatName + "%}";
            string repeatHeaderCell = GetRepeatHeaderCell_Col(repeatName);

            return repeatHeaderCell.Replace(repeatStartTag, "");
        }

        public string GetRepeatBody2(string repeatName)
        {
            string s = "";

            bool cont = true;

            while (cont)
            {
                FindRepeatMergeResult r = FindRepeatMerge(Body);
                if (r == null)
                    cont = false;
                else
                {
                    if (r.repeatName == repeatName)
                    {
                        s = r.repeatBody;
                    }
                    //int count = OnRepeatCount(r.repeatName);
                    //string s = "";
                    //for (int i = 0; i < count; i++)
                    //{
                    //    string s1 = r.repeatBody;
                    //    bool cont1 = true;
                    //    while (cont1)
                    //    {
                    //        FindMergeResult r1 = FindMerge(s1);
                    //        if (r1 == null)
                    //            cont1 = false;
                    //        else
                    //        {
                    //            string s2 = OnRepeatMergeField(r1.mergeName, r.repeatName, i);
                    //            if (s2 == null) s2 = "";
                    //            s1 = s1.Substring(0, r1.istart) + s2 + s1.Substring(r1.iend);
                    //        }
                    //    }
                    //    s += s1;
                    //}
                    //Body = Body.Substring(0, r.istart) + s + Body.Substring(r.iend);
                }
            }

            return s;
        }        

        public string MergeField_Repeater(string repeatBody, string templateFieldName, string outputVal)
        {
            string mergeField;

            decimal n;
            bool isNumeric = Decimal.TryParse(outputVal, out n);

            if (!isNumeric)
            {
                mergeField = "{%" + templateFieldName + "%}";

                int p1 = repeatBody.IndexOf(mergeField);
                int p2 = p1 + 2 + templateFieldName.Length;

                repeatBody = repeatBody.Substring(0, p1) + outputVal + repeatBody.Substring(p2 + 2);
            }
            else
            {
                mergeField = "<Data ss:Type=\"String\">{%" + templateFieldName + "%}";

                int p1 = repeatBody.IndexOf(mergeField);
                int p2 = p1 + mergeField.Length;

                repeatBody = repeatBody.Substring(0, p1) + "<Data ss:Type=\"Number\">" + outputVal + repeatBody.Substring(p2);
            }

            return repeatBody;
        }

        public void MergeField(string templateFieldName, string outputVal)
        {
            string mergeField;

            decimal n;
            bool isNumeric = Decimal.TryParse(outputVal, out n);

            if (!isNumeric)
            {
                mergeField = "{%" + templateFieldName + "%}";

                int p1 = Body.IndexOf(mergeField);
                int p2 = p1 + 2 + templateFieldName.Length;

                Body = Body.Substring(0, p1) + outputVal + Body.Substring(p2 + 2);
            }
            else
            {
                mergeField = "<Data ss:Type=\"String\">{%" + templateFieldName + "%}";

                int p1 = Body.IndexOf(mergeField);
                int p2 = p1 + mergeField.Length;

                Body = Body.Substring(0, p1) + "<Data ss:Type=\"Number\">" + outputVal + Body.Substring(p2);
            }
        }

        public void XmlEnding()
        {
            Body = Body.Replace("ss:ExpandedRowCount=", "ss:ExpandedRowCountO=");
            Body = Body.Replace("ss:ExpandedColumnCount=", "ss:ExpandedColumnCountO=");
        }

        public void SaveXml(string filePath, string fileName, string zipPassword)
        {
            //XmlDocument xdoc = new XmlDocument();
            //xdoc.LoadXml(this.Body);
            //xdoc.Save(filePath + "/" + fileName);

            using (ZipFile zip = new ZipFile())
            {
                zip.Password = zipPassword;
                //WriteDelegate
                //w.
                zip.AddEntry(fileName + ".xml", this.Body.ToString(), Encoding.UTF8);
                zip.Save(filePath + "/" + fileName + ".zip");
            }
        }


        private Range GetRowRange(string s, int index)
        {
            Range r = new Range(-1,-1);
            r.istart = s.Substring(0, index).LastIndexOf("<Row");
            r.iend = s.IndexOf("</Row>", index) + 6;
            return r;
        }

        private Range GetColRange(string s, int index)
        {
            Range r = new Range(-1, -1);
            r.istart = s.Substring(0, index).LastIndexOf("<Cell");
            r.iend = s.IndexOf("</Cell>", index) + 7;
            return r;
        }

        private FindRepeatMergeResult FindRepeatMerge(string s)
        {
            int p1 = s.IndexOf("{%REPEAT-START:");
            if(p1<0) return null;
        
            int p2 = s.IndexOf("%}", p1);
            if(p2<0) return null;

            string mergeName = s.Substring(p1 + 2, p2);
            string repeatName = mergeName.Substring(13);
        
            string endTag = "{%REPEAT-END:" + repeatName + "%}";
            int p3 = s.IndexOf(endTag, p2);
            if(p3<0) return null;

            Range r1 = GetRowRange(s, p1);
            Range r2 = GetRowRange(s, p3);

            string repeatBody = s.Substring(r1.iend, r2.istart);
        
            return new FindRepeatMergeResult(r1.istart, r2.iend, repeatName, repeatBody);
        } 



        private FindMergeResult FindMerge(string s)
        {
            int p1 = s.IndexOf("{%");
            if(p1<0) return null;
        
            int p2 = s.IndexOf("%}", p1);
            if(p2<0) return null;

            string mergeName = s.Substring(p1 + 2, p2 - (p1 + 2));
        
            return new FindMergeResult(p1, p2+2, mergeName);
        } 
    
        virtual protected string OnMergeField(string mergeName)
        {
            return "";
        }
    
        virtual protected int OnRepeatCount(string repeatName)
        {
            return 1;
        }
    
        virtual protected string OnRepeatMergeField(string mergeName, string repeatName, int index)
        {
            return "";
        }

        
        private class FindMergeResult
        {
            public int istart;
            public int iend;
            public string mergeName;
            public FindMergeResult(int istart, int iend, string mergeName)
            {
                this.istart = istart;
                this.iend = iend;
                this.mergeName = mergeName;
            }
        }
    
        public class FindRepeatMergeResult
        {
            public int istart;
            public int iend;
            public string repeatName;
            public string repeatBody;
            public FindRepeatMergeResult(int istart, int iend, string repeatName, string repeatBody)
            {
                this.istart = istart;
                this.iend = iend;
                this.repeatName = repeatName;
                this.repeatBody = repeatBody;
            }
        }
    
        private class Range
        {
            public int istart;
            public int iend;
            public Range(int istart, int iend)
            {
                this.istart = istart;
                this.iend = iend;
            }
        }

    }
}
