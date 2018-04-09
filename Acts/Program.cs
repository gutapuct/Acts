using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Acts
{
    class Program
    {
        static object _lockObject = new object();

        static void Main(string[] args)
        {
            var data = new ExcelImporter("D:\\temp\\Values.xlsx").GetData();

            var path = "D:\\temp\\Template.docx";

            for (var val = 1; val < data.GetLength(1); val++)
            {
                var dict = new Dictionary<string, string>();

                for (var l = 0; l < data.GetLength(0); l++)
                {
                    dict.Add(data[l, 0], data[l, val]);
                }

                var copypath = $"D:\\temp\\NewDoc_{val}.docx";

                lock (_lockObject)
                {
                    File.Copy(path, copypath);
                }

                using (var doc = WordprocessingDocument.Open(copypath, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    var texts = body.Descendants<Text>();

                    foreach (var pair in dict)
                    {
                        //var tokenTexts = texts.Where(t => String.Equals(t.Text, pair.Key));
                        //foreach (var token in tokenTexts)
                        //{
                        //    var parent = token.Parent;
                        //    var newToken = token.CloneNode(true);
                        //    var lines = Regex.Split(pair.Value, "\r\n|\r|\n");
                        //    ((Text)newToken).Text = lines[0];
                        //    for (int i = 1; i < lines.Length; i++)
                        //    {
                        //        parent.AppendChild<Break>(new Break());
                        //        parent.AppendChild<Text>(new Text(lines[i]));
                        //    }
                        //    token.InsertAfterSelf(newToken);
                        //    token.Remove();
                        //}

                        var tokenTexts = texts.Where(t => t.Text.Contains(pair.Key));
                        foreach (var token in tokenTexts)
                        {
                            token.Text = token.Text.Replace(pair.Key, pair.Value);
                        }
                    }

                    doc.MainDocumentPart.Document.Save();
                }
            }
        }
    }
}
