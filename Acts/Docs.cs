using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Acts
{
    public class Docs
    {
        private readonly object LockObject;
        private readonly string PathToTemplate;
        private readonly string PathToValues;
        private readonly string PathToReasons;

        public Docs(string pathToTemplate, string pathToValues, string pathToReasons)
        {
            LockObject = new object();
            PathToTemplate = pathToTemplate;
            PathToValues = pathToValues;
            PathToReasons = pathToReasons;
        }

        public void Execute()
        {
            try
            {
                var randomFolder = Guid.NewGuid().ToString();
                Console.Write("Creating acts...");
                CreateAllActs(randomFolder);
                Console.WriteLine("Done.");
                Console.WriteLine();

                Console.WriteLine("Getting new all files...");
                var files = GetStreamAllFiles(randomFolder);
                Console.WriteLine("Done.");
                Console.WriteLine();

                Console.Write("Creating the new general file...");
                SaveNewFile(files);
                Console.WriteLine("Done.");
                Console.WriteLine();

                Console.WriteLine("Deleting temp files...");
                RemoveFiles(randomFolder);
                Console.WriteLine("Done. Press any key to exit.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
#if (DEBUG)
                Console.WriteLine(ex);
#else
                Console.WriteLine(ex.Message); 
#endif
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }

        private byte[] OpenAndCombine(IList<byte[]> documents)
        {
            MemoryStream mainStream = new MemoryStream();

            mainStream.Write(documents[0], 0, documents[0].Length);
            mainStream.Position = 0;

            int pointer = 1;
            byte[] ret;
            try
            {
                using (WordprocessingDocument mainDocument = WordprocessingDocument.Open(mainStream, true))
                {

                    XElement newBody = XElement.Parse(mainDocument.MainDocumentPart.Document.Body.OuterXml);

                    for (pointer = 1; pointer < documents.Count; pointer++)
                    {
                        if (pointer % 5 == 0)
                        {
                            Console.Write(".");
                        }

                        WordprocessingDocument tempDocument = WordprocessingDocument.Open(new MemoryStream(documents[pointer]), true);
                        XElement tempBody = XElement.Parse(tempDocument.MainDocumentPart.Document.Body.OuterXml);

                        newBody.Add(tempBody);
                        mainDocument.MainDocumentPart.Document.Body = new Body(newBody.ToString());
                        mainDocument.MainDocumentPart.Document.Save();
                        mainDocument.Package.Flush();
                    }
                }
            }
            catch (OpenXmlPackageException oxmle)
            {
                Console.WriteLine($"Error while merging files: Document index {0}. {oxmle.Message}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error while merging files: Document index {0}. {e.Message}");
            }
            finally
            {
                ret = mainStream.ToArray();
                mainStream.Close();
                mainStream.Dispose();
            }
            return (ret);
        }

        private void CreateAllActs(string randomFolder)
        {
            var values = new ExcelImporter(PathToValues).GetData();
            var reasons = new ExcelImporter(PathToReasons).GetData().ToReasonModel();

            for (var value = 1; value < values.GetLength(1); value++)
            {
                if (value % 5 == 0)
                {
                    Console.Write(".");
                }

                System.Threading.Thread.Sleep(10);
                var date = DateTime.Now.ToString("yyyyMMddhhmmssffff");
                var copyPath = $"{Path.GetDirectoryName(PathToTemplate)}\\{randomFolder}\\{date}_{value}.docx";

                lock (LockObject)
                {
                    if (!Directory.Exists($"{Path.GetDirectoryName(PathToTemplate)}\\{randomFolder})"));
                    {
                        Directory.CreateDirectory($"{Path.GetDirectoryName(PathToTemplate)}\\{randomFolder}");
                    }
                    File.Copy(PathToTemplate, copyPath);
                }

                var pairsForReplacings = new Dictionary<string, string>();

                // Add name of columns in the dictionary
                for (var d = 0; d < values.GetLength(0); d++)
                {
                    if (pairsForReplacings.ContainsKey(values[d, 0]))
                    {
                        throw new Exception("Some values consist more than one columns. Also, you can't use the same values in the different columns.");
                    }
                    pairsForReplacings.Add(values[d, 0], values[d, value]);
                }

                if (pairsForReplacings.ContainsKey("reason"))
                {
                    throw new Exception("Your Values can't contain \"reason\" column!");
                }

                if (pairsForReplacings.ContainsKey("Name") || pairsForReplacings.ContainsKey("name"))
                {
                    var equipmentName = pairsForReplacings.ContainsKey("Name") ? pairsForReplacings["Name"] : pairsForReplacings["name"];
                    pairsForReplacings.Add("reason", reasons.GetReasonByEquipmentName(equipmentName));
                }
                else
                {
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    Console.WriteLine("We couldn't find the column named \"Name\" or \"name\". The reasins will not be replaced!");
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }

                using (var doc = WordprocessingDocument.Open(copyPath, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    var texts = body.Descendants<Text>();

                    foreach (var pair in pairsForReplacings.Where(i => !String.IsNullOrWhiteSpace(i.Key)))
                    {
                        var tokenTexts = texts.Where(t => t.Text.Contains(pair.Key));
                        foreach (var token in tokenTexts)
                        {
                            token.Text = token.Text.Replace(pair.Key, pair.Value);
                        }
                    }

                    if (texts.Any())
                    {
                        Paragraph PageBreakParagraph = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                        doc.MainDocumentPart.Document.Body.Append(PageBreakParagraph);
                    }

                    doc.MainDocumentPart.Document.Save();
                }
            }
            Console.WriteLine($"Acts created: {values.GetLength(1)}");
        }

        private IList<byte[]> GetStreamAllFiles(string randomFolder)
        {
            var fileList = new List<byte[]>();
            var files = Directory.GetFiles($"{Path.GetDirectoryName(PathToTemplate)}\\{randomFolder}");

            foreach (var file in files)
            {
                byte[] textByteArray = File.ReadAllBytes(file);
                fileList.Add(textByteArray);
            }

            return fileList;
        }

        private void SaveNewFile(IList<byte[]> files)
        {
            var result = this.OpenAndCombine(files);

            using (var newFile = File.Create($"{Path.GetDirectoryName(PathToTemplate)}\\Acts.docx")) { }
            File.WriteAllBytes($"{Path.GetDirectoryName(PathToTemplate)}\\Acts.docx", result);
        }

        private void RemoveFiles(string randomFolder)
        {
            var files = Directory.GetFiles($"{Path.GetDirectoryName(PathToTemplate)}\\{randomFolder}");

            foreach (var file in files)
            {
                File.Delete(file);
            }

            Directory.Delete($"{Path.GetDirectoryName(PathToTemplate)}\\{randomFolder}");
        }
    }
}