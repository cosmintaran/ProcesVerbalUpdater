using CommandLine;
using Serilog;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

//using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace PvUpdater
{
    class Program
    {

        private static string paragraph = string.Intern("Biroul de Înregistrare  Sistematică");
        private static string textToReplace = string.Intern("Nr. data.");
        //private static string pattern = @"^Nr\.\s*\d{1,2}\s*data\s*\.\s*\d{2}\.\d{2}\.\d{4}\s*$";
        //private static Regex regex = new Regex(pattern);
        static async Task Main(string[] args)
        {
            SetApplicationCulture();

            ConfigureLogger();
            Log.Information("Application is starting....");
            if (!ValidateArguments(args))
                return;

            var (startNumber, path, output ,data, birou) = ParseArguments(args);
            if (startNumber < 1 && !ConfirmRestart(startNumber))
                return;

            await ProcessFilesAsync(startNumber, data, path, output, birou).ConfigureAwait(false);

            Log.Information("Application is closing....");
            Log.CloseAndFlush();
        }

        private static async Task ProcessFilesAsync(int startNumber, DateTime data, string path, string output, string birou)
        {
            string fileExtension = Path.GetExtension(path);

            if (!string.IsNullOrWhiteSpace(fileExtension) && (fileExtension.Equals(".doc") || fileExtension.Equals(".docx")))
            {
                await ExecuteSingleFileAsync(startNumber, data, path, output, birou);
            }
            else if (Directory.Exists(path))
            {
                await ExecuteMultipleFilesAsync(startNumber, data, path, output, birou);
            }
            else
            {
                Log.Error("Invalid file extension for {path}!", path);
                await Task.CompletedTask;
            }
        }
        private static Task ExecuteSingleFileAsync(int startNumber, DateTime data, string path,string output, string birou)
        {
            return Task.Run(() =>
            {
                Word.Application wordApp = new Word.Application()
                {
                    Visible = false,
                };
                ProcessSingleFile(startNumber, data, path, wordApp, output, birou);
            });
        }

        private static Task ExecuteMultipleFilesAsync(int startNumber, DateTime data, string path, string output, string birou)
        {
            return Task.Run(() =>
            {
                Word.Application wordApp = new Word.Application()
                {
                    Visible = false,
                };

                try
                {

                    var files = Directory.GetFiles(path, "*.*")
                                .Where(file => file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                                               file.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                                .ToArray();

                    if(files.Length == 0)
                    {
                        Log.Error("No word files fond at location {path}!", path);
                    }

                    bool allFilesHaveNumbers = files.All(file =>
                    {
                        var fileName = Path.GetFileNameWithoutExtension(file);
                        return Regex.IsMatch(fileName, @"\d+");
                    });

                    if (allFilesHaveNumbers)
                    {
                        var sortedFiles = files.OrderBy(file =>
                                            {
                                                var match = Regex.Match(Path.GetFileNameWithoutExtension(file), @"\d+");
                                                return match.Success ? int.Parse(match.Value) : int.MaxValue;
                                            }).ToArray();

                        foreach (var filePath in sortedFiles)
                        {
                            var strNr = Path.GetFileNameWithoutExtension(filePath);
                            if(int.TryParse(strNr, out var nr))
                            {
                                ProcessSingleFile(nr, data, filePath, wordApp, output, birou);
                            }
                            else
                            {
                                Log.Error("File name is not a number {path}! Document was skipped!", path);
                            }

                        }

                    }
                    else
                    {
                        foreach (var filePath in files)
                        {
                            ProcessSingleFile(startNumber, data, filePath, wordApp, output, birou);
                            ++startNumber;
                        }
                    }
   
                }
                finally
                {
                    CloseWordObjects(wordApp, null);
                }
            });
        }

        private static void ProcessSingleFile(int startNumber, DateTime data, string path, Word.Application wordApp, string output, string birou)
        {
            Log.Information("Processing document {document}....", Path.GetFileName(path));
            Word.Document doc = null;
            try
            {
                Word.Paragraph firstParagraph = null;
                string stringToReplace = null;

                doc = wordApp.Documents.Open(path, ReadOnly: true);
                if (!string.IsNullOrWhiteSpace(birou) && birou.IndexOf("timis", StringComparison.OrdinalIgnoreCase) > -1)
                {
                    firstParagraph = GetFirstNrPAragraph(doc);
                    if (firstParagraph == null)
                    {
                        Log.Warning("Document:{doc} cannot be processed because the format is not as expected.\n" +
                            "{textToReplace}!", Path.GetFileName(path), textToReplace);
                        return;
                    }

                    stringToReplace = GetStringToReplaceFromParagraph(firstParagraph);
                }
                else
                {
                    firstParagraph = FindParagraphByText(doc, paragraph);
                    if (firstParagraph == null)
                    {
                        Log.Warning("Document:{doc} cannot be processed because the format is not as expected.\n" +
                            "{textToReplace}!", Path.GetFileName(path), textToReplace);
                        return;

                    }

                    stringToReplace = GetStringToReplaceFromParagraph(firstParagraph);
                }

                StringBuilder stringBuilder = NumberBuilder(startNumber, data, birou);
                firstParagraph.Range.Find.Execute(FindText: stringToReplace, ReplaceWith: stringBuilder.ToString(), Replace: Word.WdReplace.wdReplaceOne);

                if (string.IsNullOrEmpty(output))
                {
                    output = Path.Combine(Path.GetDirectoryName(path), "Output");
                }

                if (!Directory.Exists(output))
                    Directory.CreateDirectory(output);

                string pathToSave = $"{Path.Combine(output, Path.GetFileNameWithoutExtension(path))}.docx";

                doc.SaveAs2(pathToSave, Word.WdSaveFormat.wdFormatXMLDocument);

                Log.Information("Document:{doc} processed!", Path.GetFileName(path));
            }
            catch (Exception ex)
            {
                LogAndHandleException(path, ex);
            }
            finally
            {
                CloseWordObjects(null, doc);
            }
        }

        private static string GetStringToReplaceFromParagraph(Word.Paragraph firstParagraph)
        {
            ReadOnlySpan<char> elementToSearch = stackalloc[] { 'N', 'r', '.' };
            ReadOnlySpan<char> dataElem = stackalloc[] { 'd', 'a', 't', 'a' };
            var spanOfParagraph = firstParagraph.Range.Text.AsSpan();
            var indexOf = spanOfParagraph.IndexOf(elementToSearch);
            var slice = spanOfParagraph.Slice(indexOf);
            //if(!slice.Contains(dataElem,StringComparison.CurrentCultureIgnoreCase))
            //{
            //    Log.Warning("Document:{doc} cannot be processed because the format is not as expected.\n" +
            //        "e.g.{textToReplace}!", Path.GetFileName(path), textToReplace);
            //    return;
            //}
            var stringToReplace = slice.Trim().ToString();
            return stringToReplace;
        }

        private static void SetApplicationCulture()
        {
            CultureInfo romanianCulture = new CultureInfo("ro-RO");
            CultureInfo.DefaultThreadCurrentCulture = romanianCulture;
            CultureInfo.DefaultThreadCurrentUICulture = romanianCulture;
        }

        private static void ConfigureLogger()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console(Serilog.Events.LogEventLevel.Verbose, outputTemplate:
                "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File("logs.log", rollingInterval: RollingInterval.Hour)
                .CreateLogger();
        }

        private static bool ValidateArguments(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Log.Error("No arguments provided.");
                return false;
            }

            return true;
        }

        private static (int StartNumber, string Path, string ouput,DateTime Data, string birouOCPI) ParseArguments(string[] args)
        {
            int startNumber = 1;
            string path = string.Empty;
            string outPut = string.Empty;
            DateTime data = DateTime.Now;
            string birou = string.Empty;
            
            string[] validFormats = { "MM/dd/yyyy", "dd/MM/yyyy", "d/M/yyyy" };

            try
            {
                Parser.Default.ParseArguments<Options>(args)
                    .WithParsed(o =>
                    {
                        startNumber = o.StartNumar;
                        path = o.Path;
                        outPut = o.OutPut;
                        birou = string.IsNullOrWhiteSpace(o.OCPI) ? "Timis" : o.OCPI;

                        if (!DateTime.TryParseExact(o.Data, validFormats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out data))
                        {
                            Log.Error(@"Invalid Date format. Please use one of the following formats:
                                    ""MM/dd/yyyy"", ""dd/MM/yyyy"", ""d/M/yyyy""");
                        }
                    });
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "Fatal error occurred: {ErrorMessage}!", ex.Message);
            }

            return ValueTuple.Create(startNumber, path, outPut, data, birou);
        }

        private static bool ConfirmRestart(int startNumber)
        {
            Log.Warning("The start number is not valid {number}. Do you want to start from 1 ? y/n", startNumber);
            return Console.ReadKey().Key == ConsoleKey.Y;
        }

        private static StringBuilder NumberBuilder(int startNumber, DateTime data, string ocpi)
        {
            StringBuilder stringBuilder = new StringBuilder();
            if (string.IsNullOrWhiteSpace(ocpi) || ocpi.StartsWith("Caras"))
            {
                stringBuilder.Append("Nr. ");
                stringBuilder.AppendFormat("{0:D2} ", startNumber);
                stringBuilder.AppendFormat("data . ");
                stringBuilder.AppendFormat("{0:D2}.", data.Day);
                stringBuilder.AppendFormat("{0:D2}.", data.Month);
                stringBuilder.AppendFormat("{0:D4}", data.Year);
                stringBuilder.Append("    ");
            }
            if (ocpi.StartsWith("Timis"))
            {
                stringBuilder.Append("Nr. ");
                stringBuilder.AppendFormat("{0:D2} ", startNumber);
                stringBuilder.Append('/');
                stringBuilder.AppendFormat("{0:D2}.", data.Day);
                stringBuilder.AppendFormat("{0:D2}.", data.Month);
                stringBuilder.AppendFormat("{0:D4}", data.Year);
                stringBuilder.Append(" ");
            }
            return stringBuilder;
        }

        static Word.Paragraph FindParagraphByText(Word.Document doc, string searchText)
        {
            foreach (Word.Paragraph paragraph in doc.Paragraphs)
            {
                if (paragraph.Range.Text.Contains(searchText))
                {
                    return paragraph;
                }
            }

            return null; // Textul nu a fost găsit în niciun paragraf
        }

        private static void CloseWordObjects(Word.Application wordApp, Word.Document doc)
        {
            if (doc != null)
            {
                doc.Close();
                Marshal.ReleaseComObject(doc);
            }

            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }
        }

        private static void LogAndHandleException(string filePath, Exception ex)
        {
            Log.Error("Error while processing file {file}. Exception:{exception}!", Path.GetFileName(filePath), ex.Message);
        }

        static Word.Paragraph FindParagraphByPattern(Word.Document doc, string pattern)
        {
            foreach (Word.Paragraph paragraph in doc.Paragraphs)
            {
                if (Regex.IsMatch(paragraph.Range.Text, pattern, RegexOptions.IgnoreCase))
                {
                    return paragraph;
                }
            }

            return null;
        }

        private static Word.Paragraph GetFirstNrPAragraph(Word.Document doc)
        {
            foreach (Word.Paragraph paragraph in doc.Paragraphs)
            {
                string text = paragraph.Range.Text.Trim();
                if (text.StartsWith("Nr.", StringComparison.InvariantCultureIgnoreCase))
                {
                    return paragraph;
                }
            }
            return null;
        }

    }

}