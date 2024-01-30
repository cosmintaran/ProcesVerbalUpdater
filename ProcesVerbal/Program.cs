using CommandLine;
using Serilog;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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

            var (startNumber, path, output ,data) = ParseArguments(args);
            if (startNumber < 1 && !ConfirmRestart(startNumber))
                return;

            await ProcessFilesAsync(startNumber, data, path, output).ConfigureAwait(false);

            Log.Information("Application is closing....");
            Log.CloseAndFlush();
        }

        private static async Task ProcessFilesAsync(int startNumber, DateTime data, string path, string output)
        {
            string fileExtension = Path.GetExtension(path);

            if (!string.IsNullOrWhiteSpace(fileExtension) && (fileExtension.Equals(".doc") || fileExtension.Equals(".docx")))
            {
                await ExecuteSingleFileAsync(startNumber, data, path, output);
            }
            else if (Directory.Exists(path))
            {
                await ExecuteMultipleFilesAsync(startNumber, data, path, output);
            }
            else
            {
                Log.Error("Invalid file extension for {path}!", path);
                await Task.CompletedTask;
            }
        }
        private static Task ExecuteSingleFileAsync(int startNumber, DateTime data, string path,string output)
        {
            return Task.Run(() =>
            {
                Word.Application wordApp = new Word.Application()
                {
                    Visible = false,
                };
                ProcessSingleFile(ref startNumber, data, path, wordApp, output);
            });
        }

        private static Task ExecuteMultipleFilesAsync(int startNumber, DateTime data, string path, string output)
        {
            return Task.Run(() =>
            {
                Word.Application wordApp = new Word.Application()
                {
                    Visible = false,
                };

                try
                {
                    bool hasOne = false;
                    foreach (var filePath in Directory.EnumerateFiles(path))
                    {
                        if (!filePath.EndsWith(".doc", StringComparison.CurrentCultureIgnoreCase) &&
                        !filePath.EndsWith(".docx", StringComparison.CurrentCultureIgnoreCase)) continue;
                        ProcessSingleFile(ref startNumber, data, filePath, wordApp, output);
                        hasOne = true;
                        if (!hasOne)
                            Log.Warning("No word files fond at location {path}!", path);
                    }
                }
                finally
                {
                    CloseWordObjects(wordApp, null);
                }
            });
        }

        private static void ProcessSingleFile(ref int startNumber, DateTime data, string path, Word.Application wordApp, string output)
        {
            Log.Information("Processing document {document}....", Path.GetFileName(path));
            Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(path, ReadOnly: true);
                var firstParagraph = FindParagraphByText(doc, paragraph);
                if (firstParagraph == null)
                {
                    Log.Warning("Document:{doc} cannot be processed because the format is not as expected.\n" +
                        "e.g.{paragraph}    {textToReplace}!", Path.GetFileName(path), paragraph, textToReplace);
                    return;
                }

                ReadOnlySpan<char> elementToSearch = stackalloc[] { 'N', 'r', '.' };
                ReadOnlySpan<char> dataElem = stackalloc[] { 'd', 'a', 't', 'a' };
                var spanOfParagraph = firstParagraph.Range.Text.AsSpan();
                var indexOf = spanOfParagraph.IndexOf(elementToSearch);
                var slice = spanOfParagraph.Slice(indexOf);
                if(!slice.Contains(dataElem,StringComparison.CurrentCultureIgnoreCase))
                {
                    Log.Warning("Document:{doc} cannot be processed because the format is not as expected.\n" +
                        "e.g.{textToReplace}!", Path.GetFileName(path), textToReplace);
                    return;
                }
                var stringToReplace = slice.ToString();

                //if (!regex.IsMatch(slice.ToString()))
                //{
                //    Log.Warning("Document:{doc} cannot be processed because the format is not as expected.\n" +
                //        "e.g.{textToReplace}", Path.GetFileName(path), textToReplace);
                //    return;
                //}

                StringBuilder stringBuilder = NumberBuilder(ref startNumber, data);

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

        private static (int StartNumber, string Path, string ouput,DateTime Data) ParseArguments(string[] args)
        {
            int startNumber = 1;
            string path = string.Empty;
            string outPut = string.Empty;
            DateTime data = DateTime.Now;

            string[] validFormats = { "MM/dd/yyyy", "dd/MM/yyyy", "d/M/yyyy" };

            try
            {
                Parser.Default.ParseArguments<Options>(args)
                    .WithParsed(o =>
                    {
                        startNumber = o.StartNumar;
                        path = o.Path;
                        outPut = o.OutPut;

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

            return ValueTuple.Create(startNumber, path, outPut, data);
        }

        private static bool ConfirmRestart(int startNumber)
        {
            Log.Warning("The start number is not valid {number}. Do you want to start from 1 ? y/n", startNumber);
            return Console.ReadKey().Key == ConsoleKey.Y;
        }

        private static StringBuilder NumberBuilder(ref int startNumber, DateTime data)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("Nr. ");
            stringBuilder.AppendFormat("{0:D2} ", startNumber++);
            stringBuilder.AppendFormat("data . ");
            stringBuilder.AppendFormat("{0:D2}.", data.Day);
            stringBuilder.AppendFormat("{0:D2}.", data.Month);
            stringBuilder.AppendFormat("{0:D4}", data.Year);
            stringBuilder.Append("    ");
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
    }

}