using CommandLine;

namespace PvUpdater
{
    public sealed class Options
    {
        [Option('p', "path", Required = true, HelpText = "Locatia fisierelor")]
        public string Path { get; set; } = string.Empty;

        [Option('o', "output", Required = true, HelpText = "Locatia de iesire a fisierelor")]
        public string OutPut { get; set; } = string.Empty;

        [Option('n', "numar", Required = true, HelpText = "Numarul de start")]
        public int StartNumar { get; set; }

        [Option('d', "data", Required = true, HelpText = "Data")]
        public string Data { get; set; } = string.Empty;
    }
}
