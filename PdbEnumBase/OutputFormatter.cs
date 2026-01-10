using System.IO;
using System.Runtime.Serialization.Json;
using System.Xml.Serialization;

namespace PdbEnum
{
    public enum OutputFormat
    {
        Human,
        Json,
        Xml
    }

    internal static class OutputFormatter
    {
        public static void WriteResult(SymbolSearchResult result, OutputFormat format, TextWriter writer)
        {
            switch (format)
            {
                case OutputFormat.Json:
                    WriteJson(result, writer);
                    break;
                case OutputFormat.Xml:
                    WriteXml(result, writer);
                    break;
                case OutputFormat.Human:
                default:
                    WriteHuman(result, writer);
                    break;
            }
        }

        private static void WriteJson(SymbolSearchResult result, TextWriter writer)
        {
            DataContractJsonSerializer serializer = new(typeof(SymbolSearchResult));
            using MemoryStream memoryStream = new();
            {
                serializer.WriteObject(memoryStream, result);
                memoryStream.Position = 0;
                using StreamReader reader = new(memoryStream);
                {
                    writer.Write(reader.ReadToEnd());
                }
            }
        }

        private static void WriteXml(SymbolSearchResult result, TextWriter writer)
        {
            XmlSerializer serializer = new(typeof(SymbolSearchResult));
            serializer.Serialize(writer, result);
        }

        private static void WriteHuman(SymbolSearchResult result, TextWriter writer)
        {
            if (!result.Success)
            {
                writer.WriteLine($"Error: {result.ErrorMessage}");
                return;
            }

            if (result.Module != null)
            {
                writer.WriteLine(result.Module.ToString());
                writer.WriteLine();
            }

            if (result.PdbInfo != null)
            {
                writer.WriteLine(result.PdbInfo.ToString());
                writer.WriteLine();
            }

            if (result.Symbol != null)
            {
                writer.WriteLine("Symbol found!");
                writer.WriteLine(result.Symbol.ToString());
            }
            else
            {
                writer.WriteLine("Symbol not found.");
            }
        }

        public static void WriteBatchResult(BatchSymbolSearchResult result, OutputFormat format, TextWriter writer)
        {
            switch (format)
            {
                case OutputFormat.Json:
                    WriteBatchJson(result, writer);
                    break;
                case OutputFormat.Xml:
                    WriteBatchXml(result, writer);
                    break;
                case OutputFormat.Human:
                default:
                    WriteBatchHuman(result, writer);
                    break;
            }
        }

        private static void WriteBatchJson(BatchSymbolSearchResult result, TextWriter writer)
        {
            DataContractJsonSerializer serializer = new(typeof(BatchSymbolSearchResult));
            using MemoryStream memoryStream = new();
            {
                serializer.WriteObject(memoryStream, result);
                memoryStream.Position = 0;
                using StreamReader reader = new(memoryStream);
                {
                    writer.Write(reader.ReadToEnd());
                }
            }
        }

        private static void WriteBatchXml(BatchSymbolSearchResult result, TextWriter writer)
        {
            XmlSerializer serializer = new(typeof(BatchSymbolSearchResult));
            serializer.Serialize(writer, result);
        }

        private static void WriteBatchHuman(BatchSymbolSearchResult result, TextWriter writer)
        {
            if (!result.Success)
            {
                writer.WriteLine($"Error: {result.ErrorMessage}");
                return;
            }

            if (result.Module != null)
            {
                writer.WriteLine(result.Module.ToString());
                writer.WriteLine();
            }

            if (result.PdbInfo != null)
            {
                writer.WriteLine(result.PdbInfo.ToString());
                writer.WriteLine();
            }

            if (result.Symbols != null && result.Symbols.Count > 0)
            {
                writer.WriteLine($"Symbol Search Results ({result.Symbols.Count} symbols):");
                writer.WriteLine();

                foreach (SymbolSearchResult symbolResult in result.Symbols)
                {
                    writer.WriteLine($"Searched for: {symbolResult.SearchedSymbolName}");
                    if (symbolResult.Symbol != null)
                    {
                        writer.WriteLine("  Found!");
                        writer.WriteLine($"  {symbolResult.Symbol}");
                    }
                    else
                    {
                        writer.WriteLine("  Not found.");
                    }
                    writer.WriteLine();
                }
            }
        }
    }
}