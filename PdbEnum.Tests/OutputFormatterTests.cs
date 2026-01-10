using NUnit.Framework;
using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml.Serialization;
using PdbEnum;

namespace PdbEnum.Tests
{
    [TestFixture]
    public class OutputFormatterTests
    {
        private SymbolSearchResult CreateTestResult()
        {
            return new SymbolSearchResult
            {
                Success = true,
                Module = new ModuleInfo
                {
                    Name = "test.dll",
                    FullPath = "C:\\Windows\\System32\\test.dll",
                    BaseAddress = 0x12340000,
                    Size = 0x10000,
                    EntryPoint = 0x12341000
                },
                PdbInfo = new PdbInfo
                {
                    PdbFileName = "test.pdb"
                },
                Symbol = new SymbolInfo
                {
                    Name = "TestFunction",
                    Address = 0x12345678,
                    Size = 0x100
                }
            };
        }

        [Test]
        public void Test_WriteResult_Human_WithSuccess()
        {
            SymbolSearchResult result = CreateTestResult();
            StringWriter writer = new StringWriter();

            OutputFormatter.WriteResult(result, OutputFormat.Human, writer);

            string output = writer.ToString();
            Assert.IsNotEmpty(output, "Human output should not be empty");
            Assert.IsTrue(output.Contains("test.dll"), "Output should contain module name");
            Assert.IsTrue(output.Contains("TestFunction"), "Output should contain symbol name");
        }

        [Test]
        public void Test_WriteResult_Human_WithError()
        {
            SymbolSearchResult result = new SymbolSearchResult
            {
                Success = false,
                ErrorMessage = "Test error message"
            };
            StringWriter writer = new StringWriter();

            OutputFormatter.WriteResult(result, OutputFormat.Human, writer);

            string output = writer.ToString();
            Assert.IsTrue(output.Contains("Error"), "Error output should contain 'Error'");
            Assert.IsTrue(output.Contains("Test error message"), "Output should contain error message");
        }

        [Test]
        public void Test_WriteResult_Json_CanSerialize()
        {
            SymbolSearchResult result = CreateTestResult();
            StringWriter writer = new StringWriter();

            OutputFormatter.WriteResult(result, OutputFormat.Json, writer);

            string output = writer.ToString();
            Assert.IsNotEmpty(output, "JSON output should not be empty");
            Assert.IsTrue(output.Contains("TestFunction"), "JSON should contain symbol name");
            Assert.IsTrue(output.Contains("\"Success\":true"), "JSON should indicate success");
        }

        [Test]
        public void Test_WriteResult_Json_CanDeserialize()
        {
            SymbolSearchResult originalResult = CreateTestResult();
            StringWriter writer = new StringWriter();

            OutputFormatter.WriteResult(originalResult, OutputFormat.Json, writer);

            string json = writer.ToString();
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(SymbolSearchResult));
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                SymbolSearchResult deserializedResult = (SymbolSearchResult)serializer.ReadObject(stream);
                
                Assert.IsNotNull(deserializedResult, "Deserialized result should not be null");
                Assert.AreEqual(originalResult.Success, deserializedResult.Success);
                Assert.AreEqual(originalResult.Symbol.Name, deserializedResult.Symbol.Name);
                Assert.AreEqual(originalResult.Module.Name, deserializedResult.Module.Name);
            }
        }

        [Test]
        public void Test_WriteResult_Xml_CanSerialize()
        {
            SymbolSearchResult result = CreateTestResult();
            StringWriter writer = new StringWriter();

            OutputFormatter.WriteResult(result, OutputFormat.Xml, writer);

            string output = writer.ToString();
            Assert.IsNotEmpty(output, "XML output should not be empty");
            Assert.IsTrue(output.Contains("<SymbolSearchResult"), "XML should have root element");
            Assert.IsTrue(output.Contains("TestFunction"), "XML should contain symbol name");
        }

        [Test]
        public void Test_WriteResult_Xml_CanDeserialize()
        {
            SymbolSearchResult originalResult = CreateTestResult();
            StringWriter writer = new StringWriter();

            OutputFormatter.WriteResult(originalResult, OutputFormat.Xml, writer);

            string xml = writer.ToString();
            
            XmlSerializer serializer = new XmlSerializer(typeof(SymbolSearchResult));
            using (StringReader reader = new StringReader(xml))
            {
                SymbolSearchResult deserializedResult = (SymbolSearchResult)serializer.Deserialize(reader);
                
                Assert.IsNotNull(deserializedResult, "Deserialized result should not be null");
                Assert.AreEqual(originalResult.Success, deserializedResult.Success);
                Assert.AreEqual(originalResult.Symbol.Name, deserializedResult.Symbol.Name);
                Assert.AreEqual(originalResult.Module.Name, deserializedResult.Module.Name);
            }
        }

        [Test]
        public void Test_WriteBatchResult_Human()
        {
            BatchSymbolSearchResult batchResult = new BatchSymbolSearchResult
            {
                Success = true,
                Module = CreateTestResult().Module,
                PdbInfo = CreateTestResult().PdbInfo,
                Symbols = new System.Collections.Generic.List<SymbolSearchResult>
                {
                    new SymbolSearchResult
                    {
                        Success = true,
                        SearchedSymbolName = "Function1",
                        Symbol = new SymbolInfo { Name = "Function1", Address = 0x1000, Size = 0x50 }
                    },
                    new SymbolSearchResult
                    {
                        Success = false,
                        SearchedSymbolName = "Function2",
                        Symbol = null
                    }
                }
            };

            StringWriter writer = new StringWriter();
            OutputFormatter.WriteBatchResult(batchResult, OutputFormat.Human, writer);

            string output = writer.ToString();
            Assert.IsNotEmpty(output, "Batch output should not be empty");
            Assert.IsTrue(output.Contains("Function1"), "Output should contain first function");
            Assert.IsTrue(output.Contains("Function2"), "Output should contain second function");
            Assert.IsTrue(output.Contains("Found!"), "Output should indicate found symbol");
            Assert.IsTrue(output.Contains("Not found"), "Output should indicate not found symbol");
        }

        [Test]
        public void Test_WriteBatchResult_Json()
        {
            BatchSymbolSearchResult batchResult = new BatchSymbolSearchResult
            {
                Success = true,
                Module = CreateTestResult().Module,
                PdbInfo = CreateTestResult().PdbInfo,
                Symbols = new System.Collections.Generic.List<SymbolSearchResult>
                {
                    new SymbolSearchResult
                    {
                        Success = true,
                        SearchedSymbolName = "Function1",
                        Symbol = new SymbolInfo { Name = "Function1", Address = 0x1000, Size = 0x50 }
                    }
                }
            };

            StringWriter writer = new StringWriter();
            OutputFormatter.WriteBatchResult(batchResult, OutputFormat.Json, writer);

            string output = writer.ToString();
            Assert.IsNotEmpty(output, "JSON batch output should not be empty");
            Assert.IsTrue(output.Contains("Function1"), "JSON should contain function name");
        }

        [Test]
        public void Test_SymbolSearchResult_Serialization_PreservesNullValues()
        {
            SymbolSearchResult result = new SymbolSearchResult
            {
                Success = true,
                Symbol = null,
                Module = null,
                PdbInfo = null
            };

            StringWriter writer = new StringWriter();
            OutputFormatter.WriteResult(result, OutputFormat.Json, writer);

            string json = writer.ToString();
            Assert.IsNotEmpty(json, "JSON should be generated even with null values");
        }
    }
}
