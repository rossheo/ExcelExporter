using CommandLine;
using System.IO;
using System;
using log4net.Config;
using log4net;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ExcelExporter
{
    class Program
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(Program));

        // https://github.com/commandlineparser/commandline
        public class Options
        {
            public const string DefaultServerHeaderPath = "../serverHeader";
            public const string DefaultClientHeaderPath = "../clientHeader";
            public const string DefaultServerJsonPath = "../serverJson";
            public const string DefaultClientJsonPath = "../clientJson";

            [Option('e', "export", Required = false, HelpText = "Export excel to json.")]
            public bool IsExport { get; set; }

            [Option('s', "server", Required = false, HelpText = "Export server data.")]
            public bool IsExportServerData { get; set; }

            [Option('c', "client", Required = false, HelpText = "Export client data.")]
            public bool IsExportClientData { get; set; }

            [Option("excel-path", Required = false, HelpText = "Excel path.")]
            public string ExcelPath { get; set; }

            [Option("server-header-path", Default = DefaultServerHeaderPath, Required = false, HelpText = "Export server header path.")]
            public string ServerHeaderPath { get; set; }

            [Option("client-header-path", Default = DefaultClientHeaderPath, Required = false, HelpText = "Export client header path.")]
            public string ClientHeaderPath { get; set; }

            [Option("server-json-path", Default = DefaultServerJsonPath, Required = false, HelpText = "Export server json path.")]
            public string ServerJsonPath { get; set; }

            [Option("client-json-path", Default = DefaultClientJsonPath, Required = false, HelpText = "Export client json path.")]
            public string ClientJsonPath { get; set; }
        }

        private static void ModifyOptions(ref Options options)
        {
            if (options.ExcelPath != null && options.ExcelPath.Length > 0)
            {
                string excelPath = string.Empty;

                if (Directory.Exists(options.ExcelPath))
                {
                    excelPath = options.ExcelPath;
                }
                else if (File.Exists(options.ExcelPath))
                {
                    excelPath = Path.GetDirectoryName(options.ExcelPath);
                }

                if (options.ServerHeaderPath == Options.DefaultServerHeaderPath)
                {
                    options.ServerHeaderPath = Path.Combine(excelPath, options.ServerHeaderPath);
                }

                if (options.ClientHeaderPath == Options.DefaultClientHeaderPath)
                {
                    options.ClientHeaderPath = Path.Combine(excelPath, options.ClientHeaderPath);
                }

                if (options.ServerJsonPath == Options.DefaultServerJsonPath)
                {
                    options.ServerJsonPath = Path.Combine(excelPath, options.ServerJsonPath);
                }

                if (options.ClientJsonPath == Options.DefaultClientJsonPath)
                {
                    options.ClientJsonPath = Path.Combine(excelPath, options.ClientJsonPath);
                }
            }
        }

        static void Main(string[] args)
        {
            Environment.ExitCode = 1;

            XmlConfigurator.Configure(new FileInfo("config/log4net.xml"));

            var parser = new Parser(config => config.HelpWriter = Console.Out);

            if (args.Length == 0)
            {
                parser.ParseArguments<Options>(new[] { "--help" });
                return;
            }

            parser.ParseArguments<Options>(args)
                .WithParsed(options =>
            {
                ModifyOptions(ref options);

                if (options.IsExport)
                {
                    if (ExcelExport(options))
                    {
                        Environment.ExitCode = 0;
                        return;
                    }
                }
            });
        }

        static bool ExcelExport(Options options)
        {
            string excelPath = options.ExcelPath;

            System.Data.DataSet rawDataSet = new System.Data.DataSet();

            using (ExcelTableToRawDataSet excelTableToRawDataSet =
                new ExcelTableToRawDataSet(excelPath))
            {
                if (!excelTableToRawDataSet.Execute(ref rawDataSet))
                {
                    Log.WarnFormat("Fail to execute excelTableToRawDataSet. excelPath: {0}",
                        excelTableToRawDataSet.ExcelPath);
                    return false;
                }
            }

            using (MergingRawDataSet mergingRawDataSet = new MergingRawDataSet())
            {
                if (!mergingRawDataSet.Execute(ref rawDataSet))
                {
                    Log.Warn("Fail to execute mergingRawDataSet.");
                    return false;
                }
            }

            // ServerData
            if (options.IsExportServerData)
            {
                System.Data.DataSet refinedServerDataSet = new System.Data.DataSet();

                using (RefineDataSet refineDataSet = new RefineDataSet())
                {
                    if (!refineDataSet.ExecuteServerData(rawDataSet, ref refinedServerDataSet))
                    {
                        Log.Warn("Fail to refine server dataset.");
                        return false;
                    }
                }

                string exportServerJsonPath = options.ServerJsonPath;

                using (ExportServerJsonFiles exportServerJsonFiles =
                    new ExportServerJsonFiles(exportServerJsonPath))
                {
                    if (!exportServerJsonFiles.Execute(refinedServerDataSet, rawDataSet))
                    {
                        Log.Warn("Fail to export server json files.");
                        return false;
                    }
                }

                string serverEnumHeaderExportPath = options.ServerHeaderPath;

                using (GenerateServerEnumHeader generateServerEnumHeader =
                    new GenerateServerEnumHeader(serverEnumHeaderExportPath))
                {
                    if (!generateServerEnumHeader.Execute(rawDataSet))
                    {
                        Log.Warn("Fail to generate server enum header.");
                        return false;
                    }
                }

                string serverDataHeaderExportPath = options.ServerHeaderPath;

                using (GenerateServerDataHeader generateServerDataHeader =
                    new GenerateServerDataHeader(serverDataHeaderExportPath))
                {
                    if (!generateServerDataHeader.Execute(rawDataSet))
                    {
                        Log.Warn("Fail to generate server data header.");
                        return false;
                    }
                }
            }

            // ClientData
            if (options.IsExportClientData)
            {
                System.Data.DataSet refinedClientDataSet = new System.Data.DataSet();

                using (RefineDataSet refineDataSet = new RefineDataSet())
                {
                    if (!refineDataSet.ExecuteClientData(rawDataSet, ref refinedClientDataSet))
                    {
                        Log.Warn("Fail to refine client dataset.");
                        return false;
                    }
                }

                string exportClientJsonPath = options.ClientJsonPath;

                using (ExportClientJsonFiles exportClientJsonFiles =
                    new ExportClientJsonFiles(exportClientJsonPath))
                {
                    if (!exportClientJsonFiles.Execute(refinedClientDataSet, rawDataSet))
                    {
                        Log.Warn("Fail to export client json files.");
                        return false;
                    }
                }

                string clientEnumHeaderPath = options.ClientHeaderPath;

                using (GenerateClientEnumHeader generateClientEnumHeader =
                    new GenerateClientEnumHeader(clientEnumHeaderPath))
                {
                    if (!generateClientEnumHeader.Execute(rawDataSet))
                    {
                        Log.Warn("Fail to generate client enum header.");
                        return false;
                    }
                }

                string clientDataHeaderPath = options.ClientHeaderPath;

                using (GenerateClientDataHeader generateClientDataHeader =
                    new GenerateClientDataHeader(clientDataHeaderPath))
                {
                    if (!generateClientDataHeader.Execute(rawDataSet))
                    {
                        Log.Warn("Fail to generate client data header.");
                        return false;
                    }
                }
            }

            Log.InfoFormat("Export excel completed: {0}", excelPath);
            return true;
        }
    }
} 