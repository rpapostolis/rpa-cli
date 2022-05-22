using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Diagnostics;
using System.ServiceModel;

namespace RPACLI.Handlers
{

    /// <summary>
    /// Handler_Clone_Desktop_Flows
    /// </summary>
    internal static class Handler_Diff_Desktop_Flows
    {
        public const string FILE1 = "padDefinition1.json";
        public const string FILE2 = "padDefinition2.json";
        public const string FILE_WILDCARD = "padDefinition*.json";
        public const string REPORTING_FOLDER = "Reporting";
        public const string PAD_ACTIONLIST_FILENAME = "modules.json";


        /// <summary>
        /// DiffDesktopFlows
        /// </summary>
        /// <param name="Username"></param>
        /// <param name="SourceEnvironmentInstanceUrl"></param>
        /// <param name="DesktopFlowId"></param>
        /// <param name="Count"></param>
        /// <param name="BatchSize"></param>
        /// <exception cref="Exception"></exception>
        internal static void DiffDesktopFlows(string username, string desktopFlow1SourceEnvironmentInstanceUrl, string desktopFlow2SourceEnvironmentInstanceUrl, bool isFromSameEnvironment, string desktopFlowId1, string desktopFlowId2)
        {
            ServiceClient service = null;

            try
            {
                service = CLIHelper.ConnectToDataverse(username, desktopFlow1SourceEnvironmentInstanceUrl, false);

                Console.WriteLine($"Running command: {AppDomain.CurrentDomain.FriendlyName} diff-desktop-flows --Username {username} --File1SourceEnvironment {desktopFlow1SourceEnvironmentInstanceUrl} --DesktopFlowId1 {desktopFlowId1} --File2SourceEnvironment {desktopFlow2SourceEnvironmentInstanceUrl} --DesktopFlowId2 {desktopFlowId2} --IsFromSameEnvironment {isFromSameEnvironment}\n");

                string reportingPath = Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "Reporting");

                DesktopFlowModule modules = null;

                // Get the Desktop flow action list 
                using (StreamReader sr = new StreamReader(Path.Combine(reportingPath, PAD_ACTIONLIST_FILENAME)))
                {
                    modules = JsonConvert.DeserializeObject<DesktopFlowModule>(sr.ReadToEnd());
                }

                string desktopFlowScript1 = String.Empty;
                string desktopFlowScript2 = String.Empty;
                string fullFilePath1 = Path.Combine(reportingPath, FILE1);
                string fullFilePath2 = Path.Combine(reportingPath, FILE2);

                Console.WriteLine($"Start Time: {DateTime.Now.ToLongTimeString()}\n");

                if (isFromSameEnvironment || desktopFlow1SourceEnvironmentInstanceUrl.Equals(desktopFlow2SourceEnvironmentInstanceUrl))
                {
                    Console.WriteLine($"Comparing Desktop flow {desktopFlowId1} with {desktopFlowId2} from {desktopFlow1SourceEnvironmentInstanceUrl}\n");
                    Console.WriteLine($"Extracting Desktop flow definition from {desktopFlowId1}\n");
                    desktopFlowScript1 = ExtractDesktopFlowDefinition(service, desktopFlowId1);
                    Console.WriteLine($"Extracting Desktop flow definition from {desktopFlowId2}\n");
                    desktopFlowScript2 = ExtractDesktopFlowDefinition(service, desktopFlowId2);

                    bool differencesFound = !desktopFlowScript1.Equals(desktopFlowScript2);
                    GenerateDesktopFlowActionInventoryFile(service, modules, desktopFlowId1, desktopFlowScript1, fullFilePath1, differencesFound);
                    GenerateDesktopFlowActionInventoryFile(service, modules, desktopFlowId2, desktopFlowScript2, fullFilePath2, differencesFound);
                }
                else
                {
                    Console.WriteLine($"Comparing Desktop flow {desktopFlowId1} from {desktopFlow1SourceEnvironmentInstanceUrl} with {desktopFlowId2} from {desktopFlow2SourceEnvironmentInstanceUrl}\n");
                    Console.WriteLine($"Extracting Desktop flow definition from {desktopFlowId1}\n");
                    desktopFlowScript1 = ExtractDesktopFlowDefinition(service, desktopFlowId1);
                    bool differencesFound = !desktopFlowScript1.Equals(desktopFlowScript2);
                    GenerateDesktopFlowActionInventoryFile(service, modules, desktopFlowId1, desktopFlowScript1, fullFilePath1, differencesFound);
                    service = CLIHelper.ConnectToDataverse(username, desktopFlow2SourceEnvironmentInstanceUrl,false);
                    Console.WriteLine($"Extracting Desktop flow definition from {desktopFlowId2}\n");
                    desktopFlowScript2 = ExtractDesktopFlowDefinition(service, desktopFlowId2);
                    GenerateDesktopFlowActionInventoryFile(service, modules, desktopFlowId2, desktopFlowScript2, fullFilePath2, differencesFound);
                }

                Console.WriteLine($"Opening Visual Studio Code to compare scripts...");

                OpenVSCodeToCompareFiles(fullFilePath1, fullFilePath2);

                Console.WriteLine($"\nEnd Time: {DateTime.Now.ToLongTimeString()}\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (service != null)
                    service.Dispose();

                Console.WriteLine("Press <Enter> to exit.");
                Console.ReadLine();

            }
        }

        internal static void GenerateDesktopFlowActionInventoryFile(ServiceClient service, DesktopFlowModule modules, string desktopFlowId, string desktopFlowScript, string fullExportFileName, bool differencesFound)
        {
            string reportingPath = Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "Reporting");

            Entity desktopFlow = service.Retrieve("workflow", Guid.Parse(desktopFlowId), new ColumnSet("workflowid", "name","ownerid", "owninguser"));

            List<DesktopFlowActionUsage> usedActions = new List<DesktopFlowActionUsage>();

            foreach (Module mod in modules.Modules)
            {
                foreach (Action act in mod.Actions)
                {
                    foreach (string selector in act.SelectorIds)
                    {
                        // We are not interested in Set Variable actions hence we're not searching for text: 'SET NewVar TO '. However, this could be 
                        // done by using this RegEx: SET+.*[a-zA-Z]+.*TO+
                        int regExMatchCount = CLIHelper.FindStringOccurrencesWithRegEx(desktopFlowScript, @"(" + selector + @"\b)", true);

                        // Only if we find an occurance
                        if (regExMatchCount > 0)
                        {
                            usedActions.Add(new DesktopFlowActionUsage
                            {
                                actionname = act.FriendlyName,
                                desktopflowactionid = act.Id,
                                modulesource = mod.ModuleSource,
                                modulename = mod.FriendlyName,
                                selectorid = selector,
                                desktopflowid = "[In header info]",
                                desktopflowname= "[In header info]",
                                occurrencecount = regExMatchCount
                            });

                        }
                    }
                }
            }

            try
            {

                FileInfo diffFile = new FileInfo(fullExportFileName);
                diffFile.Delete();

                using (File.Create(fullExportFileName)) { };

            }
            catch (IOException ioExp)
            {
                Console.WriteLine(ioExp.Message);
            }

            DesktopFlowActionUsageExport dfExport = new DesktopFlowActionUsageExport();
            dfExport.environmentid = service.OrganizationDetail.EnvironmentId;
            dfExport.environmenturlname = service.OrganizationDetail.UrlName;
            dfExport.desktopflowhasdifferences = differencesFound;
            dfExport.desktopflowid = desktopFlowId;
            dfExport.desktopflowname = desktopFlow["name"].ToString();
            dfExport.ownerid = ((EntityReference)desktopFlow["ownerid"]).Id.ToString();
            dfExport.ownername = ((EntityReference)desktopFlow["ownerid"]).Name;
            dfExport.actions = usedActions;

            using (StreamWriter sw = new StreamWriter(fullExportFileName, true))
            {
                sw.Write(JsonConvert.SerializeObject(dfExport, Formatting.Indented));
            }

        }
        internal static string ExtractDesktopFlowDefinition(ServiceClient service, string desktopFlowId)
        {
            Entity desktopFlow = service.Retrieve("workflow", Guid.Parse(desktopFlowId), new ColumnSet("name", "workflowid", "clientdata"));
            string desktopFlowScript = string.Empty;

            if (desktopFlow != null)
            {
                try
                {
                    string clientdata = desktopFlow.Contains("clientdata") ? desktopFlow.Attributes["clientdata"].ToString() : "";

                    if (!string.IsNullOrEmpty(clientdata))
                    {
                        string desktopFlowZip = CLIHelper.GetZippedDesktopFlowScript(clientdata);

                        if (!string.IsNullOrWhiteSpace(desktopFlowZip))
                        {
                            // Unpack and extract Desktop flow script
                            string desktopFlowScriptWithVars = CLIHelper.UnzipDesktopFlowScript(desktopFlowZip, true);

                            // Get cleaned-up script without input and output variables
                            desktopFlowScript = CLIHelper.RemoveVariablesFromDesktopFlowScript(desktopFlowScriptWithVars);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine($"The application terminated with an error. {ex.Message}");

                    // Display the details of the inner exception.
                    if (ex.InnerException != null)
                    {
                        FaultException<OrganizationServiceFault> fe = ex.InnerException
                            as FaultException<OrganizationServiceFault>;
                        if (fe != null)
                        {
                            Console.WriteLine($"Timestamp: {fe.Detail.Timestamp}");
                            Console.WriteLine($"Code: {fe.Detail.ErrorCode}");
                            Console.WriteLine($"Message: {fe.Detail.Message}");
                            Console.WriteLine($"Trace: {fe.Detail.TraceText}");
                        }
                    }
                    else
                        throw;
                }
            }
            else
                throw new Exception($"Desktop flow with Id {desktopFlowId} cannot be found.");

            return desktopFlowScript;
        }

        public static void OpenVSCodeToCompareFiles(string filename1, string filename2)
        {
            var pi = new ProcessStartInfo
            {
                UseShellExecute = true,
                FileName = "code",
                Arguments = $"-d {filename1} {filename2}",
                WindowStyle = ProcessWindowStyle.Hidden
            };

            try
            {
                Process.Start(pi);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
