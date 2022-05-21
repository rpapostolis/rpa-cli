using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Diagnostics;
using System.Globalization;

namespace RPACLI.Handlers
{
    /// <summary>
    /// Handler_Get_Runs
    /// </summary>
    internal static class Handler_Get_Runs
    {
        public const string DESKOPTFLOW_SESSION_FILENAME = "DesktopFlowSessions.csv";

        /// <summary>
        /// GetDesktopFlowRuns
        /// </summary>
        /// <param name="Username"></param>
        /// <param name="SourceEnvironmentInstanceUrl"></param>
        /// <param name="DesktopFlowId"></param>
        /// <param name="Timeframe"></param>
        /// <param name="Export"></param>
        /// <param name="RootExportFolder"></param>
        internal static void GetDesktopFlowRuns(string Username, string SourceEnvironmentInstanceUrl, string DesktopFlowId, int Timeframe, bool Export, string RootExportFolder)
        {
            ServiceClient service = null;
            try
            {
                service = CLIHelper.ConnectToDataverse(Username, SourceEnvironmentInstanceUrl);

                Console.WriteLine($"Running command: {AppDomain.CurrentDomain.FriendlyName} get-runs --Username {Username} --SourceEnvironmentInstanceUrl {SourceEnvironmentInstanceUrl} --DesktopFlowId {DesktopFlowId} --Timeframe {Timeframe}");

                Console.WriteLine($"\n###############################################################################################################");
                Console.WriteLine($"Retrieving {Timeframe}-day(s) old Desktop flow sessions for {DesktopFlowId}");
                Console.WriteLine($"###############################################################################################################\n");
                Console.WriteLine("{0} {1} {2} {3} {4}",
                    "Desktop Flow Name".ToString().PadRight(55, ' '),
                    "Last Run".ToString().PadRight(17, ' '),
                    "Runs".ToString().PadRight(8, ' '),
                    "Succeeded".ToString().PadRight(10, ' '),
                    "Failed".ToString().PadRight(10, ' ')
                );

                ExportDesktopFlowSessions(service,DesktopFlowId,RootExportFolder,Export,Timeframe);

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

        /// <summary>
        /// ExportDesktopFlowSessions
        /// </summary>
        /// <param name="service"></param>
        /// <param name="desktopFlowId"></param>
        /// <param name="rootExportFolder"></param>
        /// <param name="export"></param>
        /// <param name="timeframe"></param>
        static void ExportDesktopFlowSessions(ServiceClient service, string desktopFlowId, string rootExportFolder, bool export,int timeframe)
        {
            // Set the number of records per page to retrieve to 1000.
            int fetchCount = 1000;
            // Initialize the page number.
            int pageNumber = 1;
            // Specify the current paging cookie.
            // For retrieving the first page, pagingCookie should be null.
            string pagingCookie = null;
            // Create the FetchXml string for retrieving all Desktop flow session.
            string fetchXml = string.Format(@"<fetch>
                                                  <entity name='flowsession'>
                                                    <attribute name='errordetails' />
                                                    <attribute name='statecode' />
                                                    <attribute name='statecodename' />
                                                    <attribute name='regardingobjectid' />
                                                    <attribute name='regardingobjectidname' />
                                                    <attribute name='context' />
                                                    <attribute name='machinegroupid' />
                                                    <attribute name='errorcode' />
                                                    <attribute name='machineid' />
                                                    <attribute name='statuscodename' />
                                                    <attribute name='machinegroupidname' />
                                                    <attribute name='machineidname' />
                                                    <attribute name='statuscode' />
                                                    <attribute name='createdon' />
                                                    <attribute name='flowsessionid' />
                                                    <attribute name='completedon' />
                                                    <attribute name='errormessage' />
                                                    <attribute name='gateway' />
                                                    <attribute name='startedon' />
                                                    <filter type='and'>
                                                      <condition attribute='regardingobjectid' operator='eq' value='" + desktopFlowId + @"' />"
                                                       + (timeframe > 0 ? "<condition attribute='createdon' operator='last-x-days' value='" + timeframe + @"' />" : "") +
                                                    @"</filter>
                                                    <filter type='or'>
                                                      <condition attribute='statuscode' operator='eq' value='8' />
                                                      <condition attribute='statuscode' operator='eq' value='4' />
                                                    </filter>
                                                    <order attribute='createdon' descending='true' />
                                                  </entity>
                                                </fetch>");

            int totalSessions = 0;
            int totalSuccesses = 0;
            int totalFailures = 0;

            string lastRundate = "";
            bool isLatestRun = true;

            List<DesktopFlowSession> desktopFlowSessions = new List<DesktopFlowSession>();

            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CLIHelper.CreateXml(fetchXml, pagingCookie, pageNumber, fetchCount);

                // Excute the fetch query and get the xml result.
                var fr = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                EntityCollection desktopFlowSessionCollection = ((RetrieveMultipleResponse)service.Execute(fr)).EntityCollection;
                string desktopFlowName = string.Empty;

                if (desktopFlowSessionCollection.Entities.Count < 1) 
                { 
                    Console.WriteLine("\n\n****************\nDone - No Flow Sessions available.\n****************\n");
                    return;
                }

                foreach (var s in desktopFlowSessionCollection.Entities)
                {
                    if (isLatestRun)
                        lastRundate = s.Contains("createdon") ? s.Attributes["createdon"].ToString() : "";

                    desktopFlowSessions.Add(new DesktopFlowSession
                    {
                        errordetails = s.Contains("errordetails") ? s.Attributes["errordetails"].ToString() : "",
                        statecode = s.Contains("statecode") ? s.FormattedValues["statecode"].ToString() : "",
                        regardingobjectid = s.Contains("regardingobjectid") ? ((EntityReference)s.Attributes["regardingobjectid"]).Id.ToString() : "",
                        regardingobjectidname = s.Contains("regardingobjectid") ? s.FormattedValues["regardingobjectid"].ToString() : "",
                        context = s.Contains("context") ? s.Attributes["context"].ToString() : "",
                        machinegroupid = s.Contains("machinegroupid") ? ((EntityReference)s.Attributes["machinegroupid"]).Id.ToString() : "",
                        errorcode = s.Contains("errorcode") ? s.Attributes["errorcode"].ToString() : "",
                        machineid = s.Contains("machineid") ? ((EntityReference)s.Attributes["machineid"]).Id.ToString() : "",
                        createdon = s.Contains("createdon") ? s.Attributes["createdon"].ToString() : "",
                        machinegroupidname = s.Contains("machinegroupid") ? s.FormattedValues["machinegroupid"].ToString() : "",
                        machineidname = s.Contains("machineid") ? s.FormattedValues["machineid"].ToString() : "",
                        statuscode = s.Contains("statuscode") ? s.FormattedValues["statuscode"].ToString() : "",
                        flowsessionid = s.Contains("flowsessionid") ? s.Attributes["flowsessionid"].ToString() : "",
                        completedon = s.Contains("completedon") ? s.Attributes["completedon"].ToString() : "",
                        errormessage = s.Contains("errormessage") ? s.Attributes["errormessage"].ToString() : "",
                        gateway = s.Contains("gateway") ? s.Attributes["gateway"].ToString() : "",
                        startedon = s.Contains("startedon") ? s.Attributes["startedon"].ToString() : ""
                    });

                    if (s.FormattedValues["statuscode"].ToString().Equals("Failed"))
                        totalFailures++;
                    else
                        totalSuccesses++;

                    totalSessions++;

                    isLatestRun = false;

                    desktopFlowName = s.Contains("regardingobjectid") ? s.FormattedValues["regardingobjectid"].ToString() : "";
                }

                // Check for morerecords, if it returns 1.
                if (desktopFlowSessionCollection.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = desktopFlowSessionCollection.PagingCookie;

                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                }
                else
                {
                    if (desktopFlowSessions.First() != null) 
                    { 
                        Console.WriteLine("{0} {1} {2} {3} {4}\n\n****************\nDone\n****************\n",
                            desktopFlowSessions.First().regardingobjectidname.PadRight(55, ' '),
                            string.IsNullOrWhiteSpace(lastRundate) ? "None".PadRight(17, ' ') : DateTime.Parse(lastRundate).ToShortDateString().PadRight(17, ' '),
                            totalSessions.ToString().PadRight(8, ' '),
                            totalSuccesses.ToString().PadRight(10, ' '),
                            totalFailures.ToString().PadRight(10, ' ')
                        );
                    }
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }

            if (export) { 
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    NewLine = Environment.NewLine,
                    HasHeaderRecord = true
                };

                string reportingPath = Path.Combine(rootExportFolder, "Reporting");

                using (var desktopFlowSessionWriter = new StreamWriter(Path.Combine(reportingPath, desktopFlowId + "_" + DESKOPTFLOW_SESSION_FILENAME)))
                using (var csvDesktopFlowSessions = new CsvWriter(desktopFlowSessionWriter, config))
                {
                    csvDesktopFlowSessions.WriteRecords(desktopFlowSessions);
                }
                if (Directory.Exists(reportingPath))
                    Process.Start("explorer.exe", reportingPath);
            }
        }
    }
}
