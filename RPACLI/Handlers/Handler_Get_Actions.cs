using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.PowerPlatform.Dataverse.Client;
using System.ServiceModel;
using CsvHelper.Configuration;
using System.Globalization;
using CsvHelper;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Text;

namespace RPACLI.Handlers
{

    /// <summary>
    /// Handler_Get_Actions
    /// </summary>
    internal static class Handler_Get_Actions
    {
        public const string PROSESSING_LOG_FILENAME = "_processingLog.txt";
        public const string PAD_ACTIONLIST_FILENAME = "modules.json";
        public const string DESKOPTFLOW_FILENAME = "DesktopFlows.csv";
        public const string DESKOPTFLOW_ACTION_FILENAME = "DesktopFlowActions.csv";
        public const string DESKOPTFLOW_SESSION_FILENAME = "DesktopFlowSessions.csv";


        /// <summary>
        /// GenerateDesktopFlowActionStats
        /// </summary>
        /// <param name="username"></param>
        /// <param name="outputTarget"></param>
        /// <param name="sourceEnvironmentInstanceUrl"></param>
        /// <param name="exportRuns"></param>
        /// <param name="batchSize"></param>
        /// <param name="continueFlag"></param>
        /// <param name="targetEnvironmentInstanceUrl"></param>
        /// <param name="targetEnvironmentId"></param>
        /// <param name="rootExportFolder"></param>
        /// <exception cref="Exception"></exception>
        internal static void GenerateDesktopFlowActionStats(
            string username, string outputTarget, string sourceEnvironmentInstanceUrl, bool exportRuns, int batchSize, 
            bool continueFlag, string targetEnvironmentInstanceUrl, string targetEnvironmentId,
            string rootExportFolder)
        {

            string target = outputTarget.ToLower();

            if (!target.Equals("csv") && !target.Equals("autocoe") && !target.ToLower().Equals("ppcoe"))
                throw new Exception($"Wrong target parameter specified. The entered value was: {outputTarget} but we expect the value to be either CSV, AutoCoE or PPCoE.");

            ServiceClient service = null;

            try
            {
                service = CLIHelper.ConnectToDataverse(username, sourceEnvironmentInstanceUrl);

                List<string> processedDesktopFlowIdsFromLog = new List<string>();

                // Initialize the page number.
                int pageNumber = 1;
                // Initialize the number of records.
                int recordCount = 0;
                // Specify the current paging cookie.
                // For retrieving the first page, pagingCookie should be null.
                string pagingCookie = null;
                // Create the FetchXml string for retrieving all Desktop flows and their maker info.
                string fetchXml = @"<fetch>
                                        <entity name='workflow'>
                                        <attribute name='category' />
                                        <attribute name='modifiedon' />
                                        <attribute name='uiflowtype' />
                                        <attribute name='businessprocesstype' />
                                        <attribute name='ownerid' />
                                        <attribute name='mode' />
                                        <attribute name='name' />
                                        <attribute name='workflowidunique' />
                                        <attribute name='type' />
                                        <attribute name='workflowid' />
                                        <attribute name='modifiedby' />
                                        <attribute name='createdby' />
                                        <attribute name='createdon' />
                                        <attribute name='clientdata' />
                                        <filter>
                                            <condition attribute='category' operator='eq' value='6' />";
                        
                        // If the continue flag has been set to true then we can get
                        // the previously processed Desktop flow Ids and exclude these
                        // from the current prcess to speed-up the overall loading process
                        if (continueFlag)
                        {
                            processedDesktopFlowIdsFromLog = ReadDesktopFlowIdsProcessingLog(rootExportFolder);

                            if (processedDesktopFlowIdsFromLog.Count > 0)
                            {
                                StringBuilder sb = new StringBuilder();

                                sb.AppendLine("<condition attribute='workflowid' operator='not-in'>");
                                processedDesktopFlowIdsFromLog.ForEach(x => sb.AppendLine("<value>" + x + "</value>"));
                                sb.AppendLine("</condition>");

                                fetchXml += sb.ToString();
                            }
                        };

                        fetchXml += @"</filter>
                                        <order attribute='name' />
                                        <link-entity name='systemuser' from='systemuserid' to='owninguser' alias='User'>
                                            <attribute name='organizationid' />
                                            <attribute name='fullname' />
                                            <attribute name='personalemailaddress' />
                                            <attribute name='createdbyname' />
                                            <attribute name='domainname' />
                                            <attribute name='firstname' />
                                            <attribute name='middlename' />
                                            <attribute name='internalemailaddress' />
                                            <attribute name='lastname' />
                                        </link-entity>
                                        </entity>
                                    </fetch>";
                
                List<DesktopFlowDefinition> deskFlows = new List<DesktopFlowDefinition>();
                List<DesktopFlowAction> deskFlowActions = new List<DesktopFlowAction>();
                List<DesktopFlowSession> deskFlowSessions = new List<DesktopFlowSession>();
                List<DesktopFlowActionUsage> usedActions = new List<DesktopFlowActionUsage>();

                List<string> missingActions = new List<string>();
                List<string> makers = new List<string>();
                List<string> processedDesktopFlowIds = new List<string>();

                int totalActionsUsed = 0;
                int totalOccurrences = 0;
                int totalUpserts = 0;

                if (target.Equals("autocoe"))
                {
                    Console.WriteLine($"Running command: {AppDomain.CurrentDomain.FriendlyName} get-actions " +
                        $"--Username {username} " +
                        $"--Target {outputTarget} " +
                        $"--Batchsize {batchSize} " +
                        $"--Continue {continueFlag} " +
                        $"--SourceEnvironmentInstanceUrl {sourceEnvironmentInstanceUrl} " +
                        $"--TargetEnvironmentInstanceUrl {targetEnvironmentInstanceUrl} " +
                        $"--TargetEnvironmentId {targetEnvironmentId}" );
        
                    deskFlowActions = PrepareAutoCoEUpdate(service, rootExportFolder, continueFlag);
                }
                else
                {
                    Console.WriteLine($"Running command: {AppDomain.CurrentDomain.FriendlyName} get-actions " +
                        $"--Username {username} " +
                        $"--Target {target} " +
                        $"--Batchsize {batchSize} " +
                        $"--ExportRuns {exportRuns} " +
                        $"--Continue {continueFlag} " +
                        $"--SourceEnvironmentInstanceUrl {sourceEnvironmentInstanceUrl} " +
                        $"--RootExportFolder {rootExportFolder}");

                    deskFlowActions = PrepareCSVExport(rootExportFolder, continueFlag);
                }

                Console.WriteLine($"\n###############################################################################################################");
                Console.WriteLine($"Retrieving paged Desktop flow definitions in batches of {batchSize}");
                Console.WriteLine($"Note: Some of the standard actions such as 'Set Variable', 'Loop While' etc. are excluded from parsing.");
                Console.WriteLine($"###############################################################################################################\n");
                Console.WriteLine("\n****************\nStarting query process for desktop flows... This might take a few minutes...\n****************");
                Console.WriteLine($"Start Time: {DateTime.Now.ToLongTimeString()}\n");

                while (true)
                {
                    // Build fetchXml string with the placeholders.
                    string xml = CLIHelper.CreateXml(fetchXml, pagingCookie, pageNumber, batchSize);

                    // Excute the fetch query and get the xml result.
                    var fr = new RetrieveMultipleRequest
                    {
                        Query = new FetchExpression(xml)
                    };

                    EntityCollection desktopFlowCollection = ((RetrieveMultipleResponse)service.Execute(fr)).EntityCollection;

                    Console.WriteLine($"\n****************\nProcessing batch {pageNumber}... {recordCount+desktopFlowCollection.Entities.Count} Desktop flows loaded... \n****************\n");

                    string desktopFlowScript = string.Empty;

                    if (exportRuns)
                    {
                        Console.WriteLine("{0} {1} {2} {3} {4} {5} {6}",
                            "#".ToString().PadRight(4, ' '),
                            "Desktop Flow Name".ToString().PadRight(47, ' '),
                            "Actions".ToString().PadRight(10, ' '),
                            "Last Run".ToString().PadRight(19, ' '),
                            "Runs".ToString().PadRight(8, ' '),
                            "Succeeded".ToString().PadRight(10, ' '),
                            "Failed".ToString().PadRight(10, ' ')
                        );
                    }
                    else
                    {
                        Console.WriteLine("{0} {1} {2}",
                            "#".ToString().PadRight(4, ' '),
                            "Desktop Flow Name".ToString().PadRight(47, ' '),
                            "Actions".ToString().PadRight(10, ' ')
                        );
                    }

                    foreach (var c in desktopFlowCollection.Entities)
                    {
                        string clientdata = c.Contains("clientdata") ? c.Attributes["clientdata"].ToString() : "";

                        if (!string.IsNullOrEmpty(clientdata))
                        {
                            string desktopFlowZip = CLIHelper.GetZippedDesktopFlowScript(clientdata);

                            if (!string.IsNullOrWhiteSpace(desktopFlowZip))
                            {
                                // Unpack and extract Desktop flow script
                                string desktopFlowScriptWithVars = CLIHelper.UnzipDesktopFlowScript(desktopFlowZip, true);

                                // Get cleaned-up script without input and output variables
                                desktopFlowScript = CLIHelper.RemoveVariablesFromDesktopFlowScript(desktopFlowScriptWithVars);

                                DesktopFlowDefinition newDesktopFlowDef = new DesktopFlowDefinition
                                {
                                    category = c.Contains("category") ? c.FormattedValues["category"].ToString() : "",
                                    modifiedon = c.Contains("modifiedon") ? DateTime.Parse(c.Attributes["modifiedon"].ToString()).ToString("G",DateTimeFormatInfo.InvariantInfo) : "",
                                    uiflowtype = c.Contains("uiflowtype") ? c.FormattedValues["uiflowtype"].ToString() : "",
                                    businessprocesstype = c.Contains("businessprocesstype") ? c.FormattedValues["businessprocesstype"].ToString() : "",
                                    mode = c.Contains("mode") ? c.FormattedValues["mode"].ToString() : "",
                                    name = c.Contains("name") ? c.Attributes["name"].ToString() : "",
                                    workflowidunique = c.Contains("workflowidunique") ? c.Attributes["workflowidunique"].ToString() : "",
                                    type = c.Contains("type") ? c.FormattedValues["type"].ToString() : "",
                                    workflowid = c.Contains("workflowid") ? c.Attributes["workflowid"].ToString() : "",
                                    createdon = c.Contains("createdon") ? DateTime.Parse(c.Attributes["createdon"].ToString()).ToString("G",DateTimeFormatInfo.InvariantInfo) : "",
                                    clientdata = "Removed during processing",
                                    organizationid = c.Contains("User.organizationid") ? c.GetAttributeValue<AliasedValue>("User.organizationid").Value.ToString() : "",
                                    organizationidname = c.Contains("User.organizationidname") ? c.GetAttributeValue<AliasedValue>("User.organizationidname").Value.ToString() : "",
                                    ownerid = c.Contains("ownerid") ? ((EntityReference)c.Attributes["ownerid"]).Id.ToString() : "",
                                    fullname = c.Contains("User.fullname") ? c.GetAttributeValue<AliasedValue>("User.fullname").Value.ToString() : "",
                                    domainname = c.Contains("User.domainname") ? c.GetAttributeValue<AliasedValue>("User.domainname").Value.ToString() : "",
                                    firstname = c.Contains("User.firstname") ? c.GetAttributeValue<AliasedValue>("User.firstname").Value.ToString() : "",
                                    middlename = c.Contains("User.middlename") ? c.GetAttributeValue<AliasedValue>("User.middlename").Value.ToString() : "",
                                    internalemailaddress = c.Contains("User.internalemailaddress") ? c.GetAttributeValue<AliasedValue>("User.internalemailaddress").Value.ToString() : "",
                                    lastname = c.Contains("User.lastname") ? c.GetAttributeValue<AliasedValue>("User.lastname").Value.ToString() : ""
                                };

                                if (exportRuns)
                                    PrepareDesktopFlowSessionCSVExport(service, ref newDesktopFlowDef, ref deskFlowSessions);

                                deskFlows.Add(newDesktopFlowDef);

                                int actionCount = 0;

                                if (!makers.Contains(((EntityReference)c.Attributes["ownerid"]).Id.ToString()))
                                    makers.Add(((EntityReference)c.Attributes["ownerid"]).Id.ToString());

                                foreach (DesktopFlowAction action in deskFlowActions)
                                {
                                    // We are not interested in Set Variable actions hence we're not searching for text: 'SET NewVar TO '. However, this could be 
                                    // done by using this RegEx: SET+.*[a-zA-Z]+.*TO+
                                    int regExMatchCount = CLIHelper.FindStringOccurrencesWithRegEx(desktopFlowScript, @"(" + action.selectorid + @"\b)", true);

                                    // Only if we find an occurance
                                    if (regExMatchCount > 0)
                                    {
                                        usedActions.Add(new DesktopFlowActionUsage
                                        {
                                            actionname = action.actionname,
                                            desktopflowactionid = action.desktopflowactionid,
                                            modulesource = action.modulesource,
                                            modulename = action.modulename,
                                            selectorid = action.selectorid,
                                            desktopflowname = c.Attributes["name"].ToString(),
                                            desktopflowid = c.Attributes["workflowid"].ToString(),
                                            occurrencecount = regExMatchCount
                                        });
                                        actionCount++;
                                        totalOccurrences += regExMatchCount;
                                    }
                                }
                                int recCnt = ++recordCount;
                                totalActionsUsed += actionCount;

                                if (exportRuns)
                                {
                                    Console.WriteLine("{0} {1} {2} {3} {4} {5} {6}",
                                        recCnt.ToString().PadRight(4, ' '),
                                        c.Attributes["name"].ToString().PadRight(47, ' '),
                                        actionCount.ToString().PadRight(10, ' '),
                                        string.IsNullOrWhiteSpace(newDesktopFlowDef.lastrundate) ? "None".PadRight(19, ' ') : newDesktopFlowDef.lastrundate.PadRight(17, ' '),
                                        newDesktopFlowDef.totalruns.ToString().PadRight(8, ' '),
                                        newDesktopFlowDef.succeededruns.ToString().PadRight(10, ' '),
                                        newDesktopFlowDef.failedruns.ToString().PadRight(10, ' '));
                                }
                                else
                                {
                                    Console.WriteLine("{0} {1} {2}",
                                        recCnt.ToString().PadRight(4, ' '),
                                        c.Attributes["name"].ToString().PadRight(47, ' '),
                                        actionCount.ToString().PadRight(10, ' '));
                                }
                            }
                        }
                    }

                    if (target.Equals("autocoe"))
                    {
                        totalUpserts = GenerateDLPImpactProfile(targetEnvironmentId, service, deskFlows, usedActions);
                    }
                    else
                    {
                        GenerateDesktopFlowCSV(rootExportFolder,
                            deskFlows,
                            usedActions,
                            deskFlowSessions,
                            exportRuns,
                            recordCount == batchSize ? true : false);
                    }

                    WriteToProcessingLog(rootExportFolder, deskFlows.Select(d => d.workflowid).ToList(), continueFlag, processedDesktopFlowIdsFromLog);

                    deskFlows = new List<DesktopFlowDefinition>();
                    deskFlowSessions = new List<DesktopFlowSession>();
                    usedActions = new List<DesktopFlowActionUsage>();

                    // Check for more records, if it returns 1.
                    if (desktopFlowCollection.MoreRecords)
                    {
                        // Increment the page number to retrieve the next page.
                        pageNumber++;

                        // Set the paging cookie to the paging cookie returned from current results.                            
                        pagingCookie = desktopFlowCollection.PagingCookie;
                    }
                    else
                    {
                        if(target.Equals("csv"))
                            Process.Start("explorer.exe", Path.Combine(rootExportFolder, "Reporting"));

                        // If no more records in the result nodes, exit the loop.
                        break;
                    }
                }

                Console.WriteLine($"\nEnd Time: {DateTime.Now.ToLongTimeString()}\n");
                Console.WriteLine($"\n################################################################");
                Console.WriteLine($"   Summary Statistics:\n");
                Console.WriteLine($"     -  Export Location:          {rootExportFolder}");
                Console.WriteLine($"     -  Total Desktop flows:      {recordCount}");
                Console.WriteLine($"     -  Total Makers:             {makers.Count}");
                Console.WriteLine($"     -  Total Actions:            {totalActionsUsed}");
                Console.WriteLine($"     -  Total Action Occurrences: {totalOccurrences}");
                if (target.Equals("autocoe"))
                    Console.WriteLine($"     -  Total Dataverse Upserts:  {totalUpserts}");
                Console.WriteLine($"\n################################################################\n");

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
        /// GenerateDLPImpactProfile
        /// </summary>
        /// <param name="targetEnvironmentId"></param>
        /// <param name="service"></param>
        /// <param name="deskFlows"></param>
        /// <param name="usedActions"></param>
        /// <returns></returns>
        internal static int GenerateDLPImpactProfile(string targetEnvironmentId, ServiceClient service, List<DesktopFlowDefinition> deskFlows, List<DesktopFlowActionUsage> usedActions)
        {
            int totalUpserts = 0;

            Console.WriteLine("\n******************************************************\nUpserting {0} Desktop flow definitions in Dataverse...\n******************************************************\n", deskFlows.Count);

            foreach (DesktopFlowDefinition dfs in deskFlows)
            {
                Guid desktopFlowDefId = Guid.Empty;

                QueryExpression query = new QueryExpression(autocoe_DesktopFlowDefinition.EntityLogicalName);
                query.ColumnSet.AddColumns("autocoe_desktopflow", "autocoe_desktopflowdefinitionid", "autocoe_environmentid", "autocoe_owneremail", "autocoe_ownername");
                query.Criteria.AddCondition(new ConditionExpression("autocoe_desktopflow", ConditionOperator.Equal, Guid.Parse(dfs.workflowid)));

                var dfdRec = service.RetrieveMultiple(query).Entities.FirstOrDefault();

                autocoe_DesktopFlowDefinition dfd = null;

                if (dfdRec != null)
                    dfd = new autocoe_DesktopFlowDefinition() { Id = dfdRec.Id };
                else
                    dfd = new autocoe_DesktopFlowDefinition();

                try
                {
                    // Only upsert in case we have a new or changed Desktop flow definition
                    if (dfdRec == null || ((EntityReference)dfdRec["autocoe_desktopflow"]).Id != new EntityReference("workflow", Guid.Parse(dfs.workflowid)).Id ||
                        dfdRec["autocoe_environmentid"].ToString() != targetEnvironmentId ||
                        dfdRec["autocoe_owneremail"].ToString() != dfs.internalemailaddress ||
                        dfdRec["autocoe_ownername"].ToString() != dfs.fullname)
                    {
                        dfd.autocoe_DesktopFlow = new EntityReference("workflow", Guid.Parse(dfs.workflowid));
                        dfd.autocoe_EnvironmentId = targetEnvironmentId;
                        dfd.autocoe_OwnerEmail = dfs.internalemailaddress;
                        dfd.autocoe_OwnerName = dfs.fullname;

                        UpsertRequest ur = new UpsertRequest() { Target = dfd };

                        Console.WriteLine("Upserting Desktop flow definition from {0}", dfs.name);

                        var response = (UpsertResponse)service.Execute(ur);

                        desktopFlowDefId = response.Target.Id;

                    }
                    else
                    {
                        desktopFlowDefId = dfdRec.Id;
                    }

                    var desktopFlowActions = usedActions.Where(a => a.desktopflowid == dfs.workflowid).ToList();

                    ExecuteMultipleRequest requestWithoutResults = new ExecuteMultipleRequest()
                    {
                        // Assign settings that define execution behavior
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = false
                        },
                        // Create an empty organization request collection.
                        Requests = new OrganizationRequestCollection()
                    };

                    QueryExpression dfdQuery = new QueryExpression(autocoe_DesktopFlowDLPImpactProfile.EntityLogicalName);
                    dfdQuery.ColumnSet.AllColumns = true;
                    dfdQuery.Criteria.AddCondition("autocoe_desktopflowdefinition", ConditionOperator.Equal, dfd.Id);

                    var dlpProfs = service.RetrieveMultiple(dfdQuery).Entities.Select(a => a.ToEntity<autocoe_DesktopFlowDLPImpactProfile>()).ToList();

                    foreach (DesktopFlowActionUsage act in desktopFlowActions)
                    {
                        if (dlpProfs.Where(
                            d => d.autocoe_DesktopFlowId == act.desktopflowid &&
                            d.autocoe_SelectorId == act.selectorid &&
                            d.autocoe_ModuleName == act.modulename &&
                            d.autocoe_ModuleSource == act.modulesource &&
                            d.autocoe_OccurrenceCount == act.occurrencecount).FirstOrDefault() != null)
                            continue; // Nothing change, so no need to upsert anything

                        KeyAttributeCollection keyColl = new KeyAttributeCollection();
                        keyColl.Add("autocoe_desktopflowid", act.desktopflowid);
                        keyColl.Add("autocoe_selectorid", act.selectorid);

                        //use alternate key for DLP Profile record
                        autocoe_DesktopFlowDLPImpactProfile dlpProfile = new Entity(autocoe_DesktopFlowDLPImpactProfile.EntityLogicalName, keyColl).ToEntity<autocoe_DesktopFlowDLPImpactProfile>();

                        dlpProfile.autocoe_ActionRecordId = act.desktopflowactionid;
                        dlpProfile.autocoe_ActionName = act.actionname;
                        dlpProfile.autocoe_ModuleSource = act.modulesource;
                        dlpProfile.autocoe_ModuleName = act.modulename;
                        dlpProfile.autocoe_SelectorId = act.selectorid;
                        dlpProfile.autocoe_DesktopFlowName = act.desktopflowname;
                        dlpProfile.autocoe_EnvironmentId = targetEnvironmentId;
                        dlpProfile.autocoe_OccurrenceCount = act.occurrencecount;
                        dlpProfile.autocoe_DesktopFlowDefinition = dfd.ToEntityReference();

                        requestWithoutResults.Requests.Add(new UpsertRequest() { Target = dlpProfile });
                    }

                    if (requestWithoutResults.Requests.Count > 0)
                    {
                        Console.WriteLine("Upserting {0} Desktop flow DLP impact profiles for {1}", requestWithoutResults.Requests.Count, dfs.name);

                        // Execute all the requests in the request collection using a single method call.
                        ExecuteMultipleResponse responseWithoutResults =
                            (ExecuteMultipleResponse)service.Execute(requestWithoutResults);

                        totalUpserts += requestWithoutResults.Requests.Count;
                    }
                    else
                    {
                        Console.WriteLine($"Nothing to upsert for {dfs.name} this time...");
                    }
                }
                // Catch any service fault exceptions that Microsoft Dataverse throws.
                catch (FaultException<OrganizationServiceFault>)
                {
                    throw;
                }
            }

            return totalUpserts;
        }

        /// <summary>
        /// PrepareAutoCoEUpdate
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        static List<DesktopFlowAction> PrepareAutoCoEUpdate(ServiceClient service,string rootExportFolder, bool continueFlag)
        {
            string reportingPath = Path.Combine(rootExportFolder, "Reporting");

            try
            {
                if (!Directory.Exists(rootExportFolder))
                {
                    Directory.CreateDirectory(rootExportFolder);
                    Directory.CreateDirectory(reportingPath);
                }

                FileInfo[] desktopFlowExportFileList = new DirectoryInfo(reportingPath).GetFiles("DesktopFlow*.csv");

                foreach (FileInfo file in desktopFlowExportFileList)
                    file.Delete();

                FileInfo[] processingLog = new DirectoryInfo(reportingPath).GetFiles(PROSESSING_LOG_FILENAME);

                if (!continueFlag)
                {
                    if (processingLog != null && processingLog.Length == 1)
                        processingLog.First().Delete();
                }
                else
                {
                    if (processingLog != null && processingLog.Length == 0)
                        using (File.Create(Path.Combine(reportingPath, PROSESSING_LOG_FILENAME))) { };
                }
            }
            catch (IOException ioExp)
            {
                Console.WriteLine(ioExp.Message);
            }

            List<autocoe_DesktopFlowAction> desktopFlowActionList;
            List<DesktopFlowAction> deskFlowActions = new List<DesktopFlowAction>();

            QueryExpression dfaQuery = new QueryExpression(autocoe_DesktopFlowAction.EntityLogicalName);
            dfaQuery.ColumnSet.AllColumns = true;

            desktopFlowActionList = service.RetrieveMultiple(dfaQuery).Entities.Select(a => a.ToEntity<autocoe_DesktopFlowAction>()).ToList();

            foreach (autocoe_DesktopFlowAction action in desktopFlowActionList)
            {
                deskFlowActions.Add(new DesktopFlowAction
                {
                    desktopflowactionid = action.Id.ToString(),
                    actionname = action.autocoe_ActionName,
                    dlpsupport = action.autocoe_DLPSupport,
                    moduledisplayname = action.autocoe_ModuleDisplayName,
                    modulename = action.autocoe_ModuleName,
                    modulesource = action.autocoe_ModuleSource,
                    selectorid = action.autocoe_SelectorId
                });
            }

            return deskFlowActions;
        }

        /// <summary>
        /// PrepareDesktopFlowSessionCSVExport
        /// </summary>
        /// <param name="service"></param>
        /// <param name="desktopFlowDef"></param>
        /// <param name="desktopFlowSessions"></param>
        internal static void PrepareDesktopFlowSessionCSVExport(ServiceClient service, ref DesktopFlowDefinition desktopFlowDef, ref List<DesktopFlowSession> desktopFlowSessions)
        {
            // Set the number of records per page to retrieve.
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
                                                    <filter>
                                                      <condition attribute='regardingobjectid' operator='eq' value='" + desktopFlowDef.workflowid + @"'/>
                                                    </filter>
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

                foreach (var s in desktopFlowSessionCollection.Entities)
                {
                    if (isLatestRun)
                        lastRundate = s.Contains("createdon") ? DateTime.Parse(s.Attributes["createdon"].ToString()).ToString("G", DateTimeFormatInfo.InvariantInfo) : "";

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
                        createdon = s.Contains("createdon") ? DateTime.Parse(s.Attributes["createdon"].ToString()).ToString("G",DateTimeFormatInfo.InvariantInfo) : "",
                        machinegroupidname = s.Contains("machinegroupid") ? s.FormattedValues["machinegroupid"].ToString() : "",
                        machineidname = s.Contains("machineid") ? s.FormattedValues["machineid"].ToString() : "",
                        statuscode = s.Contains("statuscode") ? s.FormattedValues["statuscode"].ToString() : "",
                        flowsessionid = s.Contains("flowsessionid") ? s.Attributes["flowsessionid"].ToString() : "",
                        completedon = s.Contains("completedon") ? DateTime.Parse(s.Attributes["completedon"].ToString()).ToString("G",DateTimeFormatInfo.InvariantInfo) : "",
                        errormessage = s.Contains("errormessage") ? s.Attributes["errormessage"].ToString() : "",
                        gateway = s.Contains("gateway") ? s.Attributes["gateway"].ToString() : "",
                        startedon = s.Contains("startedon") ? DateTime.Parse(s.Attributes["startedon"].ToString()).ToString("G",DateTimeFormatInfo.InvariantInfo) : "",
                    });

                    if (s.FormattedValues["statuscode"].ToString().Equals("Failed"))
                        totalFailures++;
                    else
                        totalSuccesses++;

                    totalSessions++;

                    isLatestRun = false;
                }

                // Check for morerecords, if it returns 1.
                if (desktopFlowSessionCollection.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = desktopFlowSessionCollection.PagingCookie;
                }
                else
                {
                    desktopFlowDef.totalruns = totalSessions;
                    desktopFlowDef.failedruns = totalFailures;
                    desktopFlowDef.succeededruns = totalSuccesses;
                    desktopFlowDef.lastrundate = lastRundate;

                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }
        }


        /// <summary>
        /// PrepareCSVExport
        /// </summary>
        /// <param name="rootExportFolder"></param>
        /// <param name="continueFlag"></param>
        /// <returns></returns>
        internal static List<DesktopFlowAction> PrepareCSVExport(string rootExportFolder, bool continueFlag)
        {
            string reportingPath = Path.Combine(rootExportFolder, "Reporting");
            
            try
            {
                if (!Directory.Exists(rootExportFolder))
                {
                    Directory.CreateDirectory(rootExportFolder);
                    Directory.CreateDirectory(reportingPath);
                }

                FileInfo[] desktopFlowExportFileList = new DirectoryInfo(reportingPath).GetFiles("DesktopFlow*.csv");

                foreach (FileInfo file in desktopFlowExportFileList) 
                    file.Delete();

                FileInfo[] processingLog = new DirectoryInfo(reportingPath).GetFiles(PROSESSING_LOG_FILENAME);

                if (!continueFlag) 
                {
                    if (processingLog != null && processingLog.Length == 1)
                        processingLog.First().Delete();
                }
                else
                {
                    if (processingLog != null || processingLog.Length == 0)
                        using (File.Create(Path.Combine(reportingPath, PROSESSING_LOG_FILENAME))) { };
                }
            }
            catch (IOException ioExp)
            {
                Console.WriteLine(ioExp.Message);
            }

            List<DesktopFlowAction> deskFlowActions = new List<DesktopFlowAction>();

            // Get the Desktop flow action list from a referenced CSV files and load its context to a custom list of objects
            using (StreamReader sr = new StreamReader(Path.Combine(rootExportFolder, "Reporting", PAD_ACTIONLIST_FILENAME)))
            {
                DesktopFlowModule modules = JsonConvert.DeserializeObject<DesktopFlowModule>(sr.ReadToEnd());

                foreach (Module mod in modules.Modules)
                {
                    foreach (Action act in mod.Actions)
                    {
                        foreach (string selector in act.SelectorIds)
                        {
                            deskFlowActions.Add(new DesktopFlowAction
                            {
                                actionname = act.FriendlyName,
                                desktopflowactionid = act.Id,
                                dlpsupport = act.DLPSupport,
                                moduledisplayname = mod.FriendlyName,
                                modulename = mod.Id,
                                modulesource = mod.ModuleSource,
                                selectorid = selector
                            });
                        }
                    }
                }
            }
            return deskFlowActions;
        }



        /// <summary>
        /// GenerateDesktopFlowCSV
        /// </summary>
        /// <param name="rootExportFolder"></param>
        /// <param name="deskFlows"></param>
        /// <param name="usedActions"></param>
        /// <param name="desktopFlowSessions"></param>
        /// <param name="exportRuns"></param>
        /// <param name="isFirstRecordSet"></param>
        internal static void GenerateDesktopFlowCSV(string rootExportFolder, List<DesktopFlowDefinition> deskFlows, List<DesktopFlowActionUsage> usedActions, List<DesktopFlowSession> desktopFlowSessions, bool exportRuns, bool isFirstRecordSet)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                NewLine = Environment.NewLine,
                HasHeaderRecord = isFirstRecordSet
            };

            string reportingPath = Path.Combine(rootExportFolder, "Reporting");

            using (var desktopFlowWriter = new StreamWriter(Path.Combine(reportingPath, DESKOPTFLOW_FILENAME),true))
            using (var desktopFlowActionWriter = new StreamWriter(Path.Combine(reportingPath, DESKOPTFLOW_ACTION_FILENAME), true))
            using (var csvDesktopFlows = new CsvWriter(desktopFlowWriter, config))
            using (var csvDesktopFlowActions = new CsvWriter(desktopFlowActionWriter, config))
            {
                csvDesktopFlows.WriteRecords(deskFlows);
                csvDesktopFlowActions.WriteRecords(usedActions);
            }

            if (exportRuns)
            {
                using (var desktopFlowSessionWriter = new StreamWriter(Path.Combine(reportingPath, DESKOPTFLOW_SESSION_FILENAME), true))
                using (var csvDesktopFlowSessions = new CsvWriter(desktopFlowSessionWriter, config))
                {
                    csvDesktopFlowSessions.WriteRecords(desktopFlowSessions);
                }
            }
        }



        /// <summary>
        /// WriteToProcessingLog
        /// </summary>
        /// <param name="rootExportFolder"></param>
        /// <param name="desktopFlowIds"></param>
        /// <param name="continueFlag"></param>
        /// <param name="existingProcessedDesktopFlowIds"></param>
        internal static void WriteToProcessingLog(string rootExportFolder, List<string> desktopFlowIds, bool continueFlag, List<string> existingProcessedDesktopFlowIds)
        {
            try { 
                string reportingPath = Path.Combine(rootExportFolder, "Reporting");

                using (FileStream fs = new FileStream(Path.Combine(reportingPath, PROSESSING_LOG_FILENAME), FileMode.Append, FileAccess.Write))
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    if (continueFlag && existingProcessedDesktopFlowIds.Count > 0)
                    {
                        var uniqueIdList = desktopFlowIds.Union(existingProcessedDesktopFlowIds).ToList();
                        uniqueIdList.ForEach(x => sw.WriteLine(x));
                    }
                    else
                        desktopFlowIds.ForEach(x => sw.WriteLine(x));
                }
            }
            catch (IOException ioExp)
            {
                Console.WriteLine(ioExp.Message);
            }
        }


        /// <summary>
        /// ReadDesktopFlowIdsProcessingLog
        /// </summary>
        /// <param name="rootExportFolder"></param>
        /// <returns></returns>
        internal static List<string> ReadDesktopFlowIdsProcessingLog(string rootExportFolder)
        {
            string processingFilePath = Path.Combine(rootExportFolder, "Reporting", PROSESSING_LOG_FILENAME);

            if (File.Exists(processingFilePath))
            {
                List<string> resultsArray = new List<string>();
                string[] text = File.ReadAllLines(processingFilePath);

                foreach (string line in text) 
                    resultsArray.Add(line.Trim());
                
                return resultsArray;
            }
            return new List<string>();
        }
    }
}