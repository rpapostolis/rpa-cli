using System.CommandLine;
using System.Diagnostics;
using RPACLI.Handlers;

string rootExportFolder = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);

#region Command: rootCommand

// Add a root command:
var rootCommand = new RootCommand("RPA CLI");
 

rootCommand.Description = @"                                               
                         _____  _____           _____ _      _____ 
                        |  __ \|  __ \ /\      / ____| |    |_   _|
                        | |__) | |__) /  \    | |    | |      | |  
                        |  _  /|  ___/ /\ \   | |    | |      | |  
                        | | \ \| |  / ____ \  | |____| |____ _| |_ 
                        |_|  \_\_| /_/    \_\  \_____|______|_____|
                                                            
The RPA CLI command-line tool might be used to extract Desktop flow definitions, actions, run logs and more.

Author:
Apostolis Papaioannou

Disclaimer:
This command-line tool is unsupported, experimental, NOT an official Microsoft product and is provided 'as is' without support and warranty of any kind, either express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, and non-infringement. In no event shall the authors or copyright holders be liable for any claims, damages or other liabilities, whether in contract, tort or otherwise, arising out of, arising out of or in connection with the Software or the use or other dealings with the Software.";

#endregion

#region Command: get-actions 

var sUsername = new Option<string>(
    "--username",
    description: "Username of a highly privileged account under which the process will run.");
var sTarget = new Option<string>(
    "--target",
    getDefaultValue: () => "CSV",
    description: "The output target the extrated Desktop flow action stats should be stored in. Available media types are CSV and AutoCoE (automation CoE starter kit). If you select AutoCoE, then the extracted action stats will be stored in 'Automation CoE Main' tables called autocoe_DesktopFlowDefinition and autocoe_DesktopFlowDLPImpactProfile.")
    .FromAmong("CSV", "AutoCoE");
var sSourceEnvironmentInstanceUrl = new Option<string>(
    "--sourceenvironmentinstanceurl",
    description: "The source environment instance url from which you want to extract Desktop flows from ie https://mysourceorg.crm.dynamics.com .");
var sTargetEnvironmentInstanceUrl = new Option<string>(
    "--targetenvironmentinstanceurl",
    description: "The target environment instance url of the environment in which you want store the extracted Desktop flow telemetry in ie https://mytargetorg.crm.dynamics.com .");
var sTargetEnvironmentId = new Option<string>(
    "--targetenvironmentid",
    description: "The target environment Id of the environment in which you want store the extracted Desktop flow telemetry in ie 7a74e5d3-7ff3-4bc4-b7bb-37b1428c9a1aa.");
var bExportRuns = new Option<bool>(
    "--exportruns",
    getDefaultValue: () => false,
    description: "Do you want to export Desktop flow sessions as well? Please note that the processing time takes significantly longer if you're enabling this option.");
var iBatchSize = new Option<int>(
    "--batchsize",
    getDefaultValue: () => 10,
    description: "The number of Desktop flows to be processed per data page.");
var bContinue = new Option<bool>(
    "--continue",
    getDefaultValue: () => false,
    description: "Do you want to continue where the process has stopped during the previous run?");
var eFgcolorOption = new Option<ConsoleColor>(
    name: "--fgcolor",
    description: "Foreground color of text displayed on the console.",
    getDefaultValue: () => ConsoleColor.Cyan);
var bLightModeOption = new Option<bool>(
    name: "--light-mode",
    description: "Background color of text displayed on the console: default is black, light mode is white.");

sUsername.AddAlias("-u");
sUsername.Arity = ArgumentArity.ExactlyOne;
sUsername.IsRequired = true;
sTarget.AddAlias("-t");
sTarget.Arity = ArgumentArity.ExactlyOne;
sTarget.IsRequired = true;
bExportRuns.AddAlias("-er");
bExportRuns.Arity = ArgumentArity.ZeroOrOne;
bExportRuns.IsRequired = false;
iBatchSize.AddAlias("-b");
iBatchSize.Arity = ArgumentArity.ZeroOrOne;
iBatchSize.IsRequired = false;
bContinue.AddAlias("-c");
bContinue.Arity = ArgumentArity.ZeroOrOne;
bContinue.IsRequired = false;
sSourceEnvironmentInstanceUrl.AddAlias("-s");
sSourceEnvironmentInstanceUrl.Arity = ArgumentArity.ExactlyOne;
sSourceEnvironmentInstanceUrl.IsRequired = true;
sTargetEnvironmentInstanceUrl.AddAlias("-te");
sTargetEnvironmentInstanceUrl.Arity = ArgumentArity.ZeroOrOne;
sTargetEnvironmentInstanceUrl.IsRequired = false;
sTargetEnvironmentId.AddAlias("-ti");
sTargetEnvironmentId.Arity = ArgumentArity.ZeroOrOne;
sTargetEnvironmentId.IsRequired = false;
eFgcolorOption.AddAlias("-f");
eFgcolorOption.IsRequired = false;
bLightModeOption.AddAlias("-x");
bLightModeOption.IsRequired = false;

var getActions = new Command("get-actions", "This command scans all Desktop flows that the specified user has access to in an environment and extracts core Desktop flow metadata together with detailed action usage telemetry")
{
    sUsername,
    sTarget,
    sSourceEnvironmentInstanceUrl,
    bExportRuns,
    iBatchSize,
    bContinue,
    sTargetEnvironmentInstanceUrl,
    sTargetEnvironmentId,
    eFgcolorOption,
    bLightModeOption
};

getActions.Description = @"This command scans all Desktop flows that the specified user has access to in an environment and extracts core Desktop flow metadata together with detailed action usage telemetry.

With the new PAD action-based DLP preview feature, administrators and CoE teams can now define which action modules and individual actions can be used as part of Desktop flows created with Power Automate for Desktop. In the case of policy violations (ie VBScript was allowed but is now being restricted by DLP policy), the platform notifies the maker that this action is not allowed anymore, hence the flow will be suspended.

This is a great governance improvement and fits perfectly with other DLP features across cloud flows and Power Apps. To minimize the potential impact on already deployed bots, you can use this new CLI tool to extract Desktop flow metadata along with action usage telemetry. The results can then be visualized in Power BI Desktop, where proactive impact analysis can be performed on planned DLP changes.";

getActions.SetHandler((string u, string t, string s, bool er, int b, bool c, string te, string i, ConsoleColor fgColor, bool lightMode) =>
{
    Console.BackgroundColor = lightMode ? ConsoleColor.White : ConsoleColor.Black;
    Console.ForegroundColor = fgColor;

    Handler_Get_Actions.GenerateDesktopFlowActionStats(u, t, s, er, b, c, te, i, rootExportFolder);

}, sUsername, sTarget, sSourceEnvironmentInstanceUrl, bExportRuns, iBatchSize, bContinue, sTargetEnvironmentInstanceUrl, sTargetEnvironmentId, eFgcolorOption, bLightModeOption);

rootCommand.AddCommand(getActions);

#endregion

#region Command: get-runs

var sDesktopFlowId = new Option<string>(
    "--desktopflowid",
    description: "The Desktop flow id which you wish to process data for.");
var iTimeframe = new Option<int>(
    "--timeframe",
    getDefaultValue: () => 0,
    description: "The timeframe in days for which the run logs should be retrieved. 0 = All history");
var bExport = new Option<bool>(
    "--export",
    getDefaultValue: () => false,
    description: "A flag that indicates if the flow sessions should be exported to CSV.");

iTimeframe.AddAlias("-t");
iTimeframe.Arity = ArgumentArity.ExactlyOne;
iTimeframe.IsRequired = true;
bExport.AddAlias("-e");
bExport.Arity = ArgumentArity.ExactlyOne;
bExport.IsRequired = true;
sDesktopFlowId.AddAlias("-d");
sDesktopFlowId.Arity = ArgumentArity.ExactlyOne;
sDesktopFlowId.IsRequired = true;

var getDesktopFlowRuns = new Command("get-desktop-runs", "Retrieves Desktop flow runs") { 
    sUsername,
    sSourceEnvironmentInstanceUrl,
    sDesktopFlowId, 
    iTimeframe,
    bExport,
    eFgcolorOption,
    bLightModeOption
};

getDesktopFlowRuns.Description = @"This command retrieves flow runs for a specific Desktop flow and timeframe.";

getDesktopFlowRuns.SetHandler((string u, string s, string d, int t, bool e, ConsoleColor fgColor, bool lightMode) =>
{
    Console.BackgroundColor = lightMode ? ConsoleColor.White : ConsoleColor.Black;
    Console.ForegroundColor = fgColor;

    Handler_Get_Runs.GetDesktopFlowRuns(u, s, d, t, e, rootExportFolder);

}, sUsername, sSourceEnvironmentInstanceUrl, sDesktopFlowId, iTimeframe, bExport, eFgcolorOption, bLightModeOption);

rootCommand.AddCommand(getDesktopFlowRuns);

#endregion

#region Command: clone-desktop-flows

var sCount = new Option<int>(
    "--count",
    getDefaultValue: () => 0,
    description: "The number of Desktop flow copies to generate");

sCount.AddAlias("-c");
sCount.Arity = ArgumentArity.ExactlyOne;
sCount.IsRequired = true;

var cloneDesktopFlows = new Command("clone-desktop-flows", "Clones copies of Desktop flows from a template Desktop flow") {
    sUsername,
    sSourceEnvironmentInstanceUrl,
    sDesktopFlowId, 
    sCount,
    iBatchSize,
    eFgcolorOption,
    bLightModeOption
};

cloneDesktopFlows.Description = @"This command clones copies of Desktop flows from a template Desktop flow.";

cloneDesktopFlows.SetHandler((string u, string s, string d, int c, int b, ConsoleColor fgColor, bool lightMode) =>
{
    Console.BackgroundColor = lightMode ? ConsoleColor.White : ConsoleColor.Black;
    Console.ForegroundColor = fgColor; 
    
    Handler_Clone_Desktop_Flows.CloneDesktopFlows(u, s, d, c, b);

}, sCount, sSourceEnvironmentInstanceUrl, sDesktopFlowId, sCount, iBatchSize, eFgcolorOption, bLightModeOption);

rootCommand.AddCommand(cloneDesktopFlows);

#endregion

#region Command: diff-desktop-flows

var sDesktopFlow1SourceEnvironmentInstanceUrl = new Option<string>(
    "--file1sourceenvironment",
    description: "The source environment instance url of the first Desktop flow you want compare with the second one ie https://mysourceorg.crm.dynamics.com .");

var sDesktopFlow2SourceEnvironmentInstanceUrl = new Option<string>(
    "--file2sourceenvironment",
    description: "The source environment instance url of the second Desktop flow you want compare with the first one ie https://mysourceorg.crm.dynamics.com .");

var bIsFromSameEnvironment = new Option<bool>(
    "--sameenvironment",
    getDefaultValue: () => true,
    description: "A flag that indicates if the Desktop flows to be compared are from the same Dataverse envrionment.");

var sDesktopFlowId1 = new Option<string>(
    "--desktopflowid1",
    description: "The first Desktop flow id that you want to compare the second one with.");

var sDesktopFlowId2 = new Option<string>(
    "--desktopflowid2",
    description: "The second Desktop flow id that you want to compare the first one with.");

sDesktopFlow1SourceEnvironmentInstanceUrl.AddAlias("-source1");
sDesktopFlow1SourceEnvironmentInstanceUrl.Arity = ArgumentArity.ExactlyOne;
sDesktopFlow1SourceEnvironmentInstanceUrl.IsRequired = true;

sDesktopFlow2SourceEnvironmentInstanceUrl.AddAlias("-source2");
sDesktopFlow2SourceEnvironmentInstanceUrl.Arity = ArgumentArity.ExactlyOne;
sDesktopFlow2SourceEnvironmentInstanceUrl.IsRequired = false;

sDesktopFlowId1.AddAlias("-flowid1");
sDesktopFlowId1.Arity = ArgumentArity.ExactlyOne;
sDesktopFlowId1.IsRequired = true;

sDesktopFlowId2.AddAlias("-flowid2");
sDesktopFlowId2.Arity = ArgumentArity.ExactlyOne;
sDesktopFlowId2.IsRequired = true;

bIsFromSameEnvironment.AddAlias("-s");
bIsFromSameEnvironment.Arity = ArgumentArity.ExactlyOne;
bIsFromSameEnvironment.IsRequired = false;

var diffDesktopFlows = new Command("diff-desktop-flows", "Detect differences between two Desktop flows and show an action usage diff in VS Code. This command requires VS Code to be installed and added to the PATH.") {
    sUsername,
    sDesktopFlow1SourceEnvironmentInstanceUrl,
    sDesktopFlow2SourceEnvironmentInstanceUrl,
    bIsFromSameEnvironment,
    sDesktopFlowId1,
    sDesktopFlowId2,
    eFgcolorOption,
    bLightModeOption
};

diffDesktopFlows.Description = @"This command detects differences between two Desktop flows and shows an action usage diff in VS Code.";

diffDesktopFlows.SetHandler((string u, string s1, string s2, bool s, string d1, string d2, ConsoleColor fgColor, bool lightMode) =>
{
    Console.BackgroundColor = lightMode ? ConsoleColor.White : ConsoleColor.Black;
    Console.ForegroundColor = fgColor;

    Handler_Diff_Desktop_Flows.DiffDesktopFlows(u, s1, s2, s, d1, d2);

}, sUsername,
    sDesktopFlow1SourceEnvironmentInstanceUrl,
    sDesktopFlow2SourceEnvironmentInstanceUrl,
    bIsFromSameEnvironment,
    sDesktopFlowId1,
    sDesktopFlowId2,
    eFgcolorOption,
    bLightModeOption);

rootCommand.AddCommand(diffDesktopFlows);

#endregion

//Parse the incoming args and invoke the respective handler
return rootCommand.Invoke(args);