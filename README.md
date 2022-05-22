# RPA CLI

### Overview

The RPA CLI command-line tool can be used to extract Desktop flow definitions, actions, run logs and more.

## Pre-requisites
- You will need a highly privileged Dataverse environment user account (ie user with Sys Admin role or similar) in order to extract desktop flow and session information environment(s) wide.
- Before using the **AutoCoE** target option, make sure you have solution version **0_0_1_4** or higher installed.

### Getting started
![RPACLI](https://user-images.githubusercontent.com/10453029/169645497-af04a2d8-867b-4ea4-8c3b-9eeb61e5a451.gif)

### RPACLIE v0.2 includes a new diff function
![Untitled Project](https://user-images.githubusercontent.com/10453029/169693830-0ec13d88-b415-4016-8dfe-b2b1b818243c.gif)

#### Run it

1. Download the [latest release](https://github.com/rpapostolis/rpa-cli/releases)
2. Unzip RPACLI_v[X].[X]_win-x[XX].zip archive
3. Open a command prompt and navigate to the previously extracted folder
4. Run one of the [commands below](#example-usage-exporting-desktop-flow-metadata-and-action-statistics-to-csv)
5. If you've chosen CSV as output target you can navigate to the **Reporting** folder and open either the **[Desktop Flow Action Analysis without Runs.pbit](https://github.com/rpapostolis/rpa-cli/blob/main/RPACLI/Reporting/Desktop%20Flow%20Action%20Analysis%20without%20Runs.pbit)** or **[Desktop Flow Action Analysis with Runs.pbit](https://github.com/rpapostolis/rpa-cli/blob/main/RPACLI/Reporting/Desktop%20Flow%20Action%20Analysis%20with%20Runs.pbit)** Power BI template. Once it opens you're asked to provide the Reporting folder and here simply enter the same folder path where the Power BI template is located.
If you have chosen AutoCoE. The minimum solution version installed must be **0_0_1_4**.

#### Build it

1. Clone this repo with `git clone https://github.com/rpapostolis/rpa-cli.git`
2. Open in VS 2022 (or VS Code with additional launch settings steps)
3. Build and run (with desired launchSetting config)

#### Use it

<pre>
<b>.\RPACLI.exe --help</b></pre>

```
Description:

                           _____  _____           _____ _      _____
                          |  __ \|  __ \ /\      / ____| |    |_   _|
                          | |__) | |__) /  \    | |    | |      | |
                          |  _  /|  ___/ /\ \   | |    | |      | |
                          | | \ \| |  / ____ \  | |____| |____ _| |_
                          |_|  \_\_| /_/    \_\  \_____|______|_____|

  The RPA CLI command tool can be used to extract Desktop flow definitions, actions,
  run logs and more.

  Author:
  Apostolis Papaioannou

  Disclaimer:
  This command-line tool is unsupported, experimental, NOT an official Microsoft product
  and is provided 'as is' without support and warranty of any kind, either express or
  implied, including but not limited to warranties of merchantability, fitness for a
  particular purpose, and non-infringement. In no event shall the authors or copyright
  holders be liable for any claims, damages or other liabilities, whether in contract, tort
  or otherwise, arising out of, arising out of or in connection with the Software or the
  use or other dealings with the Software.

Usage:
  RPACLI [command] [options]

Options:
  --version       Show version information
  -?, -h, --help  Show help and usage information

Commands:
  get-actions          This command scans all Desktop flows that the specified user has
                       access to in an environment and extracts core Desktop flow metadata
                       together with detailed action usage telemetry.

                       With the new PAD action-based DLP preview feature, administrators and
                       CoE teams can now define which action modules and individual actions
                       can be used as part of Desktop flows created with Power Automate for
                       Desktop. In the case of policy violations (ie VBScript was allowed
                       but is now being restricted by DLP policy), the platform notifies the
                       maker that this action is not allowed anymore, hence the flow will be
                       suspended.

                       This is a great governance improvement and fits perfectly with other
                       DLP features across cloud flows and Power Apps. To minimize the
                       potential impact on already deployed bots, you can use this new CLI
                       tool to extract Desktop flow metadata along with action usage
                       telemetry. The results can then be visualized in Power BI Desktop,
                       where proactive impact analysis can be performed on planned DLP
                       changes.
  get-desktop-runs     This command retrieves flow runs for a specific Desktop flow and
                       timeframe.
  clone-desktop-flows  This command clones copies of Desktop flows from a template Desktop
                       flow.
  diff-desktop-flows   This command detects differences between two Desktop flows and shows
                       an action usage diff in VS Code.

```  

<pre><b>.\RPACLI.exe get-actions --help</b></pre>

```
Description:
  This command scans all desktop flows that the specified user has access to in an environment and
  extracts core desktop flow metadata together with detailed action usage telemetry.

  With the new PAD action-based DLP preview feature, administrators and CoE teams can now define
  which action modules and individual actions can be used as part of desktop flows created with
  Power Automate for Desktop. In the case of policy violations (ie VBScript was allowed but is now
  being restricted by DLP policy), the platform notifies the maker that this action is not allowed
  anymore, hence saving the flow is not possible.

  This is a great governance improvement and fits perfectly with other DLP features across cloud
  flows and Power Apps. To minimize the potential impact on already deployed bots, you can use this
  new CLI tool to extract desktop flow metadata along with action usage telemetry. The results can
  then be visualized in Power BI Desktop, where proactive impact analysis can be performed on
  planned DLP changes.

Usage:
  RPACLI get-actions [options]

Options:
  -u, --username <username> (REQUIRED)              Username of a highly privileged account under
                                                    which the process will run.
  -t, --target <AutoCoE|CSV> (REQUIRED)             The output target the extrated Desktop flow
                                                    action stats should be stored in. Available
                                                    media types are CSV and AutoCoE (automation CoE
                                                    starter kit). If you select AutoCoE, then the
                                                    extracted action stats will be stored in
                                                    'Automation CoE Main' tables called
                                                    autocoe_DesktopFlowDefinition and
                                                    autocoe_DesktopFlowDLPImpactProfile. [default:
                                                    CSV]
  -s, --sourceenvironmentinstanceurl                The source environment instance url from which
  <sourceenvironmentinstanceurl> (REQUIRED)         you want to extract Desktop flows from ie
                                                    https://mysourceorg.crm.dynamics.com .
  -er, --exportruns                                 Do you want to export desktop flow sessions as
                                                    well? Please note that the processing time
                                                    takes significantly longer if you're enabling
                                                    this option. [default: False]
  -b, --batchsize <batchsize>                       The number of desktop flows to be processed per
                                                    data page. [default: 10]
  -c, --continue                                    Do you want to continue where the process has
                                                    stopped during the previous run? [default:
                                                    False]
  -te, --targetenvironmentinstanceurl               The target environment instance url of the
  <targetenvironmentinstanceurl>                    environment in which you want store the
                                                    extracted Desktop flow telemetry in ie
                                                    https://mytargetorg.crm.dynamics.com .
  -ti, --targetenvironmentid <targetenvironmentid>  The target environment Id of the environment in
                                                    which you want store the extracted Desktop flow
                                                    telemetry in ie
                                                    7a74e5d3-7ff3-4bc4-b7bb-37b1428c9a1aa.
  -f, --fgcolor                                     Foreground color of text displayed on the
  <Black|Blue|Cyan|DarkBlue|DarkCyan|DarkGray|Dark  console. [default: Cyan]
  Green|DarkMagenta|DarkRed|DarkYellow|Gray|Green|
  Magenta|Red|White|Yellow>
  -x, --light-mode                                  Background color of text displayed on the
                                                    console: default is black, light mode is white.
  -?, -h, --help                                    Show help and usage information

```

<pre><b>.\RPACLI get-desktop-runs</b></pre>

```
Description:
  This command retrieves flow runs for a specific desktop flow and timeframe.

Usage:
  RPACLI get-desktop-runs [options]

Options:
  -u, --username <username> (REQUIRED)          Username of a highly privileged account
                                                under which the process will run.
  -s, --sourceenvironmentinstanceurl            The source environment instance url from
  <sourceenvironmentinstanceurl> (REQUIRED)     which you want to extract Desktop flows
                                                from ie
                                                https://mysourceorg.crm.dynamics.com .
  -d, --desktopflowid <desktopflowid>           The desktop flow id which you wish to
  (REQUIRED)                                    process data for.
  -t, --timeframe <timeframe> (REQUIRED)        The timeframe in days for which the run
                                                logs should be retrieved. 0 = All history
                                                [default: 0]
  -e, --export (REQUIRED)                       A flag that indicates if the flow sessions
                                                should be exported to CSV. [default: False]
  -f, --fgcolor                                 Foreground color of text displayed on the
  <Black|Blue|Cyan|DarkBlue|DarkCyan|DarkGray|  console. [default: Cyan]
  DarkGreen|DarkMagenta|DarkRed|DarkYellow|Gra
  y|Green|Magenta|Red|White|Yellow>
  -x, --light-mode                              Background color of text displayed on the
                                                console: default is black, light mode is
                                                white.
  -?, -h, --help                                Show help and usage information


```

<pre><b>.\RPACLI clone-desktop-flows</b></pre>

```
Description:
  This command clones copies of Desktop flows from a template desktop flow.

Usage:
  RPACLI clone-desktop-flows [options]

Options:
  -u, --username <username> (REQUIRED)              Username of a highly privileged account under
                                                    which the process will run.
  -s, --sourceenvironmentinstanceurl                The source environment instance url from which
  <sourceenvironmentinstanceurl> (REQUIRED)         you want to extract Desktop flows from ie
                                                    https://mysourceorg.crm.dynamics.com .
  -d, --desktopflowid <desktopflowid> (REQUIRED)    The desktop flow id which you wish to process
                                                    data for.
  -c, --count <count> (REQUIRED)                    The number of desktop flow copies to generate
                                                    [default: 0]
  -b, --batchsize <batchsize>                       The number of desktop flows to be processed per
                                                    data page. [default: 10]
  -f, --fgcolor                                     Foreground color of text displayed on the
  <Black|Blue|Cyan|DarkBlue|DarkCyan|DarkGray|Dark  console. [default: Cyan]
  Green|DarkMagenta|DarkRed|DarkYellow|Gray|Green|
  Magenta|Red|White|Yellow>
  -x, --light-mode                                  Background color of text displayed on the
                                                    console: default is black, light mode is white.
  -?, -h, --help                                    Show help and usage information
  
```  


<pre><b>.\RPACLI.exe diff-desktop-flows --help</b></pre>

```
Description:
  This command detects differences between two Desktop flows and shows an action usage diff
  in VS Code.

Usage:
  RPACLI diff-desktop-flows [options]

Options:
  -u, --username <username> (REQUIRED)          Username of a highly privileged account
                                                under which the process will run.
  -source1, --file1sourceenvironment            The source environment instance url of the
  <file1sourceenvironment> (REQUIRED)           first Desktop flow you want compare with the
                                                second one ie
                                                https://mysourceorg.crm.dynamics.com .
  -source2, --file2sourceenvironment            The source environment instance url of the
  <file2sourceenvironment>                      second Desktop flow you want compare with
                                                the first one ie
                                                https://mysourceorg.crm.dynamics.com .
  -s, --sameenvironment                         A flag that indicates if the Desktop flows
                                                to be compared are from the same Dataverse
                                                envrionment. [default: True]
  -flowid1, --desktopflowid1 <desktopflowid1>   The first Desktop flow id that you want to
  (REQUIRED)                                    compare the second one with.
  -flowid2, --desktopflowid2 <desktopflowid2>   The second Desktop flow id that you want to
  (REQUIRED)                                    compare the first one with.
  -f, --fgcolor                                 Foreground color of text displayed on the
  <Black|Blue|Cyan|DarkBlue|DarkCyan|DarkGray|  console. [default: Cyan]
  DarkGreen|DarkMagenta|DarkRed|DarkYellow|Gra
  y|Green|Magenta|Red|White|Yellow>
  -x, --light-mode                              Background color of text displayed on the
                                                console: default is black, light mode is
                                                white.
  -?, -h, --help                                Show help and usage information
  
```

#### Example usage: Exporting Desktop Flow metadata and action statistics to CSV

<pre><b>.\RPACLI.exe get-actions</b> -u myname@mytenant.onmicrosoft.com -b 10 -t CSV 
-s https://[myorg].crm[region].dynamics.com -er true -c true
</pre>

**Using default options, export without runs (flow sessions)**

<pre><b>.\RPACLI.exe get-actions</b> -u myname@mytenant.onmicrosoft.com -s https://[myorg].crm[region].dynamics.com</pre>

**Export with runs (flow sessions)**

<pre><b>.\RPACLI.exe get-actions</b> -u myname@mytenant.onmicrosoft.com -s https://[myorg].crm[region].dynamics.com 
-er true</pre>

Adding -c continues from where the export left off. Desktop flow Ids are stored inside the generated file **_processingLog.txt**.
Using -c will only export Desktop flows that are new (not written to the **_processingLog.txt**).

<pre><b>.\RPACLI.exe get-actions</b> -u myname@mytenant.onmicrosoft.com -s https://[myorg].crm[region].dynamics.com 
-c true</pre>

Use -b to change the batch size. The more Desktop flows to export the longer the processing.
> Important
> For exporting large amounts of Desktop flows, stay around the default batch size (10).

<pre><b>.\RPACLI.exe get-actions</b> -u myname@mytenant.onmicrosoft.com -s https://[myorg].crm[region].dynamics.com 
-c true -b 15</pre>

#### Example usage: Bootstrapping or updating Desktop Flow metadata and action statistics in the Automation CoE Starter Kit

<pre><b>.\RPACLI.exe get-actions</b> -u myname@mytenant.onmicrosoft.com -b 10 -t AutoCoE 
-s https://[myorgurl].crm.dynamics.com -te https://[myorg].crm[region].dynamics.com 
-ti 7a74e5d3-7ff3-4bc4-b7bb-37b1428c5d4a -er true -c true
</pre>

#### Example usage: Get Desktop flow inventory of an environment (Same source and target environment)

<pre><b>.\RPACLI.exe get-actions</b> --username myname@mytenant.onmicrosoft.com 
--sourceenvironmentinstanceurl https://[myorg].crm[region].dynamics.com/ --target AutoCoE 
--targetenvironmentinstanceurl https://[myorg].crm[region].dynamics.com/ 
--targetenvironmentid 7a74e5d3-7ff3-4bc4-b7bb-37b1428c5d4a --exportruns true -c true</pre>

#### Example usage: Exporting Desktop Flow Sessions (only) to CSV

<pre><b>.\RPACLI.exe get-desktop-runs</b> -u myname@mytenant.onmicrosoft.com -e true 
-s https://[myorg].crm[region].dynamics.com -d 17eb9a9c-3198-4a99-ac99-da5ea0da2cda -t 21
</pre>

#### Example usage: Clone Desktop Flows based on a template Desktop Flow

<pre><b>.\RPACLI.exe clone-desktop-flows</b> -u myname@mytenant.onmicrosoft.com 
-s https://[myorg].crm[region].dynamics.com -d 39069fe2-e299-4cf9-99c3-999aeafcf999 
-c 100 -b 10
</pre>

#### Example usage: Compare two Desktop flows from two different environments and show their diffs in VS Code

<pre><b>.\RPACLI.exe diff-desktop-flows</b> -u myname@mytenant.onmicrosoft.com 
diff-desktop-flows -u apostolis@pasandbox.onmicrosoft.com -s false 
-source1 https://[myorg1].crm[region].dynamics.com 
-source2 https://[myorg2].crm[region].dynamics.com 
-flowid1 999de370-110e-40c6-a276-bd83e89bc07b 
-flowid2 da736383-3c2e-45b7-81ef-4f131a8e404b
</pre>

## Solution software packages used

Following nuget packages are used, please check their individual license terms:

CsvHelper  
- Nuget package https://www.nuget.org/packages/CsvHelper/
- Project https://joshclose.github.io/CsvHelper/

Microsoft.PowerPlatform.Dataverse.Client
- Nuget https://www.nuget.org/packages/Microsoft.PowerPlatform.Dataverse.Client/
- Project https://github.com/microsoft/PowerPlatform-DataverseServiceClient/

System.CommandLine
- Nuget https://www.nuget.org/packages/System.CommandLine/
- Project https://github.com/dotnet/command-line-api

Newtonsoft 
- Nuget package https://www.nuget.org/packages/Newtonsoft.Json/
- Project https://www.newtonsoft.com/json

### Important note

The [modules.json](https://github.com/rpapostolis/rpa-cli/blob/main/RPACLI/Reporting/modules.json) file that is part of the solution lists Power Automate for Desktop actions that were available in the product at the time of this repo's release. However, it is important to note that these actions may change at any time without prior notice and thus affect RPA CLI functionality. There is currently no intention to update this repo on a regular basis, so use is at your own risk. Please make sure you’ve also noted the [Disclaimer](#disclaimer) below.

## License

This project's base code is licensed under the [MIT license](https://github.com/rpapostolis/rpa-cli/blob/main/LICENSE).
## Disclaimer

**This command-line tool is unsupported, experimental, NOT an official Microsoft product and is provided 'as is' without support and warranty of any kind**, either express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, and non-infringement. In no event shall the authors or copyright holders be liable for any claims, damages or other liabilities, whether in contract, tort or otherwise, arising out of, arising out of or in connection with the Software or the use or other dealings with the Software.
