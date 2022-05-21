using System.Text;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
using Newtonsoft.Json;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace RPACLI
{

    /// <summary>
    /// CLIHelper
    /// </summary>
    public static class CLIHelper
    {

        /// <summary>
        /// ConnectToDataverse
        /// </summary>
        /// <param name="username"></param>
        /// <param name="sourceenvinstanceurl"></param>
        /// <returns></returns>
        public static ServiceClient ConnectToDataverse(string username, string sourceenvinstanceurl)
        {
            ServiceClient service = null;

            try
            {
                Console.WriteLine("Connecting to Microsoft Dataverse... Please Check your Web Browser to login.");

                StringBuilder connectionString = new StringBuilder();

                connectionString.Append("AuthType=OAuth;");
                connectionString.Append($"Username={username};");
                connectionString.Append($"Integrated Security=true;");
                connectionString.Append($"Url={sourceenvinstanceurl};");
                connectionString.Append($"AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;");// Standard App Id available globally
                connectionString.Append($"RedirectUri=http://localhost;");
                connectionString.Append($"LoginPrompt=Auto");                           // Promt for interactive Login

                // Try to create via connection string. 
                service = new ServiceClient(connectionString.ToString());
                
                if (service.IsReady)
                {
                    Console.Clear();
                    Console.WriteLine($"...successfully connected to {service.ConnectedOrgFriendlyName}!\n");
                }
                else
                {
                    const string UNABLE_TO_LOGIN_ERROR = "Unable to Login to Microsoft Dataverse";
                    if (service.LastError.Equals(UNABLE_TO_LOGIN_ERROR))
                    {
                        Console.WriteLine("Check the connection string values passed in to the application.");
                        throw new Exception(service.LastError);
                    }
                    else
                    {
                        throw service.LastException;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return service;
        }

        /// <summary>
        /// GetZippedDesktopFlowScript
        /// </summary>
        /// <param name="clientdata"></param>
        /// <returns></returns>
        public static string GetZippedDesktopFlowScript(string clientdata)
        {
            if (!string.IsNullOrEmpty(clientdata))
            {
                return JsonConvert.DeserializeObject<WorkflowClientData>(clientdata).properties.definition.package;
            }
            else
                return string.Empty;
        }

        /// <summary>
        /// UnzipDesktopFlowScript
        /// </summary>
        /// <param name="clientData"></param>
        /// <param name="formatScript"></param>
        /// <returns></returns>
        public static string UnzipDesktopFlowScript(string clientData, bool formatScript)
        {
            if (!string.IsNullOrEmpty(clientData))
            {
                var base64EncodedBytes = System.Convert.FromBase64String(clientData);

                Stream data = new MemoryStream(base64EncodedBytes);
                ZipArchive archive = new ZipArchive(data);

                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    using (StreamReader sr = new StreamReader(entry.Open()))
                    {
                        if (!formatScript)
                        {
                            return sr.ReadToEnd().Replace("\0", "");
                        }
                        else
                        {
                            StringBuilder sb = new StringBuilder();

                            var multiLineText = Regex.Split(sr.ReadToEnd().Replace("\0", ""), "\r\n|\r|\n", RegexOptions.Multiline);

                            foreach (string value in multiLineText)
                            {
                                sb.Append(value);
                                sb.Append(' ');
                            }
                            return sb.ToString();
                        }
                    }
                }
                return "";
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// RemoveVariablesFromDesktopFlowScript
        /// </summary>
        /// <param name="originalDesktopFlowScript"></param>
        /// <returns></returns>
        public static string RemoveVariablesFromDesktopFlowScript(string originalDesktopFlowScript) 
        {
            string script;

            // Check first if we have a desktop flow without input or output parameters
            var matchParameterCases = FindStringPositionWithRegex(originalDesktopFlowScript, @"(\@SENSITIVE: \[] \b)(?!.*\1)", true);
            if (matchParameterCases.Length > 0) // No input or output params defined in script
            {
                // Remove variable definition from script as these might contain sensitive information in them
                script = originalDesktopFlowScript.Substring(matchParameterCases.Index + matchParameterCases.Length, originalDesktopFlowScript.Length - (matchParameterCases.Index + matchParameterCases.Length));
            }
            else  // Input or output params defined in script
            { 
                // The regex function is used to find the end of the Sensitive, Input and Output Variables so we can remove them in the next step
                matchParameterCases = FindStringPositionWithRegex(originalDesktopFlowScript, @"(\ }  \b)(?!.*\1)|(\] \b)(?!.*\1)", true);
                // Remove variable definition from script as these might contain sensitive information 
                script = originalDesktopFlowScript.Substring(matchParameterCases.Index + matchParameterCases.Length, originalDesktopFlowScript.Length - (matchParameterCases.Index + matchParameterCases.Length));
            }

            return script;
        }

        /// <summary>
        /// FindStringPositionWithRegex
        /// </summary>
        /// <param name="inputString"></param>
        /// <param name="pattern"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        public static Match FindStringPositionWithRegex(string inputString, string pattern, bool ignoreCase)
        {
            try
            {
                return Regex.Match(inputString, pattern,
                                   ignoreCase ? RegexOptions.IgnoreCase | RegexOptions.Compiled : RegexOptions.Compiled,
                                   TimeSpan.FromSeconds(1));
            }
            catch (Exception ex)
            {
                throw new Exception("FindStringPositionWithRegex threw an exception: ", ex);
            }
        }

        /// <summary>
        /// FindStringOccurrencesWithRegEx
        /// </summary>
        /// <param name="inputString"></param>
        /// <param name="pattern"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        public static int FindStringOccurrencesWithRegEx(string inputString, string pattern, bool ignoreCase)
        {
            try
            {
                return Regex.Matches(inputString, pattern,
                                     ignoreCase ? RegexOptions.IgnoreCase | RegexOptions.Compiled : RegexOptions.Compiled,
                                     TimeSpan.FromSeconds(1)).Count;
            }
            catch (Exception ex)
            {
                throw new Exception("FindStringOccurrencesWithRegEx threw an exception: ", ex);
            }
        }

        /// <summary>
        /// CreateXml
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="cookie"></param>
        /// <param name="page"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            var reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }
    }

    /// <summary>
    /// DesktopFlowDefinition
    /// </summary>
    public class DesktopFlowDefinition
    {
        public string category { get; set; }
        public string modifiedon { get; set; }
        public string uiflowtype { get; set; }
        public string businessprocesstype { get; set; }
        public string mode { get; set; }
        public string name { get; set; }
        public string workflowidunique { get; set; }
        public string type { get; set; }
        public string workflowid { get; set; }
        public string createdon { get; set; }
        public string clientdata { get; set; }
        public string ownerid { get; set; }
        public string organizationid { get; set; }
        public string organizationidname { get; set; }
        public string fullname { get; set; }
        public string domainname { get; set; }
        public string firstname { get; set; }
        public string middlename { get; set; }
        public string internalemailaddress { get; set; }
        public string lastname { get; set; }
        public int succeededruns { get; set; }
        public int failedruns { get; set; }
        public int totalruns { get; set; }
        public string lastrundate { get; set; }

    }

    /// <summary>
    /// DesktopFlowSession
    /// </summary>
    public class DesktopFlowSession
    {
        public string errordetails { get; set; }
        public string statecode { get; set; }
        public string regardingobjectid { get; set; }
        public string regardingobjectidname { get; set; }
        public string context { get; set; }
        public string machinegroupid { get; set; }
        public string errorcode { get; set; }
        public string machineid { get; set; }
        public string machinegroupidname { get; set; }
        public string machineidname { get; set; }
        public string statuscode { get; set; }
        public string createdon { get; set; }
        public string flowsessionid { get; set; }
        public string completedon { get; set; }
        public string errormessage { get; set; }
        public string gateway { get; set; }
        public string startedon { get; set; }
    }

    /// <summary>
    /// DesktopFlowAction
    /// </summary>
    public class DesktopFlowAction
    {
        public string actionname { get; set; }
        public string desktopflowactionid { get; set; }
        public bool? dlpsupport { get; set; }
        public string moduledisplayname { get; set; }
        public string modulename { get; set; }
        public string modulesource { get; set; }
        public string selectorid { get; set; }
    }


    /// <summary>
    /// Action
    /// </summary>
    public class Action
    {
        public string FriendlyName { get; set; }
        public string Id { get; set; }
        public List<string> SelectorIds { get; set; }
        public bool DLPSupport { get; set; }
    }


    /// <summary>
    /// Module
    /// </summary>
    public class Module
    {
        public List<Action> Actions { get; set; }
        public string FriendlyName { get; set; }
        public string Id { get; set; }
        public string ModuleSource { get; set; }
    }


    /// <summary>
    /// DesktopFlowModule
    /// </summary>
    public class DesktopFlowModule
    {
        public List<Module> Modules { get; set; }
    }

    /// <summary>
    /// DesktopFlowActionUsage
    /// </summary>
    public class DesktopFlowActionUsage
    {
        public string desktopflowid { get; set; }
        public string desktopflowactionid { get; set; }
        public string selectorid { get; set; }
        public string actionname { get; set; }
        public string modulesource { get; set; }
        public string modulename { get; set; }
        public string desktopflowname { get; set; }
        public int occurrencecount { get; set; }
    }

    /// <summary>
    /// Definition
    /// </summary>
    public class Definition
    {
        public string package { get; set; }
        public string name { get; set; }
    }

    /// <summary>
    /// Properties
    /// </summary>
    public class Properties
    {        
        public Definition definition { get; set; }
    }

    /// <summary>
    /// WorkflowClientData
    /// </summary>
    public class WorkflowClientData
    {
        public Properties properties { get; set; }
    }
}
