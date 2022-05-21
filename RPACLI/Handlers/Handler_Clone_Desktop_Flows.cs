using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel;

namespace RPACLI.Handlers
{

    /// <summary>
    /// Handler_Clone_Desktop_Flows
    /// </summary>
    internal static class Handler_Clone_Desktop_Flows
    {
        /// <summary>
        /// CloneDesktopFlows
        /// </summary>
        /// <param name="Username"></param>
        /// <param name="SourceEnvironmentInstanceUrl"></param>
        /// <param name="DesktopFlowId"></param>
        /// <param name="Count"></param>
        /// <param name="BatchSize"></param>
        /// <exception cref="Exception"></exception>
        internal static void CloneDesktopFlows(string Username, string SourceEnvironmentInstanceUrl, string DesktopFlowId, int Count, int BatchSize)
        {
            ServiceClient service = null;
            try
            {
                service = CLIHelper.ConnectToDataverse(Username, SourceEnvironmentInstanceUrl);

                Console.WriteLine($"Running command: {AppDomain.CurrentDomain.FriendlyName} clone-desktop-flows --Username {Username} --SourceEnvironmentInstanceUrl {SourceEnvironmentInstanceUrl} --DesktopFlowId {DesktopFlowId} --Count {Count}");

                Console.WriteLine($"\n###############################################################################################################");
                Console.WriteLine($"Cloning {Count} Desktop flows based on Id: {DesktopFlowId}");
                Console.WriteLine($"###############################################################################################################\n");
                Console.WriteLine("{0} {1}",
                    "Desktop Flow Name".ToString().PadRight(65, ' '),
                    "Count".ToString().PadRight(20, ' ')
                );

                Entity desktopFlow = service.Retrieve("workflow", Guid.Parse(DesktopFlowId), new ColumnSet() { AllColumns = true });

                Console.WriteLine($"Start Time: {DateTime.Now.ToLongTimeString()}\n");
                if (desktopFlow != null)
                {
                    try
                    {
                        ExecuteMultipleRequest requestWithoutResults = new ExecuteMultipleRequest()
                        {
                            // Assign settings that define execution behavior: continue on error, return responses. 
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError = false,
                                ReturnResponses = false
                            },
                            // Create an empty organization request collection.
                            Requests = new OrganizationRequestCollection()
                        };

                        for (int i = 0; i < Count; i++)
                        {
                            if (i > 0 && i % BatchSize == 0)
                            {
                                Entity clonedDesktopFlow = new Entity("workflow");

                                foreach (KeyValuePair<string, object> attr in desktopFlow.Attributes)
                                {
                                    if (attr.Key == "statecode" || attr.Key == "statuscode" || attr.Key == "workflowid")
                                        continue;

                                    clonedDesktopFlow[attr.Key] = attr.Value;
                                }

                                clonedDesktopFlow["name"] = $"{clonedDesktopFlow["name"].ToString()} - Copy({i + 7417})";

                                requestWithoutResults.Requests.Add(new CreateRequest() { Target = clonedDesktopFlow });
                                Console.WriteLine("{0} {1}",
                                    clonedDesktopFlow["name"].ToString().PadRight(65, ' '),
                                    (7417 + i).ToString().PadRight(20, ' ')
                                );

                                ExecuteMultipleResponse responseWithoutResultsPaged = (ExecuteMultipleResponse)service.Execute(requestWithoutResults);

                                // If we have an error than the count would be > 0
                                if (responseWithoutResultsPaged.Responses.Count > 0)
                                {
                                    foreach (var response in responseWithoutResultsPaged.Responses)
                                    {
                                        if (response.Fault != null)
                                        {
                                            throw new Exception(response.Fault.ToString());
                                        }
                                    }
                                }

                                requestWithoutResults = new ExecuteMultipleRequest()
                                {
                                    // Assign settings that define execution behavior: continue on error, return responses. 
                                    Settings = new ExecuteMultipleSettings()
                                    {
                                        ContinueOnError = false,
                                        ReturnResponses = false
                                    },
                                    // Create an empty organization request collection.
                                    Requests = new OrganizationRequestCollection()
                                };
                            }
                            else
                            {
                                Entity clonedDesktopFlow = new Entity("workflow");

                                foreach (KeyValuePair<string, object> attr in desktopFlow.Attributes)
                                {
                                    if (attr.Key == "statecode" || attr.Key == "statuscode" || attr.Key == "workflowid")
                                        continue;

                                    clonedDesktopFlow[attr.Key] = attr.Value;
                                }

                                clonedDesktopFlow["name"] = $"{clonedDesktopFlow["name"].ToString()} - Copy({i + 7417})";

                                Console.WriteLine("{0} {1}",
                                    clonedDesktopFlow["name"].ToString().PadRight(65, ' '),
                                    (7417 + i).ToString().PadRight(20, ' ')
                                );

                                requestWithoutResults.Requests.Add(new CreateRequest() { Target = clonedDesktopFlow });
                            }
                        }

                        ExecuteMultipleResponse responseWithoutResults = (ExecuteMultipleResponse)service.Execute(requestWithoutResults);

                        // If we have an error than the count would be > 0
                        if (responseWithoutResults.Responses.Count > 0)
                        {
                            foreach (var response in responseWithoutResults.Responses)
                            {
                                if (response.Fault != null)
                                {
                                    throw new Exception(response.Fault.ToString());
                                }
                            }
                        }
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        Console.WriteLine($"The application terminated with an error.");
                        Console.WriteLine($"Timestamp: { ex.Detail.Timestamp}");
                        Console.WriteLine($"Code: {ex.Detail.ErrorCode}");
                        Console.WriteLine($"Message: {ex.Detail.Message}");
                    }
                    catch (System.TimeoutException ex)
                    {
                        Console.WriteLine($"The application terminated with a timeout exception.");
                        Console.WriteLine($"Message: {ex.Message}");
                        Console.WriteLine($"Stack Trace: {ex.StackTrace}");
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
                    throw new Exception($"Desktop flow with Id {DesktopFlowId} cannot be found.");

                Console.WriteLine($"\nEnd Time: {DateTime.Now.ToLongTimeString()}");

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
    }
}
