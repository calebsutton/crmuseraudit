// System namespaces
using System;
using System.IO;
using System.Text;

// XRM namespaces
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

// Commandline namespace
using CommandLine;


namespace CRMUserAudit
{
    class Program
    {
        // Defining command line options
        public class Options
        {
            private int _days;

            [Option("url", Required = true, HelpText = "URL of Dynamics 365 instance.")]
            public string Url { get; set; }

            [Option("username", Required = true, HelpText = "Username with Audit access.")]
            public string Username { get; set; }

            [Option("password", Required = true, HelpText = "Password for user with Audit access.")]
            public string Password { get; set; }

            [Option("path", Required = false, Default = ".\\", HelpText = "Path to export results.")]
            public string Path { get; set; }

            [Option("filename", Required = false, Default = "CRMAuditExport.csv", HelpText = "Filename to export results.")]
            public string Filename { get; set; }

            [Option("days", Required = false, Default = 30, HelpText = "Number of days to export data for.")]
            public int Days
            {
                get => _days;
                set
                {
                    // Make sure Days option is a negative value
                    if (value > 0)
                    {
                        value *= -1;  
                    }
                    _days = value;
                }
            }

            [Option("filteruser", Required = false, Default = null, HelpText = "Username to filter.  If not specified, will export all users except SYSTEM")]
            public string FilterUser { get; set; }
        }

        static void Main(string[] args)
        {
            // Parse command line options and launch application
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(options =>
                {
                    // Build path to save CSV
                    String filePath = Path.Combine(options.Path, options.Filename); 

                    // Create CSV
                    StringBuilder csv = new StringBuilder();
            
                    // Add header line to top of CSV
                    String csvHeader = "Action,User,Operation,Related Object,Date"; 
                    csv.AppendLine(csvHeader);

                    // Setup CRM connection
                    String connectionString =
                        $"Url={options.Url}; Username={options.Username}; Password={options.Password}; authtype=Office365";
                    CrmServiceClient conn = new CrmServiceClient(connectionString);
                    OrganizationServiceProxy orgServiceProxy = conn.OrganizationServiceProxy;

                    orgServiceProxy.EnableProxyTypes();

                    // Create a new query for Audit entity
                    var query = new QueryExpression(Audit.EntityLogicalName)
                    {
                        ColumnSet = new ColumnSet(true),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };

                    if (String.IsNullOrEmpty(options.FilterUser))
                    {
                        // Filter out SYSTEM user
                        query.Criteria.AddCondition("useridname", ConditionOperator.NotEqual, "SYSTEM");
                    } else
                    {
                        // Filter specified user
                        query.Criteria.AddCondition("useridname", ConditionOperator.Equal, options.FilterUser);
                    }
            

                    // Filter out web user logins
                    query.Criteria.AddCondition("objecttypecode", ConditionOperator.NotEqual, "sockeye_webuserloginattempt");

                    // Filter by date using Days option, time period between now and X days ago           
                    query.Criteria.AddCondition( "createdon", ConditionOperator.GreaterEqual, (DateTime.Now).AddDays(options.Days));

                    // execute query
                    var results = orgServiceProxy.RetrieveMultiple(query);

                    // Loop through results and append to CSV
                    foreach (Entity audit in results.Entities)
                    {
                        OptionSetValue action = (OptionSetValue)audit["action"];  // Get Action value
                        Audit_Action actionName = (Audit_Action)action.Value; // Get Action name from optionset using value
                        EntityReference user = (EntityReference)audit["userid"]; // Get User entity reference from userid
                        OptionSetValue operation = (OptionSetValue)audit["operation"]; // Get Operation value
                        Audit_Operation operationName = (Audit_Operation)operation.Value; // Get Operation name from optionset using value
                        DateTime createdon = DateTime.Parse(audit["createdon"].ToString()); // get createdon date and cast from string to datetime
                        EntityReference objectid = (EntityReference)audit["objectid"]; // Get operation target entity reference from objectid

                        // build line for csv
                        String newLine = $"{actionName},{user.Name},{operationName},{objectid},{(createdon.ToLocalTime())}";

                        // append line to csv
                        csv.AppendLine(newLine);
                    }

                    // write csv to disk
                    File.WriteAllText(filePath, csv.ToString());
                });
        }
    }
}
