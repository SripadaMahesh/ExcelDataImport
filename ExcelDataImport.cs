using Microsoft.Office.Interop.Excel;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;

namespace D365.CRM.Apps.ExcelDataImport
{
    public class ExcelDataImport
    {
        static void Main(string[] args)
        {
            readExcel();
        }
        private static void readExcel()
        {
            string filePath = @"C:\Users\Mahesh Sripada\source\repos\D365.CRM.Apps.ExcelDataImport\Data\Employees.xlsx";
            Application application = new Application();
            Workbook wb = application.Workbooks.Open(filePath);
            Worksheet ws = wb.Worksheets[1];
            Range rows = ws.Rows;
            Range columns = ws.Columns;
            ServiceClient serviceClient = GetConnection();
            if (serviceClient.IsReady)
            {
                OrganizationRequestCollection requests = new OrganizationRequestCollection();
                Range firstRow = rows[1];
                foreach (Range row in rows)
                {
                    UpsertRequest upsertRequest = new UpsertRequest();
                    KeyAttributeCollection keys = new KeyAttributeCollection();
                    Entity employee = new Entity("cr19c_employee", keys);
                    RetrieveEntityRequest entityRequest = new RetrieveEntityRequest { EntityFilters = EntityFilters.Attributes, LogicalName = "cr19c_employee" };
                    RetrieveEntityResponse entityChangesResponse = (RetrieveEntityResponse)serviceClient.Execute(entityRequest);
                    EntityMetadata entityMetadata = entityChangesResponse.EntityMetadata;
                    foreach (var attribute in entityMetadata.Attributes)
                    {
                        string attributeLabel = attribute.DisplayName.UserLocalizedLabel.Label;
                        Range rangeCells = firstRow.Cells;
                        Range cells = row.Cells[1];
                        Console.WriteLine($"Range of the cells  {cells.Text}");
                    }
                    upsertRequest.Target = employee;
                    requests.Add(upsertRequest);
                }

                ExecuteMultipleRequest batch = new ExecuteMultipleRequest();
                batch.Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = true,
                    ReturnResponses = true,
                };
                batch.Requests = requests;
                var response = serviceClient.Execute(batch);
                Console.WriteLine($"Records are created {response.Results.Count}");
                serviceClient.Dispose();
            }
            else
            {
                Console.WriteLine("Could not connect to dataverse");
            }
        }
        private static ServiceClient GetConnection()
        {
            ServiceClient serviceClient = null;
            try
            {
                serviceClient = new ServiceClient(ConfigurationManager.ConnectionStrings["Xrm"].ConnectionString);
            }
            catch (Exception ex)
            {
                Console.WriteLine(" An exception occured " + ex.Message);
            }
            return serviceClient;
        }
    }
}
