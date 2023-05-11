using Syncfusion.XlsIO;
using System.Data;

namespace SyncfusionExcelTest.Infra
{
    public class FileService : IFileService
    {
        public async Task GenerateExcel()
        {

            DataTable contacts = new DataTable();
            contacts.Columns.Add("HrCode");
            contacts.Columns.Add("FirstName");
            contacts.Columns.Add("Status");
            //var dv = dataSet11.Contacts.DefaultView;
            contacts.Rows.Add("-1", "Vimlesh", "Active");
            contacts.Rows.Add("-1", "Danny", "Active");
            contacts.Rows.Add("000256", "==> ADDITIONAL SUPPORT OF SOFTWARE DEVELOPMENT (art. 2.4 Collaboration Agreement)", "InActive");
            contacts.Rows.Add("162359", "Peter", "Active");
            contacts.Rows.Add("000236", "Mac", "Processing");
            contacts.Rows.Add("000010", "Joel", "Active");
            contacts.Rows.Add("023562", "Sammy", "Active");

            var filePath = "C:\\POC\\TestApplication\\DataTablePOC\\DataTablePOC\\InputTemplate.xlsx";
            var saveFilePath = "C:\\POC\\TestApplication\\DataTablePOC\\DataTablePOC\\OutputFile.xlsx";
            using (StreamReader filestream = new StreamReader(filePath))
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2013;
                    application.EnableIncrementalFormula = true;

                    IWorkbook workbook = application.Workbooks.Open(filestream.BaseStream);

                    IWorksheet worksheet = workbook.Worksheets[0];
                    //worksheet.SetText(2, 1, "%Contacts.HrCode;insert:copystyles");
                    worksheet.SetText(2, 2, "%Contacts.FirstName;insert:copystyles");
                    worksheet.SetText(2, 3, "%Contacts.Status;merge");
                    workbook.Version = ExcelVersion.Excel2016;
                    ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();
                    //marker.ApplyMarkers(UnknownVariableAction.Skip);
                    marker.AddVariable("Contacts", contacts, VariableTypeAction.None);
                    //marker.AddVariable("FirstName", dataSet11.Contacts, VariableTypeAction.DetectNumberFormat);
                    //marker.AddVariable("Status", dataSet11.Contacts, VariableTypeAction.DetectNumberFormat);
                    marker.ApplyMarkers(UnknownVariableAction.Skip);
                    using (var savefilestream = new FileStream(saveFilePath, FileMode.OpenOrCreate))
                        workbook.SaveAs(savefilestream);

                }
            }
        }
    }
}
