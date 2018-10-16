using System;
using Microsoft.SharePoint.Client;
using System.Security;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace SharePointAssessment
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter UserName:");
            string UserName= Console.ReadLine();
            //string UserName = "harsha.poosarla@acuvate.com";
            Console.WriteLine("Enter Password:");
            SecureString Password = GetPassword();
            using (var Context = new ClientContext("https://acuvatehyd.sharepoint.com/teams/SharePointDemo1"))
            {
                Context.Credentials = new SharePointOnlineCredentials(UserName, Password);
                //ExcelPackage(ctx);
                ReadFile(Context);
                ReadData(Context);
                UploadFile(Context);
            }
        }
        public static void ReadData(ClientContext context)
        {
            Excel.Application ExcelApp;
            Excel.Workbook ExcelWorkBook;
            Excel.Worksheet ExcelWorkSheet;
            Excel.Range ExcelRange;

            ExcelApp = new Excel.Application();
            ExcelWorkBook = ExcelApp.Workbooks.Open(@"D:\harsha853\SharePointAssessment.xlsx");
            ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            int MaximumRows = ExcelWorkSheet.UsedRange.Rows.Count;
            int MaximumColumns = ExcelWorkSheet.UsedRange.Columns.Count;

            ExcelRange = ExcelWorkSheet.UsedRange;
            string Reason;
            string UploadStatus;
            for (int Row = 2; Row < MaximumRows; Row++)
            {
                string FilePath = (ExcelRange.Cells[Row, 1] as Excel.Range).Value2;
                string Status = (ExcelRange.Cells[Row, 2] as Excel.Range).Value2;
                string CreatedBy = (ExcelRange.Cells[Row, 3] as Excel.Range).Value2;
                string Department = (ExcelRange.Cells[Row,6]as Excel.Range).Value2;
                AddFilesFromExcel(context, FilePath, CreatedBy, Status, Department, out Reason);
                UploadStatus = String.IsNullOrEmpty(Reason) ? "File Uploaded Successfully" : "Failed to Upload File";
                ExcelRange.Cells[Row, 4] = UploadStatus;
                ExcelRange.Cells[Row, 5] = Reason;
            }
            ExcelWorkBook.Save();
            ExcelWorkBook.Close();
            ExcelApp.Quit();
        }
        public static string AddFilesFromExcel(ClientContext context, string filepathstring, string createdby, string status,string department, out string reason)
        {
            List DepartmentList = context.Web.Lists.GetByTitle("Department");
            context.Load(DepartmentList);
            context.ExecuteQuery();

            CamlQuery CamlQuery = new CamlQuery();
            CamlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name ='_x0062_zf9'/><Value Type='Text'>" + department + "</Value></Eq></Where></Query><RowLimit></RowLimit></View>";
            ListItemCollection DepartmentListItems = DepartmentList.GetItems(CamlQuery);
            context.Load(DepartmentListItems);
            context.ExecuteQuery();

            string[] Array = filepathstring.Split('/');
            string FileNameForURL = Array[Array.Length - 1];
            FileInfo FileInfo = new FileInfo(filepathstring);
            long Filesize = FileInfo.Length;
            if (Filesize <= 1.5e+7)
            {
                try
                {
                    List List = context.Web.Lists.GetByTitle("DemoLib");

                    FileCreationInformation FileToUpload = new FileCreationInformation();

                    FileToUpload.Content = System.IO.File.ReadAllBytes(filepathstring);
                    FileToUpload.Overwrite = true;
                    FileToUpload.Url = "DemoLib/" + FileNameForURL;
                    Microsoft.SharePoint.Client.File uploadfile = List.RootFolder.Files.Add(FileToUpload);
                    Array = status.Split(',');
                    ListItem FileItem = uploadfile.ListItemAllFields;
                    FileItem["FileLeafRef"] = FileNameForURL;
                    FileItem["UploadStatus"] = Array;
                    FileItem["FileType"] = FileInfo.Extension;
                    FileItem["CreatedBy"] = createdby;

                    FileItem["Department"] = DepartmentListItems[0].Id;

                    FileItem.Update();
                    context.ExecuteQuery();
                    reason = "";
                    return reason;
                }
                catch (Exception e)
                {
                    ErrorLog.Errorlog(e);
                    return reason = e.Message;
                }
            }
            else
            {
                return reason = FileNameForURL + " file size exceeds the specified limit";
            }
        }
        public static void UploadFile(ClientContext context)
        {
            List DestList = context.Web.Lists.GetByTitle("DemoLib");
            FileCreationInformation FileCreationInformation = new FileCreationInformation();
            FileCreationInformation.Content = System.IO.File.ReadAllBytes(@"D:\harsha853\SharePointAssessment.xlsx");
            FileCreationInformation.Overwrite = true;
            FileCreationInformation.Url = "DemoLib/SharePointAssessment.xlsx";
            Microsoft.SharePoint.Client.File uploadfile = DestList.RootFolder.Files.Add(FileCreationInformation);
            uploadfile.Update();
            context.ExecuteQuery();
        }
        public static void ReadFile(ClientContext context)
        {
            var List = context.Web.Lists.GetByTitle("DemoLib");
            var ListItem = List.GetItemById(13);
            context.Load(List);
            context.Load(ListItem, i => i.File);
            context.ExecuteQuery();

            var FileRef = ListItem.File.ServerRelativeUrl;
            var FileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, FileRef);
            var FileName = System.IO.Path.Combine(@"D:\harsha853", (string)ListItem.File.Name);
            using (var FileStream = System.IO.File.Create(FileName))
            {
                FileInfo.Stream.CopyTo(FileStream);
            }
        }
        //    private static void ExcelPackage(ClientContext ctx)
        //{
        //    Microsoft.SharePoint.Client.File file = ctx.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/r/teams/SharePointDemo1/_layouts/15/Doc.aspx?sourcedoc=%7Ba5f7bba6-4627-4bfd-82f5-eccb8e7efd8c%7D&action=default&uid=%7BA5F7BBA6-4627-4BFD-82F5-ECCB8E7EFD8C%7D&ListItemId=4&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod");
        //    ClientResult<Stream> data = file.OpenBinaryStream();
        //    ctx.Load(file);
        //    ctx.ExecuteQuery();
        //    using (var pack = new OfficeOpenXml.ExcelPackage())
        //    {
        //        using (MemoryStream memoryStream = new MemoryStream())
        //        {
        //            if (data != null)
        //            {
        //                data.Value.CopyTo(memoryStream);
        //                pack.Load(memoryStream);
        //                var worksheet = pack.Workbook.Worksheets.First();
        //                DataTable dataTable = new DataTable();
        //                bool hasHeader = true;
        //                foreach (var firstCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
        //                {
        //                    {
        //                        var print = dataTable.Columns.Add(hasHeader ? firstCell.Text : string.Format("Column {0}", firstCell.Start.Column));
        //                        Console.WriteLine(print);
        //                    }
        //                    var startRow = hasHeader ? 2 : 1;
        //                    for (var rowNumber = startRow; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
        //                    {
        //                        var WorkSheetRow = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
        //                        int n = 1;
        //                        string filetoupload = WorkSheetRow[rowNumber, n].Text;
        //                        int split = filetoupload.LastIndexOf('.');
        //                        string rhs = split < 0 ? "" : filetoupload.Substring(split + 1);
        //                        FileInfo sizeoffile = new FileInfo(filetoupload);
        //                        long size = sizeoffile.Length;
        //                        string filename = sizeoffile.Name;
        //                        string deptoffile = WorkSheetRow[rowNumber, 6].Text;
        //                        string createdby = WorkSheetRow[rowNumber,3].Text;
        //                        string ups = WorkSheetRow[rowNumber, 2].Text;
        //                        string[] upssplit = ups.Split(',');
        //                        Console.WriteLine(deptoffile);
        //                        if ( size <= 1.5e+7)
        //                        {
        //                            List documentlibrary = ctx.Web.Lists.GetByTitle("DemoLib");
        //                            ctx.Load(documentlibrary);
        //                            var filecreationinfo = new FileCreationInformation();
        //                            filecreationinfo.Content = System.IO.File.ReadAllBytes(filetoupload);
        //                            filecreationinfo.Overwrite = true;
        //                            filecreationinfo.Url = Path.Combine("DemoLib/", Path.GetFileName(filename));
        //                            Microsoft.SharePoint.Client.File files = documentlibrary.RootFolder.Files.Add(filecreationinfo);
        //                            files.Update();
        //                            ctx.ExecuteQuery();
        //                            ListItem listItem = files.ListItemAllFields;
        //                            listItem["Department"] = deptoffile;
        //                            listItem["FileType"] = rhs;
        //                            listItem["CreatedBy"] =createdby;
        //                            listItem["UploadStatus"] =upssplit;
        //                            listItem.Update();
        //                            ctx.ExecuteQuery();
        //                            Console.WriteLine("Successfully Uploaded");
        //                        }
        //                        else
        //                        {
        //                            Console.WriteLine("Failed to Upload Files");
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    Console.ReadKey();
        //    }
        //}
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            SecureString SecurePassword = new SecureString();
            do
            {
               info = Console.ReadKey(true);
               if (info.Key != ConsoleKey.Enter)
               {
                  SecurePassword.AppendChar(info.KeyChar);
               }
            }
            while (info.Key != ConsoleKey.Enter);
            return SecurePassword;
        }
    }
}