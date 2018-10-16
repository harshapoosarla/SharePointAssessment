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
                ReadExcelFile(Context);
                ReadData(Context);
                UploadExcelSheet(Context);
                
            }
        }
        public static void ReadData(ClientContext Context)
        {
            Excel.Application ExcelApp;
            Excel.Workbook ExcelWorkBook;
            Excel.Worksheet ExcelWorkSheet;
            Excel.Range ExcelRange;

            ExcelApp = new Excel.Application();
            ExcelWorkBook = ExcelApp.Workbooks.Open(@"D:\harsha853\SharePointAssessment.xlsx");
            ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelRange = ExcelWorkSheet.UsedRange;
            string Reason;
            string UploadStatus;
            for (int row = 2; row < 8; row++)
            {
                string FilePath = (ExcelRange.Cells[row, 1] as Excel.Range).Value2;
                string status = (ExcelRange.Cells[row, 2] as Excel.Range).Value2;
                string CreatedBy = (ExcelRange.Cells[row, 3] as Excel.Range).Value2;
                string Department = (ExcelRange.Cells[row,6]as Excel.Range).Value2;
                AddFilesFromExcel(Context, FilePath, CreatedBy, status, Department, out Reason);
                UploadStatus = String.IsNullOrEmpty(Reason) ? "File Uploaded Successfully" : "Failed to Upload File";
                ExcelRange.Cells[row, 4] = UploadStatus;
                ExcelRange.Cells[row, 5] = Reason;
            }
            ExcelWorkBook.Save();
            ExcelWorkBook.Close();
            ExcelApp.Quit();
        }
        public static string AddFilesFromExcel(ClientContext Context, string FilepathString, string CreatedBy, string Status,string Department, out string Reason)
        {
            List DeptList = Context.Web.Lists.GetByTitle("Department");
            Context.Load(DeptList);
            Context.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name ='_x0062_zf9'/><Value Type='Text'>" + Department + "</Value></Eq></Where></Query><RowLimit></RowLimit></View>";
            ListItemCollection DepartmentListItems = DeptList.GetItems(camlQuery);
            Context.Load(DepartmentListItems);
            Context.ExecuteQuery();

            string[] array = FilepathString.Split('/');
            string FileNameForURL = array[array.Length - 1];
            FileInfo fileInfo = new FileInfo(FilepathString);
            long filesize = fileInfo.Length;
            if (filesize <= 1.5e+7)
            {
                try
                {
                    List list = Context.Web.Lists.GetByTitle("DemoLib");

                    FileCreationInformation fileToUpload = new FileCreationInformation();

                    fileToUpload.Content = System.IO.File.ReadAllBytes(FilepathString);
                    fileToUpload.Overwrite = true;
                    fileToUpload.Url = "DemoLib/" + FileNameForURL;
                    Microsoft.SharePoint.Client.File uploadfile = list.RootFolder.Files.Add(fileToUpload);
                    array = Status.Split(',');
                    ListItem fileitem = uploadfile.ListItemAllFields;
                    fileitem["FileLeafRef"] = FileNameForURL;
                    fileitem["UploadStatus"] = array;
                    fileitem["FileType"] = fileInfo.Extension;
                    fileitem["CreatedBy"] = CreatedBy;

                    fileitem["Department"] = DepartmentListItems[0].Id;

                    fileitem.Update();
                    Context.ExecuteQuery();
                    Reason = "";
                    return Reason;
                }
                catch (Exception E)
                {
                    return Reason = E.Message;
                }
            }
            else
            {
                return Reason = FileNameForURL + " file size exceeds the specified limit";
            }
        }
        public static void UploadExcelSheet(ClientContext Context)
        {
            List DestList = Context.Web.Lists.GetByTitle("DemoLib");
            FileCreationInformation fileCreationInformation = new FileCreationInformation();
            fileCreationInformation.Content = System.IO.File.ReadAllBytes(@"D:\harsha853\SharePointAssessment.xlsx");
            fileCreationInformation.Overwrite = true;
            fileCreationInformation.Url = "DemoLib/SharePointAssessment.xlsx";
            Microsoft.SharePoint.Client.File uploadfile = DestList.RootFolder.Files.Add(fileCreationInformation);
            uploadfile.Update();
            Context.ExecuteQuery();
        }
        public static void ReadExcelFile(ClientContext Context)
        {
            var list = Context.Web.Lists.GetByTitle("DemoLib");
            var listItem = list.GetItemById(13);
            Context.Load(list);
            Context.Load(listItem, i => i.File);
            Context.ExecuteQuery();

            var fileRef = listItem.File.ServerRelativeUrl;
            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(Context, fileRef);
            var fileName = System.IO.Path.Combine(@"D:\harsha853", (string)listItem.File.Name);
            using (var fileStream = System.IO.File.Create(fileName))
            {
                fileInfo.Stream.CopyTo(fileStream);
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