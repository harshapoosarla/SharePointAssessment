using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Data;
using System.IO;

namespace SharePointAssessment
{
    class Program
    {
        static void Main(string[] args)
        {
            string UserName = "harsha.poosarla@acuvate.com";
            Console.WriteLine("Enter your password.");
            SecureString Password = GetPassword();
            using (var ctx = new ClientContext("https://acuvatehyd.sharepoint.com/teams/SharePointDemo1"))
            {
                ctx.Credentials = new SharePointOnlineCredentials(UserName, Password);

                ExcelPackage(ctx);
            }
        }
        private static void ExcelPackage(ClientContext ctx)
        {
            Microsoft.SharePoint.Client.File file = ctx.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/r/teams/SharePointDemo1/_layouts/15/Doc.aspx?sourcedoc=%7Ba5f7bba6-4627-4bfd-82f5-eccb8e7efd8c%7D&action=default&uid=%7BA5F7BBA6-4627-4BFD-82F5-ECCB8E7EFD8C%7D&ListItemId=4&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod");
            ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
            ctx.Load(file);
            ctx.ExecuteQuery();
            using (var pack = new OfficeOpenXml.ExcelPackage())
            {
                using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(memoryStream);
                        pack.Load(memoryStream);
                        var worksheet = pack.Workbook.Worksheets.First();
                        DataTable dataTable = new DataTable();
                        bool hasHeader = true;

                        foreach (var firstCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                        {
                            {
                                var print = dataTable.Columns.Add(hasHeader ? firstCell.Text : string.Format("Column {0}", firstCell.Start.Column));
                                Console.WriteLine(print);
                            }
                            var startRow = hasHeader ? 2 : 1;
                            for (var rowNumber = startRow; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                            {
                                var WorkSheetRow = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                                int j = 1;

                                string filetoupload = WorkSheetRow[rowNumber, j].Text;

                                int split = filetoupload.LastIndexOf('.');
                                string lhs = split < 0 ? filetoupload : filetoupload.Substring(0, split);
                                string rhs = split < 0 ? "" : filetoupload.Substring(split + 1);
                                System.IO.FileInfo sizeoffile = new System.IO.FileInfo(filetoupload);
                                long size = sizeoffile.Length;
                                string filename = sizeoffile.Name;
                                string deptoffile = WorkSheetRow[rowNumber, 6].Text;
                                Console.WriteLine(deptoffile);
                                if ( size <= 1.5e+7)
                                {
                                    List documentlibrary = ctx.Web.Lists.GetByTitle("UploadedDocuments");
                                    var filecreationinfo = new FileCreationInformation();
                                    filecreationinfo.Content = System.IO.File.ReadAllBytes(filetoupload);
                                    filecreationinfo.Overwrite = true;
                                    filecreationinfo.Url =filename;

                                    Microsoft.SharePoint.Client.File files = documentlibrary.RootFolder.Files.Add(filecreationinfo);
                                    ctx.ExecuteQuery();
                                    ListItem listItem = files.ListItemAllFields;
                                    listItem["Department"] = deptoffile;
                                    listItem["FileType"] = rhs;
                                    listItem.Update();

                                    ctx.Load(files);
                                    ctx.ExecuteQuery();
                                    Console.WriteLine("Successfully Uploaded");
                                }
                                else
                                {
                                    Console.WriteLine("Failed to Upload Files");
                                }
                            }
                        }
                    }
                }
                Console.ReadKey();
            }
        }
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