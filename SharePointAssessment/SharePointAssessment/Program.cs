using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Configuration;
using System.Net.Mail;
using Microsoft.SharePoint.Client.Utilities;
//using ExcelServiceTest.XLService;

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
            }
        }
                //File file = ctx.Web.GetFileByServerRelativeUrl("/Shared%20Documents/testdata.xlsx");
                //ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                //ctx.Load(file);
                //ctx.ExecuteQuery();
                //using (var pck = new OfficeOpenXml.ExcelPackage())
                //{
                //    //using (var stream = File.OpenRead(""))
                //    //{
                //    //    pck.Load(stream);
                //    //}
                //    using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                //    {
                //        if (data != null)
                //        {
                //            data.Value.CopyTo(mStream);
                //            pck.Load(mStream);
                //            var ws = pck.Workbook.Worksheets.First();
                //            DataTable tbl = new DataTable();
                //            bool hasHeader = true; // adjust it accordingly( i've mentioned that this is a simple approach)
                //            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                //            {
                //                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                //            }
                //            var startRow = hasHeader ? 2 : 1;
                //            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                //            {
                //                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                //                var row = tbl.NewRow();
                //                foreach (var cell in wsRow)
                //                {
                //                    if (null != cell.Hyperlink)
                //                        row[cell.Start.Column - 1] = cell.Hyperlink;
                //                    else
                //                        row[cell.Start.Column - 1] = cell.Text;
                //                }
                //                tbl.Rows.Add(row);
                //            }
                //            Console.WriteLine('1');
                //    Web web = ctx.Web;
                //    List list = ctx.Web.Lists.GetByTitle("Documents");
                //    var file = list.RootFolder.Files.GetByUrl("SharePointAssessment.xlsx");
                //    ctx.Load(file);
                //    var listItem = list.GetItemById(1);
                //    ctx.Load(list);
                //    ctx.Load(listItem, i => i.File);
                //    ctx.ExecuteQuery();
                //    var fileRef = listItem.File.ServerRelativeUrl;
                //    var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileRef);
                //    var fileName = Path.Combine("C:\\Users\\epmadmin\\Desktop\\YourFolderName", (string)listItem.File.Name);
                //    using (var fileStream = System.IO.File.Create(fileName))
                //    {
                //        fileInfo.Stream.CopyTo(fileStream);
                //    }
                //    Console.WriteLine("found");
                //    Console.ReadKey();
                //}
                //private static void ReadFile(string FileURL,string SheetName,string Range)
                //{
                //    try
                //    {
                //        ExcelServices ObjXl = new ExcelServices();
                //    }
                //    Web web = ctx.Web;
                //    List list = ctx.Web.Lists.GetByTitle("Documents");
                //    var file = list.RootFolder.Files.GetByUrl("SharePointAssessment.xlsx");
                //    ctx.Load(file);
                //    var listItem = list.GetItemById(1);
                //    ctx.Load(list);
                //    ctx.Load(listItem, i => i.File);
                //    ctx.ExecuteQuery();
                //    var fileRef = listItem.File.ServerRelativeUrl;
                //    var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileRef);
                //    var fileName = Path.Combine("C:\\Users\\epmadmin\\Desktop\\YourFolderName", (string)listItem.File.Name);
                //    using (var fileStream = System.IO.File.Create(fileName))
                //    {
                //        fileInfo.Stream.CopyTo(fileStream);
                //    }
                //    Console.WriteLine("found");
                //    Console.ReadKey();
                //}
                //private static void ReadFile(string FileURL,string SheetName,string Range)
                //{
                //    try
                //    {
                //        ExcelServices ObjXl = new ExcelServices();
                //    }
        private static void ReadFileName(ClientContext clientContext)
        {
            string fileName = string.Empty;
            bool isError = true;
            const string fldTitle = "Title";
            const string lstDocName = "Documents";
            const string strFolderServerRelativeUrl = "/teams/SharePointDemo1/Shared%20Document";
            string strErrorMsg = string.Empty;
            try
            {
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
                camlQuery.FolderServerRelativeUrl = strFolderServerRelativeUrl;
                ListItemCollection listItems = list.GetItems(camlQuery);
                clientContext.Load(listItems, items => items.Include(i => i[fldTitle]));
                clientContext.ExecuteQuery();
                for (int i = 0; i < listItems.Count; i++)
                {
                    ListItem itemOfInterest = listItems[i];
                    if (itemOfInterest[fldTitle] != null)
                    {
                        fileName = itemOfInterest[fldTitle].ToString();
                        if (i == 0)
                        {
                            ReadExcelData(clientContext, itemOfInterest[fldTitle].ToString());
                        }
                    }
                }
                isError = false;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                    Console.ReadKey();
                }
            }
        }
        private static void ReadExcelData(ClientContext clientContext, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            const string lstDocName = "Documents";
            try
            {
                DataTable dataTable = new DataTable("EmployeeExcelDataTable");
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }
                    }
                }
                UpdateSPList(clientContext, dataTable, fileName);
                isError = false;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                    Console.ReadKey();
                }
            }
        }
        private static void UpdateSPList(ClientContext clientContext, DataTable dataTable, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            Int32 count = 0;
            const string lstName = "EmployeesData";
            const string lstColTitle = "Title";
            const string lstColAddress = "Address";
            try
            {
                string fileExtension = ".xlsx";
                string fileNameWithOutExtension = fileName.Substring(0, fileName.Length - fileExtension.Length);
                if (fileNameWithOutExtension.Trim() == lstName)
                {
                    List oList = clientContext.Web.Lists.GetByTitle(fileNameWithOutExtension);
                    foreach (DataRow row in dataTable.Rows)
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem oListItem = oList.AddItem(itemCreateInfo);
                        oListItem[lstColTitle] = row[0];
                        oListItem[lstColAddress] = row[1];
                        oListItem.Update();
                        clientContext.ExecuteQuery();
                        count++;
                    }
                }
                else
                {
                    count = 0;
                }
                if (count == 0)
                {
                    Console.Write("Error: List: '" + fileNameWithOutExtension + "' is not found in SharePoint.");
                }
                isError = false;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                    Console.ReadKey();
                }
            }
        }
        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                    Console.ReadKey();
                }
            }
            return value;
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