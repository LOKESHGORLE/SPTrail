using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;


namespace UploadFile
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Enter your password.");
            Credentials Credential = new Credentials();
            Statistics Stats = new Statistics();

            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ExampleGratia"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(Stats.userName, Credential.password);


                //GetFile(clientContext);
                //AddFiles(clientContext);
                //ReadExcelData(clientContext, "SharePointUploadList.xlsx");
                //ADDFile(clientContext);

                GetExcelFile(clientContext);
                ReadData(clientContext);
                UploadExcelSheet(clientContext);

                Console.Read();
            }
        }

        public static int GetItemId(ClientContext cxt, string ItemName)
        {
            Statistics Stats = new Statistics();

            List list = cxt.Web.Lists.GetByTitle(Stats.ExcelDocLibName);
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='Text'>" + ItemName + "</Value></Eq></Where></Query></View>";
            ListItemCollection items = list.GetItems(camlQuery);
            cxt.Load(items);
            cxt.ExecuteQuery();


            int ItemID = items[0].Id;
            return ItemID;
           // Console.WriteLine("item id of " + title + " is  " + itemid);
        }


        //public static int GetLookUpItemId(string ItemName)
        //{
        //    DataTable Dept;
        //    Dept.Select()
        //    // Console.WriteLine("item id of " + title + " is  " + itemid);
        //}

        public static int GetLookUpItemId(ClientContext cxt, string ItemName)
        {
            Statistics Stats = new Statistics();

            List list = cxt.Web.Lists.GetByTitle("Department");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + ItemName + "</Value></Eq></Where></Query></View>";
            ListItemCollection DeptItems = list.GetItems(camlQuery);
            cxt.Load(DeptItems);
            cxt.ExecuteQuery();

            int ItemID = DeptItems[0].Id;
            return ItemID;
            // Console.WriteLine("item id of " + title + " is  " + itemid);
        }
        public static void GetExcelFile(ClientContext cxt)
        {
            Statistics Stats = new Statistics();

            var list = cxt.Web.Lists.GetByTitle(Stats.ExcelDocLibName);
            int DocID = GetItemId(cxt, Stats.ExcelFileName);
            var listItem = list.GetItemById(DocID);
            cxt.Load(list);
            cxt.Load(listItem, i => i.File);
            cxt.ExecuteQuery();

            var fileRef = listItem.File.ServerRelativeUrl;
            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(cxt, fileRef);
            var fileName = System.IO.Path.Combine(Stats.LocalDestinationFolder, Stats.ExcelFileName);// (string)listItem.File.Name);
            using (var fileStream = System.IO.File.Create(fileName))
            {
                fileInfo.Stream.CopyTo(fileStream);
            }

        }

        public static void ReadData(ClientContext cxt)
        {
            Statistics Stats = new Statistics();

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            xlApp = new Excel.Application();
            var LocalFilePath = System.IO.Path.Combine(Stats.LocalDestinationFolder, Stats.ExcelFileName);
            xlWorkBook = xlApp.Workbooks.Open(LocalFilePath);//@"D:\SPAssessment\SharePointUploadList.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //int lastrow = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            range = xlWorkSheet.UsedRange;
            string Reason;
            string UploadStatus;
            for (int row = 2; row < 6; row++)
            {

                string FilePath = (range.Cells[row, 1] as Excel.Range).Value2; // Column : FilePath
                string status = (range.Cells[row, 2] as Excel.Range).Value2;
                string CreatedBy = (range.Cells[row, 3] as Excel.Range).Value2;
                string DeptName = (range.Cells[row, 6] as Excel.Range).Value2;
                AddFilesFromExcel(cxt, FilePath, CreatedBy, status, DeptName, out Reason);
                UploadStatus = String.IsNullOrEmpty(Reason) ? "Uploaded" : "Failed";
                range.Cells[row, 4] = UploadStatus;
                range.Cells[row, 5] = Reason;
            }

            xlWorkBook.Save();
            xlWorkBook.Close();
            xlApp.Quit();


        }
        public static string AddFilesFromExcel(ClientContext cxt, string FilepathString, string CreatedBy, string Status, string DepartmentName,out string Reason)
        {
            Statistics Stats = new Statistics();

            string[] farr = FilepathString.Split('/');
            //string FilepathString = row[datacolumn].ToString();
            string FileNameForURL = farr[farr.Length - 1];



            System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilepathString);

            long filesize = fileInfo.Length;

            if (filesize < 15000)
            {
                try
                {
                    int ID = GetLookUpItemId(cxt, DepartmentName);
                    List l = cxt.Web.Lists.GetByTitle(Stats.FilesUploadToDocLib); 



                     FileCreationInformation fileToUpload = new FileCreationInformation();
                    fileToUpload.Content = System.IO.File.ReadAllBytes(FilepathString);
                    fileToUpload.Overwrite = true;
                    fileToUpload.Url = Stats.FilesUploadToDocLib + "/" + FileNameForURL;



                    File uploadfile = l.RootFolder.Files.Add(fileToUpload);



                    farr = Status.Split(',');
                    ListItem fileitem = uploadfile.ListItemAllFields;
                    cxt.Load(fileitem);
                    cxt.ExecuteQuery();
                    fileitem["Title"] = FileNameForURL;
                    fileitem["Multiselectcheck"] = farr;
                    fileitem["FileType"] = fileInfo.Extension;
                    fileitem["CreatedBy"] = CreatedBy;
                    //FieldLookupValue lookup =fileitem["Dept"] as FieldLookupValue;
                    ///check for look up column by passing only the id.///
                   
                    //lookup.LookupId = ID;
                    fileitem["Dept"] = ID;
                    fileitem.Update();
                    // cxt.Load(item);

                    //uploadfile.Update();
                    cxt.ExecuteQuery();
                    //Console.ReadLine();
                    Reason = "";
                    return Reason;
                }
                catch (Exception ex)
                {
                    return Reason = ex.Message;
                }
            }
            else
            {
                return Reason = FileNameForURL + " file size exceed";
            }


        }
        public static void UploadExcelSheet(ClientContext cxt)
        {
            Statistics Stats = new Statistics();

            List DestList = cxt.Web.Lists.GetByTitle(Stats.ExcelDocLibName);
            FileCreationInformation Fci = new FileCreationInformation();
            Fci.Content = System.IO.File.ReadAllBytes(Stats.LocalDestinationFolder+"/"+Stats.ExcelFileName);
            Fci.Overwrite = true;
            Fci.Url =Stats.ExcelDocLibName+"/"+Stats.ExcelFileName ;//"LokeshPractice/SharePointUploadList.xlsx";

            File uploadfile = DestList.RootFolder.Files.Add(Fci);
            uploadfile.Update();
            cxt.ExecuteQuery();


        }

        public static void ADDFile(ClientContext cxt)
        {
            var pathstring = @"D:/SPAssessment/SharePointUploadList.xlsx";
            List l = cxt.Web.Lists.GetByTitle("LokeshPractice");
           
            FileCreationInformation fileToUpload = new FileCreationInformation();
            fileToUpload.Content = System.IO.File.ReadAllBytes(pathstring);
            fileToUpload.Url = "LokeshPractice/SharePointUploadList.xlsx";

            
            Microsoft.SharePoint.Client.File uploadfile = l.RootFolder.Files.Add(fileToUpload);


            ListItem item = uploadfile.ListItemAllFields;
            item["Title"] = "File generated using Code";
            
            item.Update();
           
            cxt.ExecuteQuery();
        }



        //public static void GetFile(ClientContext cxt)
        //{

        //    List UploadToList = cxt.Web.Lists.GetByTitle("LokeshPractice");
        //    CamlQuery camlQuery = new CamlQuery();
        //    camlQuery.ViewXml = @"<View><Query></Query></View>";
        //    //camlQuery.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='Name'/><Value Type='Text'>SharePointUploadList</Value></Eq></Where></Query></View>";
        //    ListItemCollection FilesinLib = UploadToList.GetItems(camlQuery);
        //    cxt.Load(FilesinLib);
        //    cxt.ExecuteQuery();


        //    foreach (ListItem file in FilesinLib)
        //    {
        //        Console.WriteLine(file.File.Author);
        //        File ExcelFile = file.File;
        //        cxt.Load(ExcelFile);
        //        cxt.ExecuteQuery();
        //        string FileUrl = ExcelFile.ServerRelativeUrl;

        //        //Console.WriteLine(file.FieldValues["Title"].ToString());
        //        Console.WriteLine(file.FieldValues["FileLeafRef"]);
        //    }



        //}

        //private static void ReadExcelData(ClientContext clientContext, string fileName)
        //{

        //    string strErrorMsg = string.Empty;
        //    const string lstDocName = "LokeshPractice";
        //    try
        //    {
        //        DataTable dataTable = new DataTable("ExcelDataTable");
        //        List list = clientContext.Web.Lists.GetByTitle(lstDocName);
        //        clientContext.Load(list.RootFolder);
        //        clientContext.ExecuteQuery();
        //        string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
        //        File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
        //        ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
        //        clientContext.Load(file);
        //        clientContext.ExecuteQuery();

        //        using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
        //        {
        //            if (data != null)
        //            {
        //                data.Value.CopyTo(mStream);
        //                using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
        //                {
        //                    //WorkbookPart workbookPart = document.WorkbookPart;
        //                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
        //                    string relationshipId = sheets.First().Id.Value;
        //                    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
        //                    //Worksheet workSheet = worksheetPart.Worksheet;
        //                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        //                    IEnumerable<Row> rows = sheetData.Descendants<Row>();
        //                    foreach (Cell cell in rows.ElementAt(0))
        //                    {
        //                        string str = GetCellValue(clientContext, document, cell);
        //                        dataTable.Columns.Add(str);
        //                    }
        //                    foreach (Row row in rows)
        //                    {
        //                        if (row != null)
        //                        {
        //                            DataRow dataRow = dataTable.NewRow();
        //                            for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
        //                            {
        //                                dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
        //                            }
        //                            dataTable.Rows.Add(dataRow);
        //                        }
        //                    }
        //                    dataTable.Rows.RemoveAt(0);
        //                }

        //                for (int datarow = 0; datarow < 3; datarow++)
        //                {
        //                    DataRow r = dataTable.Rows[datarow];
        //                    AddFiles(clientContext, r);

        //                }

        //            }
        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message + "0000");
        //    }
        //}

        //public static void AddFiles(ClientContext cxt, DataRow row)
        //{

        //    int datacolumn = 0;

        //    string[] farr = row[datacolumn].ToString().Split('/');
        //    string FilepathString = row[datacolumn].ToString();
        //    string FileNameForURL = farr[farr.Length - 1];



        //    System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilepathString);

        //    long filesize = fileInfo.Length;
        //    //var pathstring = @"D:/SPAssessment/FilesToUpload/SharePointUploadList.xlsx";
        //    if (filesize < 15000)
        //    {

        //        List l = cxt.Web.Lists.GetByTitle("LokeshPractice");



        //        FileCreationInformation fileToUpload = new FileCreationInformation();
        //        fileToUpload.Content = System.IO.File.ReadAllBytes(FilepathString);
        //        fileToUpload.Overwrite = true;
        //        fileToUpload.Url = "LokeshPractice/" + FileNameForURL;


        //        //fileToUpload.Content.GetLength(filesize);


        //        // fileToUpload.Url = "LokeshPractice/SharePointUploadList.xlsx";

        //        //folder.Folders.GetByUrl("LokeshPractice").Folders.GetByUrl("created Folder");

        //        //  var list = cxt.Web.Lists.GetByTitle("LokeshPractice");
        //        File uploadfile = l.RootFolder.Files.Add(fileToUpload);

        //        //File fil = folder.Files.Add(fileToUpload);

        //        farr = row["Status"].ToString().Split(',');
        //        ListItem fileitem = uploadfile.ListItemAllFields;
        //        fileitem["Title"] = "File generated using Code";
        //        fileitem["Multiselectcheck"] = farr;
        //        fileitem["File_x0020_Type"] = fileInfo.Extension;
        //        fileitem["CreatedBy"] = row["CreatedBy"];
        //        fileitem.Update();
        //        // cxt.Load(item);

        //        //uploadfile.Update();
        //        cxt.ExecuteQuery();
        //        //Console.ReadLine();
        //    }
        //    else
        //    {
        //        Console.WriteLine(FileNameForURL + " file size exceed");
        //    }

        //}

        //private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        //{
        //    bool isError = true;
        //    string strErrorMsg = string.Empty;
        //    string value = string.Empty;
        //    try
        //    {
        //        if (cell != null)
        //        {
        //            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
        //            if (cell.CellValue != null)
        //            {
        //                value = cell.CellValue.InnerXml;
        //                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //                {
        //                    if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
        //                    {
        //                        isError = false;
        //                        return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
        //                    }
        //                }
        //                else
        //                {
        //                    isError = false;
        //                    return value;
        //                }
        //            }
        //        }
        //        isError = false;
        //        return string.Empty;
        //    }
        //    catch (Exception e)
        //    {
        //        isError = true;
        //        strErrorMsg = e.Message;
        //    }
        //    finally
        //    {
        //        if (isError)
        //        {
        //            //Logging
        //        }
        //    }
        //    return value;
        //}



    }
}
