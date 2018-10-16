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
            try
            {
                Console.WriteLine("Enter your password.");
                Credentials Credential = new Credentials();
                Statistics Stats = new Statistics();

                using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ExampleGratia"))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(Stats.UserName, Credential.password);

                    GetExcelFile(clientContext);  //Downloads Excel File from SharePoint//
                    ReadData(clientContext);    // Reads Data and Uploads Files into Document Library//
                    UploadExcelSheet(clientContext);  //Upload Updated Excel FIle//

                    Console.Read();
                }
            }
            catch(Exception ex)
            {
                ErrorLog.ErrorlogWrite(ex);
            }
        }

        /*------ to get the ID of the Excel File from the DOcument Library-----*/

        public static int GetItemId(ClientContext cxt, string ItemName)
        {
            Statistics Stats = new Statistics();
            try
            {
                List list = cxt.Web.Lists.GetByTitle(Stats.ExcelDocLibName);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='Text'>" + ItemName + "</Value></Eq></Where></Query></View>";
                ListItemCollection items = list.GetItems(camlQuery);
                cxt.Load(items);
                cxt.ExecuteQuery();
                int ItemID = items[0].Id;
                return ItemID;
            }
            catch(Exception ex)
            {
              
                throw;
                               
            }
           
        }

        
            
            /*-- To get ID of the department Name passed-*/
        public static int GetLookUpItemId(ClientContext cxt, string ItemName)
        {
            Statistics Stats = new Statistics();
            try
            {
                List list = cxt.Web.Lists.GetByTitle("Department");// everytime it is called to get ID from Department list only. So hard coded//
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + ItemName + "</Value></Eq></Where></Query></View>";
                ListItemCollection DeptItems = list.GetItems(camlQuery);
                cxt.Load(DeptItems);
                cxt.ExecuteQuery();

                int ItemID = DeptItems[0].Id;
                return ItemID;
            }
            catch(Exception ex)
            {
                throw;
            }
           
        }

        /*--- Get Excel File from Sharepoint into Local Machine---*/
        public static void GetExcelFile(ClientContext cxt)
        {
            Statistics Stats = new Statistics();
            try
            {
                var list = cxt.Web.Lists.GetByTitle(Stats.ExcelDocLibName);
                int DocID = GetItemId(cxt, Stats.ExcelFileName);
                var listItem = list.GetItemById(DocID); // Get ID by passing Department Name//
                cxt.Load(list);
                cxt.Load(listItem, i => i.File);
                cxt.ExecuteQuery();

                var FileRef = listItem.File.ServerRelativeUrl;
                var FileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(cxt, FileRef);
                var FileName = System.IO.Path.Combine(Stats.LocalDestinationFolder, Stats.ExcelFileName);// (string)listItem.File.Name);
                using (var fileStream = System.IO.File.Create(FileName))
                {
                    FileInfo.Stream.CopyTo(fileStream);
                }
            }
            catch(Exception ex)
            {
                throw;
            }

        }

        /*--- Read Data from the downloaded Excel File*/
        public static void ReadData(ClientContext cxt)
        {
            Statistics Stats = new Statistics();
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet exlWorkSheet;
                Excel.Range range;


                xlApp = new Excel.Application();
                string LocalFilePath = System.IO.Path.Combine(Stats.LocalDestinationFolder, Stats.ExcelFileName);
                xlWorkBook = xlApp.Workbooks.Open(LocalFilePath);
                exlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int maxrows = exlWorkSheet.UsedRange.Rows.Count;// To get the last row with data in excel sheet
                int maxcols = exlWorkSheet.UsedRange.Columns.Count;// To get last column with data

                range = exlWorkSheet.UsedRange;
                string Reason;
                string UploadStatus;
                for (int row = 2; row < 6; row++) /* starts with 2, because  first row is Column Header; 
                                                Max 4 files were Uploaded to test; "maxrows" can be used for all files*/
                {

                    string FilePath = (range.Cells[row, 1] as Excel.Range).Value2; // Column : FilePath//
                    string Status = (range.Cells[row, 2] as Excel.Range).Value2;// Column: Status(Created; To Be Approved; Verified)//
                    string CreatedBy = (range.Cells[row, 3] as Excel.Range).Value2;// Column : CreatedBy//
                    string DeptName = (range.Cells[row, 6] as Excel.Range).Value2;// Department Name//
                    UploadFilesToDocLib(cxt, FilePath, CreatedBy, Status, DeptName, out Reason); // Method called for uploading files//
                    UploadStatus = String.IsNullOrEmpty(Reason) ? "Uploaded" : "Failed"; // UploadStatus is determined from Reason 
                    range.Cells[row, 4] = UploadStatus; // Updating in  Excel Sheet//
                    range.Cells[row, 5] = Reason;
                }

                xlWorkBook.Save();
                xlWorkBook.Close();
                xlApp.Quit();
            }
            catch
            {
                throw;
            }


        }
        public static string UploadFilesToDocLib(ClientContext cxt, string FilepathString, string CreatedBy, string Status, string DepartmentName,out string Reason)
        {
            Statistics Stats = new Statistics();
            try
            {
                string[] Farr = FilepathString.Split('/');// Gets the FilePath into Array

                string FileNameForURL = Farr[Farr.Length - 1];// Gets  File Name from Path From the Array



                System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilepathString);

                long filesize = fileInfo.Length;

                if (filesize < 15000)   // file size should be less than 15KB
                {
                    try
                    {
                        int ID = GetLookUpItemId(cxt, DepartmentName);///Get ID for look up column by passing only the Dept Name.///
                        List ListToBeUpdated = cxt.Web.Lists.GetByTitle(Stats.FilesUploadToDocLib);


                        //------- File Creation-------//
                        FileCreationInformation FileToUpload = new FileCreationInformation();
                        FileToUpload.Content = System.IO.File.ReadAllBytes(FilepathString);
                        FileToUpload.Overwrite = true;
                        FileToUpload.Url = Stats.FilesUploadToDocLib + "/" + FileNameForURL;
                        File uploadfile = ListToBeUpdated.RootFolder.Files.Add(FileToUpload);



                        Farr = Status.Split(',');  // Status column in Excel is CSV . Converted to Array here
                        ListItem fileitem = uploadfile.ListItemAllFields;
                        cxt.Load(fileitem);
                        cxt.ExecuteQuery();
                        fileitem["Title"] = FileNameForURL;
                        fileitem["Multiselectcheck"] = Farr;// This Field Value is passed as Array to get the choices Checked
                        fileitem["FileType"] = fileInfo.Extension;
                        fileitem["CreatedBy"] = CreatedBy;         //columns updated
                        fileitem["Dept"] = ID;
                        fileitem.Update();
                        cxt.ExecuteQuery();

                        Reason = "";                        // if success there shall be no reason
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
            catch
            {
                throw;
            }
        }

        //------ to Upload the updated Excel Sheet------///
        public static void UploadExcelSheet(ClientContext cxt)
        {
            Statistics Stats = new Statistics();
            try
            {
                List DestList = cxt.Web.Lists.GetByTitle(Stats.ExcelDocLibName);
                FileCreationInformation Fci = new FileCreationInformation();
                Fci.Content = System.IO.File.ReadAllBytes(Stats.LocalDestinationFolder + "/" + Stats.ExcelFileName);
                Fci.Overwrite = true;
                Fci.Url = Stats.ExcelDocLibName + "/" + Stats.ExcelFileName;
                File uploadfile = DestList.RootFolder.Files.Add(Fci);
                uploadfile.Update();
                cxt.ExecuteQuery();
            }
            catch
            {
                throw;
            }


        }


        /*------- to upload excel file into sharepoint before everything starts---
         *  So there are hard coded values---------///*/
        public static void ADDFile(ClientContext cxt)
        {
            try
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
            catch(Exception ex)
            {
                ErrorLog.ErrorlogWrite(ex);
            }
        }



        //public static int GetLookUpItemId(string ItemName)
        //{
        //    DataTable Dept;
        //    Dept.Select()
        //    // Console.WriteLine("item id of " + title + " is  " + itemid);
        //}


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
