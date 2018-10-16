using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UploadFile
{
   public class Statistics
    {
        public string UserName = "lokesh.gorle@acuvate.com";    // UserName of user
        public string SharePointSiteUrl = "https://acuvatehyd.sharepoint.com/teams/ExampleGratia"; //SharePoint Site URL
        public string ExcelDocLibName= "LokeshPractice";  // Document Library where Excel Sheet is Present
        public string ExcelFileName = "SharePointUploadList.xlsx"; // Name of Exel Document
        public string LocalDestinationFolder = @"D:\SPAssessment"; // Local Folder Path Url where downloaded File has to be stored 
        public string FilesUploadToDocLib = "LokeshPractice"; // Destination Document Library where the files must be Uploaded


    }
}
