using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;
using TradeBlotter.Classes;

namespace TradeBlotter
{
    public partial class _Default : Page
    {
        private string folderPath = @"~/" + "UploadedFiles" + "/";
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            UploadFile(sender, e);
        }

        public void UploadFile(object sender, EventArgs e)
        {
           
            DataTable dtNEFTOutwardp;
            dtNEFTOutwardp = new DataTable();
            try
            {

                var Excel = new ExcelManager
                {
                    header = false,
                    TreatIntermixedasText = true
                };

                if (string.IsNullOrEmpty(FileUpload2.PostedFile.FileName))
                {
                    DisplayMessage("Please supply a file to be uploaded");
                    return;
                }

                var filename = FileUpload2.FileName.ToString();


                    //var filename = string.Empty;
                    try
                    {
                        //SAVE FILE TO PATH
                        if (FileUpload2.HasFile)
                        {

                            CreateMissingPath(folderPath);


                            filename = Path.GetFileName(FileUpload2.PostedFile.FileName);//"~/" + filename

                            var path = folderPath + filename;
                          

                            FileUpload2.SaveAs(Server.MapPath(path));
                      

                        }

                    }
                    catch (Exception ex)
                    {
                    throw (ex);
                    }



                int whichTable = 1;//1 will give me the second table while 0 first

                var dtExcelFile = Excel.ReadFromExcel(FileUpload2.PostedFile.InputStream, true, whichTable);
                if (dtExcelFile.Rows.Count > 0)
                {
                    try
                    {

                        var resultData = dtExcelFile.AsEnumerable().Where(o => o.Field<string>("Column7") != null && o.Field<string>("Column7") != "Cusip").Select(p => p.Field<string>("Column7")).ToList();

                        DisplayMessage(new JavaScriptSerializer().Serialize(resultData));
                    }
                    catch(Exception ex)
                    {  }

                   

                }
                else
                {

                    DisplayMessage("Unable to upload. Please,use the right Excel Template");
                    return;
                }

                FileUpload2.ID = null;
                 
            }
            catch (Exception ex)
            {      }
          
         }

        private void CreateMissingPath(string Path)
        {

            var pathf = Path;//"~/" + Path;
            bool FolderExists = Directory.Exists(System.Web.HttpContext.Current.Server.MapPath(pathf));
            if (!FolderExists)
            {
                Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath(pathf));
            }


        }
        public void DisplayMessage(string msg)
        {
            msg = msg.Replace("'", "").Replace("\\", "\\\\").Replace("\n", "");
            string sScript = "alert('" + msg + "');";
            ScriptManager.RegisterStartupScript(this, GetType(), "aa", sScript, true);
        }

    }
}