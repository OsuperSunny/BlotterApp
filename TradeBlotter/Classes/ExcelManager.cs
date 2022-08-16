using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;


using Excel;



namespace TradeBlotter.Classes
{
    public class ExcelManager
    {
        public ExcelManager()
        {
        }

        string Connectionstring;
        bool Header = false;
        bool TreatIntermixedAsText = false;

        public bool header
        {
            get { return Header; }
            set { Header = value; }
        }

        public bool TreatIntermixedasText
        {
            get { return TreatIntermixedAsText; }
            set { TreatIntermixedAsText = value; }
        }

        public ExcelManager(string constr)
        {
            Connectionstring = constr;
        }

        /// <summary>
        /// returns a .NET memory datatable equivalent of excel data
        /// </summary>
        /// <param name="FilePath">a datatable</param>
        /// <returns>a datatable</returns>
        public System.Data.DataTable ReadFromExcel(string FilePath)
        {

            OleDbConnection objExcelConn;
            OleDbCommand objExcelCmdSelect;
            try
            {
                string ExcelConnectionString = GetExcelConnectionString(FilePath);
                objExcelConn = new OleDbConnection(ExcelConnectionString);

                objExcelConn.Open();
                var dt = objExcelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                // DataTable ExcelList = new DataTable();

                objExcelCmdSelect = new OleDbCommand("SELECT * FROM [" + dt.Rows[0]["Table_Name"].ToString() + "]", objExcelConn);

                var objExcelAdapter = new OleDbDataAdapter(objExcelCmdSelect);

                var objExcelDataset = new DataSet();

                objExcelAdapter.Fill(objExcelDataset);

                var ExcelList = objExcelDataset.Tables[0];

                objExcelCmdSelect.Dispose();
                objExcelConn.Close();

                return ExcelList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        /// <summary>
        /// returns a .NET memory datatable equivalent of excel data
        /// </summary>
        /// <param name="FilePath">a datatable</param>
        /// <returns>a datatable</returns>
        public System.Data.DataTable ReadDivWarrantFromExcel(string FilePath)
        {

            OleDbConnection objExcelConn;
            OleDbCommand objExcelCmdSelect;
            try
            {
                string ExcelConnectionString = GetExcelConnectionString(FilePath);
                objExcelConn = new OleDbConnection(ExcelConnectionString);

                objExcelConn.Open();
                var dt = objExcelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string qry = string.Empty;
                //qry = " SELECT * FROM [" + dt.Rows[0]["Table_Name"].ToString().Trim() + "]";

                qry = " SELECT * FROM [" + dt.Rows[0]["Table_Name"].ToString().Trim() + "] where Cheque_No is not null and Warrant_No is not null ";

                objExcelCmdSelect = new OleDbCommand(qry, objExcelConn);

                var objExcelAdapter = new OleDbDataAdapter(objExcelCmdSelect);

                var objExcelDataset = new DataSet();

                objExcelAdapter.Fill(objExcelDataset);

                var excelList = objExcelDataset.Tables[0];

                objExcelCmdSelect.Dispose();
                objExcelConn.Close();

                return excelList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public System.Data.DataTable ReadFromExcel(System.IO.Stream InputStream, bool openxml, int whichTable=0)
        {

            IExcelDataReader reader = null;

            try
            {

                if (!openxml)
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(InputStream);
                }
                else
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(InputStream);
                }

                reader.IsFirstRowAsColumnNames = header;
                DataSet result = reader.AsDataSet();

                return result.Tables[whichTable];

            }

            catch (Exception ex)
            {
                if (reader != null)
                {
                    if (!string.IsNullOrEmpty(reader.ExceptionMessage))
                    {
                        throw new Exception(ex.Message);
                    }
                }
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// write an excel file to a specified permanent or temporary sql-server table 
        /// </summary>
        /// <param name="FilePath">Excel file path</param>
        /// <param name="tabname">Existent or Non-Existent Sql-Server table name</param>
        /// <param name="cn">Sql-Server Connection string</param>
        /// <param name="tabletype">true represents a permanent table, while false represents a temporary table</param>
        


        /// <summary>
        /// write an excel file to a specified permanent or temporary sql-server table 
        /// and also returns a .NET memory datatable equivalent of the excel data
        /// </summary>
        /// <param name="FilePath">Excel file path</param>
        /// <param name="tabname">Existent or Non-Existent Sql-Server table name</param>
        /// <param name="cn">Sql-Server Connection string</param>
        /// <param name="tabletype">true represents a permanent table, while false represents a temporary table</param>
        /// <returns>a datatable</returns>
        public System.Data.DataTable FromExcel2SqlServer(string FilePath, string tabname, ref SqlConnection cn, bool tabletype)
        {
            try
            {
                string stmt = string.Empty;
                SqlCommand cm = new SqlCommand();
                string ExcelConnectionString = GetExcelConnectionString(FilePath);
                var objExcelConn = new OleDbConnection(ExcelConnectionString);

                objExcelConn.Open();
                var dt = objExcelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);


                var objCmdSelect = new OleDbCommand("SELECT * FROM [" + dt.Rows[0]["Table_Name"].ToString() + "]", objExcelConn);

                var objAdapter1 = new OleDbDataAdapter(objCmdSelect);

                var objDataset1 = new DataSet();

                objAdapter1.Fill(objDataset1);

                var ExcelList = objDataset1.Tables[0];

                if (tabletype == true)  //tabletype is a permanent physical sql-server table
                {
                    stmt = "IF OBJECT_ID('" + tabname + "') IS NULL  ";
                }
                else   //tabletype is false means it is a temporary sql-server table
                {
                    stmt = "IF OBJECT_ID('TempDB.." + tabname + "','U') IS NULL  ";
                }

                stmt += "Create table " + tabname + " ( ";

                for (int k = 0; k < ExcelList.Columns.Count; k++)
                {
                    string fname = ExcelList.Columns[k].ColumnName + " " + GetSqlDataType(ExcelList.Columns[k].DataType) + ", ";
                    stmt += fname;
                }

                stmt = stmt.Trim().Substring(0, stmt.Length - 2);
                stmt += ")";

                //create table here if it is non-existent
                cm.Connection = cn;
                cm.CommandType = CommandType.Text;
                cm.CommandText = stmt;
                cm.ExecuteNonQuery();

                //do a bulk copy to sql-server
                SqlBulkCopy bk = new SqlBulkCopy(cn);
                bk.DestinationTableName = tabname;
                bk.WriteToServer(ExcelList);
                objExcelConn.Close();

                return ExcelList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        public void Write2Excel(string FilePath, ref SqlConnection cn)
        {
            SqlCommand cm = new SqlCommand();

            OleDbConnection objExcelConn;
            OleDbCommand objExcelCmdSelect;
            try
            {
                //DataTable list = new DataTable();
                cm.Connection = cn;
                cm.CommandType = CommandType.Text;
                cm.CommandText = "SELECT transid, descr, debit, valdate FROM FromExcel";
                var ad = new SqlDataAdapter(cm);
                var ds = new DataSet();
                ad.Fill(ds);
                var list = ds.Tables[0];



                string ExcelConnectionString = GetExcelConnectionString(FilePath);
                objExcelConn = new OleDbConnection(ExcelConnectionString);

                objExcelConn.Open();
                //DataTable dt = objExcelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                objExcelCmdSelect = new OleDbCommand();
                objExcelCmdSelect.Connection = objExcelConn;
                objExcelCmdSelect.CommandType = CommandType.Text;

                string sql = string.Empty;

                for (int k = 0; k < list.Rows.Count; k++)
                {
                    //sql = "insert into [" + dt.Rows[0]["Table_Name"].ToString() + "] (transid, descr, debit, valdate)values ('";
                    sql = "insert Sheet1$ (transid, descr, debit, valdate) values ('";
                    sql += list.Rows[k]["transid"].ToString() + "','" + list.Rows[k]["descr"].ToString() + "','";
                    sql += list.Rows[k]["debit"].ToString() + "','" + list.Rows[k]["valdate"].ToString() + "')";

                    objExcelCmdSelect.CommandText = sql;

                    objExcelCmdSelect.ExecuteNonQuery();

                }

                objExcelCmdSelect.Dispose();


            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
                string msg = ex.Message + "-" + ex.StackTrace;  // +"-" + ex.InnerException.Source + "-" + ex.InnerException.Message + "-" + ex.InnerException.StackTrace;
            }
        }




        /// <summary>
        /// Utility method for method CreateTable.
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private string GetSqlDataType(Type type)
        {

            string s = type.ToString();
            switch (s)
            {

                case "System.DateTime":
                    s = "datetime";
                    break;

                case "System.Double":
                case "System.Decimal":
                    s = "money";
                    break;

                case "System.Int32":
                    s = "int";
                    break;

                case "System.Boolean":
                case "System.Byte":
                    s = "bit";
                    break;

                case "System.Char":
                case "System.String":
                    s = "varchar(1000)";
                    break;

                default:
                    s = "varchar(1000)";
                    break;
            }

            return s;

        }

        private string GetExcelConnectionString(string filePath)
        {
            // Note: the Types array exactly matches the entries in openFileDialog1.Filter
            string[] Types = {
                                 "Excel 12.0 Xml", // For Excel 2007 XML (*.xlsx)
			                 	"Excel 12.0", // For Excel 2007 Binary (*.xlsb)
			                 	"Excel 12.0 Macro", // For Excel 2007 Macro-enabled (*.xlsm)
			                 	"Excel 8.0", // For Excel 97/2000/2003 (*.xls)
			                 	"Excel 5.0" // For Excel 5.0/95 (*.xls)
                             };

            string Type = "";
            string extension = System.IO.Path.GetExtension(filePath).ToLower();
            switch (extension)
            {
                case ".xlsx":
                    Type = Types[0];
                    break;
                case ".xlsb":
                    Type = Types[1];
                    break;
                case ".xlsm":
                    Type = Types[2];
                    break;
                case ".xls":
                    Type = Types[3];
                    break;
            }

            // True if columns containing different data types are treated as text
            //  (note that columns containing only integer types are still treated as integer, etc)


            // Build the actual connection string
            var builder = new OleDbConnectionStringBuilder();
            builder.DataSource = filePath;
            if (Type == "Excel 5.0" || Type == "Excel 8.0")
                builder.Provider = "Microsoft.Jet.OLEDB.4.0";
            else
                builder.Provider = "Microsoft.ACE.OLEDB.12.0";
            builder["Extended Properties"] = Type +
                                             ";HDR=" + (Header ? "Yes" : "No") +
                                             ";IMEX=" + (TreatIntermixedAsText ? "1" : "0");


            return builder.ConnectionString.Replace(":", ":\\");
        }





    }
}
