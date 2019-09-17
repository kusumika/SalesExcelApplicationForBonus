using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;  //Task 202: Select all data for all year

namespace GenerateBonusReport
{
    public partial class GenerateBonus : Form
    {
        string sBonusType = string.Empty;

        public GenerateBonus()
        {
            InitializeComponent();
            try
            //Test
            {
                //string cellValue = "-100";
                //string n = Regex.Replace(cellValue, @"[^0-9-]", "");
                // cellValue = "-1  00";
                // n = Regex.Replace(cellValue, @"[^0-9-]", "");

                // cellValue = "38 344";
                // n = Regex.Replace(cellValue, @"[^0-9-]", "");

                // cellValue = "38344";
                // n = Regex.Replace(cellValue, @"[^0-9-]", "");


            }
            catch (Exception ex)
            {
 
            }
        }

        private void GenerateBonus_Load(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// Create an Excel file, and write it to a file for Leverantor data.
        /// </summary>
        
        //KD Changes
        private void CreateExcelforLeverantor()
        {
            try
            {
                //KD Change
                DataSet ds = GetSupplierForBonusSetAmount();
                string sSupplier = string.Empty;
                DataSet dsSupplier = new DataSet();
                string sColumnName = string.Empty;
                for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                {
                    if (sSupplier != ds.Tables[0].Rows[i]["SupplierName"].ToString())
                    {
                        sColumnName = ds.Tables[0].Rows[i]["SupplierName"].ToString();
                        dsSupplier.Tables.Add(sColumnName);

                        dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[0].ColumnName, ds.Tables[0].Columns[0].DataType));
                        //dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[1].ColumnName, ds.Tables[0].Columns[1].DataType));
                        dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[2].ColumnName, ds.Tables[0].Columns[2].DataType));
                        dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[3].ColumnName, ds.Tables[0].Columns[3].DataType));
                        dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[4].ColumnName, ds.Tables[0].Columns[4].DataType));
                        dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[5].ColumnName, ds.Tables[0].Columns[5].DataType));
                        dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[6].ColumnName, ds.Tables[0].Columns[6].DataType));
                    }

                    DataRow newRow = dsSupplier.Tables[sColumnName].NewRow();
                    newRow[0] = ds.Tables[0].Rows[i][0];
                    //newRow[1] = ds.Tables[0].Rows[i]["SupplierName"].ToString();
                    newRow[1] = ds.Tables[0].Rows[i][2];
                    newRow[2] = ds.Tables[0].Rows[i][3];
                    newRow[3] = ds.Tables[0].Rows[i][4];
                    newRow[4] = ds.Tables[0].Rows[i][5];
                    newRow[5] = ds.Tables[0].Rows[i][6];

                    dsSupplier.Tables[ds.Tables[0].Rows[i]["SupplierName"].ToString()].Rows.Add(newRow);

                    sSupplier = ds.Tables[0].Rows[i]["SupplierName"].ToString();
                }

                string documentPath = ConfigurationManager.AppSettings["ReportDir"];
                string filename = "StatisticsBonusgroupsBySupplier" + ".xlsx";
                CreateExcelFile2.CreateExcelDocument(dsSupplier, documentPath + filename);
                MessageBox.Show("Excel created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Create an Excel file, and write it to a file for Medlemsnamn data.
        /// </summary>

        //KD Changes
        private void CreateExcelforMedlemsnamn()
        {
            try
            {
                //KD Change
                DataSet ds = GetMemberForBonusSetAmount();

                //Task 202: Select all data for all year

                #region Select all data for all year For Task 202
                try
                {
                    bool isFileGenerated = false;
                    string filenameCSV = "StatisticsBonusgroupsByMember_All";
                    string fileSavedPath = WriteCSVFileSourcedFromDB(filenameCSV, ds, out isFileGenerated);

                    if (!string.IsNullOrEmpty(fileSavedPath))
                    {
                        MessageBox.Show("CSV Generated Successfully");
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Error for CSV: " + ex);
                }
                #endregion


                #region the existing running code hide for Task 202: Select all data for all year

                while (false)   //Task 202: Select all data for all year, extra while loop added
                {
                    string sSupplier = string.Empty;
                    DataSet dsSupplier = new DataSet();
                    string sColumnName = string.Empty;

                    //Task 202: Generate only one file with all data 
                    //int p = 0;
                    //string trace = string.Empty;
                    try
                    {
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //KD Change, Task, MemberName changed to MemberNo
                            //p = i;
                            if (!sSupplier.Equals(ds.Tables[0].Rows[i]["MemberName"].ToString(), StringComparison.InvariantCultureIgnoreCase))
                            {
                                //KD Change, Task, MemberName changed to MemberNo
                                sColumnName = ds.Tables[0].Rows[i]["MemberName"].ToString();

                                //trace = "sColumnName: " + sColumnName;

                                //Task 202: Select all data for all year
                                //if (sColumnName != null && sColumnName.Contains("J Gustavssons"))
                                //{ }


                                dsSupplier.Tables.Add(sColumnName);

                                //trace = "dsSupplier: " + sColumnName;

                                dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn("Leverantörsnummer", ds.Tables[0].Columns[0].DataType));
                                //dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn(ds.Tables[0].Columns[1].ColumnName, ds.Tables[0].Columns[1].DataType));
                                dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn("Leverantörsnamn", ds.Tables[0].Columns[2].DataType));
                                dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn("Startdatum", ds.Tables[0].Columns[3].DataType));
                                dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn("Slutdatum", ds.Tables[0].Columns[4].DataType));
                                dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn("Bonusgrupp namn", ds.Tables[0].Columns[5].DataType));
                                dsSupplier.Tables[sColumnName].Columns.Add(new DataColumn("Bonusgrupp belopp", ds.Tables[0].Columns[6].DataType));
                            }

                            DataRow newRow = dsSupplier.Tables[sColumnName].NewRow();
                            //Task 202: Generate only one file with all data 
                            //trace = "DataRow Add";

                            newRow[0] = ds.Tables[0].Rows[i][0];
                            //newRow[1] = ds.Tables[0].Rows[i]["MemberName"];
                            newRow[1] = ds.Tables[0].Rows[i][2];
                            newRow[2] = ds.Tables[0].Rows[i][3];
                            newRow[3] = ds.Tables[0].Rows[i][4];
                            newRow[4] = ds.Tables[0].Rows[i][5];
                            newRow[5] = ds.Tables[0].Rows[i][6];
                            //Task 202: Select all data for all year
                            //KD Change, Task, MemberName changed to MemberNo
                            dsSupplier.Tables[ds.Tables[0].Rows[i]["MemberName"].ToString()].Rows.Add(newRow);

                            //Task 202: Select all data for all year
                            //trace = "DataRow Add newRow";

                            //KD Change, Task, MemberName changed to MemberNo
                            sSupplier = ds.Tables[0].Rows[i]["MemberName"].ToString();

                            //trace = "sSupplier";
                        }
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message + p + trace);
                        MessageBox.Show(ex.Message);
                    }

                    string documentPath = ConfigurationManager.AppSettings["ReportDir"];

                    //Task 202: Select all data for all year, .xlsx has changed to .CSV file
                    string filename = "StatisticsBonusgroupsByMember" + ".xlsx";

                    bool status = CreateExcelFile2.CreateExcelDocument(dsSupplier, documentPath + filename);
                    MessageBox.Show("Excel created successfully.\n" + status);
                }
                #endregion Task 202: Select all data for all year
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        
        //KD Change
        public DataSet GetSupplierForBonusSetAmount()
        {
            return ConvertDataReaderToDataSet(GetSupplierForBonusSetAmountData());
        }
        
        //KD Changes
        public IDataReader GetSupplierForBonusSetAmountData()
        {
            string connString = ConfigurationManager.AppSettings["ConnectionString"];

            SqlConnection con = null;
            SqlCommand cmd = null;
            //Creating new connection object
            con = new SqlConnection(connString);

            //Opening new Connection
            con.Open();

            cmd = new SqlCommand("sp_SelectPRIM_SupplierForBonusSetAmount", con);
            cmd.CommandTimeout = 300;

            cmd.CommandType = CommandType.StoredProcedure;


            // Execute the command

            // Create SqlDataAdapter instance and fill DataSet

            SqlDataReader reader = cmd.ExecuteReader();
            return reader;
        }

        //KD Change
        public DataSet GetMemberForBonusSetAmount()
        {
            return ConvertDataReaderToDataSet(GetMemberForBonusSetAmountData());
        }

        //KD Changes
        public IDataReader GetMemberForBonusSetAmountData()
        {
            string connString = ConfigurationManager.AppSettings["ConnectionString"];

            SqlConnection con = null;
            SqlCommand cmd = null;
            //Creating new connection object
            con = new SqlConnection(connString);

            //Opening new Connection
            con.Open();

            cmd = new SqlCommand("sp_SelectPRIM_MemberForBonusSetAmount", con);
            cmd.CommandTimeout = 300;

            cmd.CommandType = CommandType.StoredProcedure;


            // Execute the command

            // Create SqlDataAdapter instance and fill DataSet

            SqlDataReader reader = cmd.ExecuteReader();
            return reader;
        }

        /// <summary>
        /// Create DataSet from IDataReader for supliers list.
        /// </summary>
        /// <param name="IDataReader">IDataReader containing the data to be convert to the DataSet.</param>
        /// <returns>DataSet containing supliers list of Leverantor</returns>
        private DataSet ConvertDataReaderToDataSet(IDataReader reader)
        {
            DataSet ds = new DataSet();
            DataTable dataTable = new DataTable();

            DataTable schemaTable = reader.GetSchemaTable();
            DataRow row;

            string columnName;
            DataColumn column;
            int count = schemaTable.Rows.Count;

            for (int i = 0; i < count; i++)
            {
                row = schemaTable.Rows[i];
                columnName = (string)row["ColumnName"];

                column = new DataColumn(columnName, (Type)row["DataType"]);
                dataTable.Columns.Add(column);
            }

            ds.Tables.Add(dataTable);

            object[] values = new object[count];

            try
            {
                dataTable.BeginLoadData();
                while (reader.Read())
                {
                    reader.GetValues(values);
                    dataTable.LoadDataRow(values, true);
                }
            }
            finally
            {
                dataTable.EndLoadData();
                reader.Close();
            }

            return ds;
        }
        
        private void btnLEV_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            //KD Changes
            CreateExcelforLeverantor();
            sBonusType = string.Empty;
            Cursor.Current = Cursors.Default;
        }

        private void btnMED_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            CreateExcelforMedlemsnamn();
            sBonusType = string.Empty;
            Cursor.Current = Cursors.Default;
        }


        //Task 202: Generate only one file with all data 
        public string WriteCSVFileSourcedFromDB(string fileName, DataSet ds, out bool isFileGenerated)
        {
            isFileGenerated = false;
            string csvFileName = string.Empty;

            string documentPath = @"D:\KD\Bolist\GenerateBonusNewReportKD\GenerateBonusReport\Report\";
            csvFileName = documentPath + fileName + ".csv";


            try
            {
                #region Poppulate Columns in 1st Row of CSV

                if (ds.Tables[0] != null)
                {
                    string tableColumnLine = string.Empty;

                    foreach (DataColumn col in ds.Tables[0].Columns)
                    {
                        if (tableColumnLine == string.Empty)
                        {
                            tableColumnLine = !string.IsNullOrEmpty(col.ColumnName) ? col.ColumnName.Trim() : string.Empty;
                        }
                        else
                        {
                            tableColumnLine += !string.IsNullOrEmpty(col.ColumnName) ? ";" + col.ColumnName.Trim() : ";" + string.Empty;
                        }
                    }

                    if (!string.IsNullOrEmpty(tableColumnLine))
                    {
                        try
                        {
                            if (System.IO.File.Exists(csvFileName))
                            {
                                System.IO.File.Delete(csvFileName);
                            }

                            StreamWriter logStream = new StreamWriter(csvFileName, false);
                            logStream.WriteLine(tableColumnLine);
                            logStream.Close();

                            isFileGenerated = true;
                        }

                        catch (Exception ex)
                        {
                            isFileGenerated = false;
                            MessageBox.Show(ex.Message + isFileGenerated); 
                        }
                    }
                }

                #endregion

                #region Poppulate / Add All Other Data Rows to CSV

                if (ds.Tables[0] != null && isFileGenerated)
                {
                    foreach (DataRow currRow in ds.Tables[0].Rows)
                    {
                        string tableDataLine = string.Empty;

                        foreach (DataColumn currCol in ds.Tables[0].Columns)
                        {
                            if (string.IsNullOrEmpty(tableDataLine))
                            {
                                tableDataLine = currRow[currCol] != System.DBNull.Value ? currRow[currCol].ToString() : string.Empty;
                            }
                            else
                            {
                                tableDataLine += currRow[currCol] != System.DBNull.Value ? ";" + currRow[currCol].ToString() : ";" + string.Empty;
                            }
                        }

                        if (!string.IsNullOrEmpty(tableDataLine))
                        {
                            if (isFileGenerated)
                            {
                                StreamWriter logStream = new StreamWriter(csvFileName, true);
                                logStream.WriteLine(tableDataLine);
                                logStream.Close();
                            }
                            else
                            {
                                StreamWriter logStream = new StreamWriter(csvFileName, false);
                                logStream.WriteLine(tableDataLine);
                                logStream.Close();

                                isFileGenerated = true;
                            }
                        }
                    }
                }


                #endregion
            }

            catch (Exception ex)
            {
                isFileGenerated = false;
                MessageBox.Show(ex.Message); 
            }

            return csvFileName;
        }

    }

}
