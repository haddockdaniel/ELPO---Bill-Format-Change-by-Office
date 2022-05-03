using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }
            DataSet myRSPC2 = new DataSet();
            comboBoxOffice.ClearItems();
            myRSPC2.Clear();
            string SQLPC2 = "select OfcOfficeCode  + '    ' + right(OfcDesc, 30) as PC from OfficeCode order by OfcOfficeCode";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no Offices. Correct and run the tool again");
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBoxOffice.Items.Add(dr["PC"].ToString());
                comboBoxOffice.SelectedIndex = 0;
            }

            comboBoxBill.ClearItems();
            comboBoxPreBill.ClearItems();
            myRSPC2.Clear();
            SQLPC2 = "select BLCode  + '    ' + right(BLDescription, 30) as PC from BillLayout order by BLCode";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no Bill Formats. Correct and run the tool again");
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                {
                    comboBoxBill.Items.Add(dr["PC"].ToString());
                    comboBoxPreBill.Items.Add(dr["PC"].ToString());
                }
                comboBoxBill.SelectedIndex = 0;
                comboBoxPreBill.SelectedIndex = 0;
            }


            //this.comboBoxBillLayout.GetItemText(this.comboBoxBillLayout.SelectedItem).Split(' ')[0]

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            string sql = "";

            DialogResult dr = MessageBox.Show("All matters for Office: " + this.comboBoxOffice.GetItemText(this.comboBoxOffice.SelectedItem).Split(' ')[0] + " will be modified to have" + "\r\n" +
                "Bill Format: " + this.comboBoxBill.GetItemText(this.comboBoxBill.SelectedItem).Split(' ')[0] + " and PreBill Format: " + this.comboBoxPreBill.GetItemText(this.comboBoxPreBill.SelectedItem).Split(' ')[0] +"." + "\r\n" +
                "This change cannot be undone. Are you sure?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                sql = "Update bt set bt.BillToBillFormat = '" + this.comboBoxBill.GetItemText(this.comboBoxBill.SelectedItem).Split(' ')[0] +
                    "', bt.BillToEditFormat = '" + this.comboBoxPreBill.GetItemText(this.comboBoxPreBill.SelectedItem).Split(' ')[0] + "' " +
                    "from billto bt inner join matter mt on mt.matbillto = bt.billtosysnbr " +
                    "where mt.MatOfficeCode = '" + this.comboBoxOffice.GetItemText(this.comboBoxOffice.SelectedItem).Split(' ')[0] + "'";
                _jurisUtility.ExecuteNonQuery(0, sql);

                MessageBox.Show("The proces completed without error.", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                System.Environment.Exit(0);
            }



        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>


        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }




    }
}
