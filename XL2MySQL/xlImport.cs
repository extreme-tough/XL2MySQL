using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;

namespace XL2MySQL
{
    class xlImport
    {
        public string xlFilePath;
        public string server;
        public string uid,pwd,database;
        public bool clean;
        MySqlConnection oCon;
        public System.Windows.Forms.TextBox txtStatus;

        Microsoft.Office.Interop.Excel.Application oXL;
        _Workbook oWB;
        _Worksheet oSheet;
        Range oRng;

        public xlImport()
        {
            
        }

        public void import()
        {


            string conn = "server=" + server + ";uid=\"" + uid + "\";pwd=\"" + pwd + "\";database=" + database + ";";
            oCon = new MySqlConnection(conn);

            txtStatus.AppendText("Opening database connection\n");
            txtStatus.Refresh();

            MySqlCommand cmd ;
            oCon.Open();


            if (clean)
            {
                txtStatus.AppendText("Deleting old records \n");
                txtStatus.Refresh();

                cmd = new MySqlCommand();
                cmd.Connection = oCon;
                cmd.CommandText = "DELETE FROM org";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM doms";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM contact";
                cmd.ExecuteNonQuery();
            }
            txtStatus.AppendText("Opening excel file \n");
            txtStatus.Refresh();

            

            oXL = new Microsoft.Office.Interop.Excel.Application();
            oWB = oXL.Workbooks.Open(xlFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            oSheet = (_Worksheet)oWB.ActiveSheet;
            

            int lastCell = oSheet.get_Range("A1:A65988").SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            int i = 2;
            while (i <= lastCell)
            {
                txtStatus.AppendText("Importing row" + (i-1).ToString() + "\n");
                txtStatus.Refresh();


                string LICNo = "", DomName = "", DomPurpose = "", OrgName = "", OrgAddress = "", OrgPhone = "", OrgFax = "", OrgEmail = "", BillName = "", BillAddress = "", BillPhone = "", BillFax = "", BillEmail = "", AdminName = "", AdminPhone = "", AdminAddress = "";
                string AdminFax = "", AdminEmail = "", TechName = "", TechAddress = "", TechPhone = "", TechFax = "", TechEmail = "", DateApplied = "", TimeApplied = "", DateActive = "", DateToRenew = "", DNS1 = "", DNS2 = "";
                string DNS3 = "", DNS4 = "";
                string status = "", owed = "";

                try
                {
                    LICNo = oSheet.get_Range("$A$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { break;  }
                try
                {
                    DomName = oSheet.get_Range("$B$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    DomPurpose = oSheet.get_Range("$C$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    OrgName = oSheet.get_Range("$D$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    OrgAddress = oSheet.get_Range("$E$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    OrgPhone = oSheet.get_Range("$F$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    OrgFax = oSheet.get_Range("$G$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    OrgEmail = oSheet.get_Range("$H$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    BillName = oSheet.get_Range("$I$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    BillAddress = oSheet.get_Range("$J$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    BillPhone = oSheet.get_Range("$K$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    BillFax = oSheet.get_Range("$L$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    BillEmail = oSheet.get_Range("$M$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    AdminName = oSheet.get_Range("$N$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    AdminPhone = oSheet.get_Range("$O$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    AdminAddress = oSheet.get_Range("$P$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    AdminFax = oSheet.get_Range("$Q$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    AdminEmail= oSheet.get_Range("$R$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    TechName = oSheet.get_Range("$S$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    TechAddress = oSheet.get_Range("$T$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    TechPhone= oSheet.get_Range("$U$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    TechFax = oSheet.get_Range("$V$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    TechEmail = oSheet.get_Range("$W$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    object value = oSheet.get_Range("$X$" + i.ToString(), Type.Missing).Value2;

                    if (value != null)
                    {
                        if (IsDate(value.ToString()))
                            DateApplied = value.ToString();
                        else
                            DateApplied = DateTime.FromOADate((double)value).ToString("yyyy-MM-dd");
                    }

                }
                catch { }
                try
                {
                    object value = oSheet.get_Range("$Y$" + i.ToString(), Type.Missing).Value2;
                    if (value != null)
                    {
                        if (value.ToString().Contains(":"))
                            TimeApplied = value.ToString();
                        else 
                            TimeApplied = DateTime.FromOADate((double)value).ToLongTimeString().ToString();
                    }
                }
                catch { }
                try
                {
                    object value = oSheet.get_Range("$Z$" + i.ToString(), Type.Missing).Value2;

                    if (value != null)
                    {
                        if (IsDate(value.ToString()))
                            DateActive = value.ToString();
                        else
                            DateActive = DateTime.FromOADate((double)value).ToString("yyyy-MM-dd");
                    }
                }
                catch { }
                try
                {
                    object value = oSheet.get_Range("$AA$" + i.ToString(), Type.Missing).Value2;

                    if (value != null)
                    {
                        if (IsDate(value.ToString()))
                            DateToRenew = value.ToString();
                        else
                            DateToRenew = DateTime.FromOADate((double)value).ToString("yyyy-MM-dd");
                    }
                }
                catch { }
                try
                {
                    DNS1 = oSheet.get_Range("$AB$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    DNS2 = oSheet.get_Range("$AC$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    DNS3 = oSheet.get_Range("$AD$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }
                try
                {
                    DNS4 = oSheet.get_Range("$AE$" + i.ToString(), Type.Missing).Value2.ToString();
                }
                catch { }


                cmd = new MySqlCommand();
                cmd.Connection = oCon;
                cmd.CommandText = "SELECT idorg FROM org WHERE orgname='" + OrgName.Replace("'","''") + "'";
                
                object orgid = cmd.ExecuteScalar();


                cmd = new MySqlCommand();
                cmd.Connection = oCon;

                if (orgid == null)
                {
                    cmd.CommandText = "INSERT INTO org (orgname,orgaddress,orgphone,orgfax,orgemail) VALUES ('" +
                        OrgName.Replace("'", "''") + "','" + OrgAddress.Replace("'", "''") + "','" + OrgPhone + "','" + OrgFax + "','" + OrgEmail + "') ";
                    cmd.ExecuteNonQuery();

                    cmd = new MySqlCommand();
                    cmd.Connection = oCon;
                    cmd.CommandText = "SELECT LAST_INSERT_ID()";
                    orgid = cmd.ExecuteScalar();
                }

                try
                {
                    cmd = new MySqlCommand();
                    cmd.Connection = oCon;
                    cmd.CommandText = "INSERT INTO doms(licno,idorg,domname,dompurpose,dateapplied,timeapplied,dateactive,ns1,ns2,ns3,ns4,status,owed,daterenew) VALUES ('" +
                        LICNo + "'," + orgid + ",'" + DomName + "','" + DomPurpose.Replace("'", "''") + "','" + DateApplied + "','" + TimeApplied + "','" + DateActive + "','" + DNS1 + "','" +
                        DNS2 + "','" + DNS3 + "','" + DNS4 + "','" + status + "','" + owed + "','" + DateToRenew + "')";
                    cmd.ExecuteNonQuery();
                }

                catch (Exception ex)
                {
                    i++;
                    continue;
                }

                cmd = new MySqlCommand();
                cmd.Connection = oCon;
                cmd.CommandText = "INSERT INTO contact(idorg,cntcttype,cntctname,cntctaddress,cntctphone,cntctfax,cntctemail) VALUES (" +
                    orgid + ",'BILL','" + BillName.Replace("'", "''") + "','" + BillAddress.Replace("'", "''") + "','" + BillPhone + "','" + BillFax + "','" + BillEmail + "')";
                cmd.ExecuteNonQuery();

                cmd = new MySqlCommand();
                cmd.Connection = oCon;
                cmd.CommandText = "INSERT INTO contact(idorg,cntcttype,cntctname,cntctaddress,cntctphone,cntctfax,cntctemail) VALUES (" +
                    orgid + ",'ADMIN','" + AdminName.Replace("'", "''") + "','" + AdminAddress.Replace("'", "''") + "','" + AdminPhone + "','" + AdminFax + "','" + AdminEmail + "')";
                cmd.ExecuteNonQuery();

                cmd = new MySqlCommand();
                cmd.Connection = oCon;
                cmd.CommandText = "INSERT INTO contact(idorg,cntcttype,cntctname,cntctaddress,cntctphone,cntctfax,cntctemail) VALUES (" +
                    orgid + ",'TECH','" + TechName.Replace("'", "''") + "','" + TechAddress.Replace("'", "''") + "','" + TechPhone + "','" + TechFax + "','" + TechEmail + "')";
                cmd.ExecuteNonQuery();

                i++;
            }

            oCon.Close();

            oXL.Quit();
            
        }

        public bool IsDate(string strDate)
        {
            bool blnIsDate;
            blnIsDate = false;
            try
            {
                DateTime myDateTime = DateTime.Parse(strDate);
                blnIsDate = true;
            }
            catch { }
            return (blnIsDate);
        }

    }

}
