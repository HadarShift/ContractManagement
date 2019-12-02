using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.IO;
using System.Data.SqlClient;
using System.Collections;


namespace ContractManagement
{
    class contract
    {
        public int Contract_id { get; set; }
        public string Name_supplier { get; set; }
        public string Initator_name { get; set; }
        public string Subject { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public double Sum_Contract { get; set; }
        public string Currency { get; set; }
        public string Order_Num { get; set; }
        public string Comment { get; set; }
        public DateTime Ensurance_Date { get; set; }
        public string Status { get; set; }
        public int Contract_Num { get; set; }
        public int Days_Alert { get; set; }
        public int Days_Cancel { get; set; }
        public int Frequency_Alert { get; set; }
        public int Months_Renew { get; set; }
        public int Bonus_Preiod { get; set; }
        //public string Bonus_Preiod { get; set; }
        public DateTime Last_Date_Alert { get; set; }
        public DateTime Last_Update { get; set; }//תאריך עדכון אחרון של החוזה
        public string Stop_Alert { get; set; }         
        //עבור קבצים
        public string File_Name { get; set; }
        public string Email { get; set; }



        public contract(int Contract_id)
        {
            this.Contract_id = Contract_id;
        }

        public contract()
        {

        }

        public contract(int contract_Id, string name_supplier, string initator_name,
                        string subject, DateTime date_start, DateTime date_end, double sum_contract,
                        string currency, string order_num, string comment, DateTime ensurance_date,
                       string status, int contract_num, int days_alert, int days_cancel, int frequency_alert
                        , int months_renew, int bonus_period, DateTime last_date_alert, DateTime last_update,
                       string stop_alert, string file_name,string email)

        {
            Contract_id = contract_Id;
            Name_supplier = name_supplier;
            Initator_name = initator_name;
            Subject = subject;
            DateStart = date_start;
            DateEnd = date_end;
            Sum_Contract = sum_contract;
            Currency = currency;
            Order_Num = order_num;
            Comment = comment;
            Ensurance_Date = ensurance_date;
            Status = status;
            Contract_Num = contract_num;
            Days_Alert = days_alert;
            Days_Cancel = days_cancel;
            Frequency_Alert = frequency_alert;
            Months_Renew = months_renew;
            Bonus_Preiod = bonus_period;
            Last_Date_Alert = last_date_alert;
            Last_Update = last_update;
            Stop_Alert = stop_alert;
            File_Name = file_name;
            Email = email;
        }


        public DataTable SupplierList()
        {
            DBService D_AS400 = new DBService();//עבור רשימת ספקים
            DataTable Result400 = new DataTable();
            string sapak_string = $@"SELECT SPKNO as num,SPNAME as name                                 FROM GCTINVF18.NSPAK                                 GROUP BY SPKNO, SPNAME";            Result400 = D_AS400.executeSelectQueryNoParam(sapak_string);
            return Result400;
        }

        public DataTable OrderList()
        {
            DBService D_AS400 = new DBService();
            DataTable order_nums = new DataTable();//עבור רשימת הזמנות
            string orders = $@"SELECT DISTINCT  ORDP1 as OrderNum
                               FROM GCTINVF18.NPORD1
                               WHERE OPNDT1>'170000'";
            order_nums = D_AS400.executeSelectQueryNoParam(orders);
            return order_nums;
        }

        public void InsertContract()
        {
            DbServiceSQL sqlcombo = new DbServiceSQL();
            DataTable DT = new DataTable();
            string values = $@"INSERT INTO dbo.contract VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
            SqlParameter contract_id1 = new SqlParameter("a", Contract_id);
            SqlParameter name_supplier1 = new SqlParameter("b", Name_supplier);
            SqlParameter initator_name1 = new SqlParameter("c", Initator_name);
            SqlParameter subject1 = new SqlParameter("d", Subject);
            SqlParameter date_Start1 = new SqlParameter("e", DateStart.ToString("MM/dd/yyyy"));
            SqlParameter date_End1 = new SqlParameter("f", DateEnd.ToString("MM/dd/yyyy"));
            SqlParameter sum_contract1 = new SqlParameter("h", Sum_Contract);
            SqlParameter currency1 = new SqlParameter("i", Currency);
            SqlParameter order_num1 = new SqlParameter("j", Order_Num);
            SqlParameter comment1 = new SqlParameter("k", Comment);
            SqlParameter ensurance_Date1;
            if (Ensurance_Date==DateTime.MinValue)
            ensurance_Date1 = new SqlParameter("l", DBNull.Value);
            else ensurance_Date1 = new SqlParameter("l", Ensurance_Date.ToString("MM/dd/yyyy"));
            SqlParameter status1 = new SqlParameter("m", Status);
            SqlParameter contract_num1 = new SqlParameter("n", Contract_Num);
            SqlParameter days_alert1 = new SqlParameter("o", Days_Alert);
            SqlParameter days_cancel1 = new SqlParameter("p", Days_Cancel);
            SqlParameter frequency_alert1 = new SqlParameter("q", Frequency_Alert);
            SqlParameter months_renew1 = new SqlParameter("r", Months_Renew);
            SqlParameter bonus_period1 = new SqlParameter("s", Bonus_Preiod);
            SqlParameter last_date_alert1 = new SqlParameter("u", Last_Date_Alert.ToString("MM/dd/yyyy"));
            SqlParameter last_update1 = new SqlParameter("v", Last_Update.ToString("yyyy-MM-dd HH:mm:ss"));
            SqlParameter stop_alert1 = new SqlParameter("s", Stop_Alert);
            SqlParameter email1 = new SqlParameter("e", Email);
            int row_change = sqlcombo.ExecuteQuery(values, CommandType.Text, contract_id1, name_supplier1, initator_name1, subject1, date_Start1, date_End1, sum_contract1, currency1, order_num1, comment1, ensurance_Date1,
                                   status1, contract_num1, days_alert1, days_cancel1, frequency_alert1, months_renew1, bonus_period1, last_date_alert1, last_update1, stop_alert1,email1);//שליחת ערכים במידה והחוזה מסתיים ללא חידוש
            if (row_change != 0) MessageBox.Show("חוזה נוסף בהצלחה");
            Copy_file();//העלאת קובץ במידה ויש


        }



        public void UpdateContract()
        {
            DialogResult r = MessageBox.Show("?האם לעדכן את השינויים בחוזה", "אישור עדכון", MessageBoxButtons.YesNo);
            if (r == DialogResult.Yes)
            {
                try
                {
                    DbServiceSQL sqlcombo = new DbServiceSQL();
                    DataTable DT = new DataTable();
                    string for_update = $@"UPDATE contract
                                    SET Supplier_Code=(?),Initiator_Name=(?),Subject=(?),Date_Start=(?),Date_End=(?),Sum_Contract=(?),Currency=(?),
                                    Order_Num=(?),Comments=(?),Ensurance_Date=(?),Status=(?),Contract_Num=(?),Days_Alert=(?),Days_For_Cancel=(?),
                                    Frequency_Alert=(?),Months_Renew=(?),Bonus_Period=(?),Last_Update=(?),Stop_Alert=(?),Email_User=(?)
                                    WHERE Cnum= {Contract_id}";
                    SqlParameter name_supplier1 = new SqlParameter("b", Name_supplier);
                    SqlParameter initator_name1 = new SqlParameter("c", Initator_name);
                    SqlParameter subject1 = new SqlParameter("d", Subject);
                    SqlParameter date_Start1 = new SqlParameter("e", DateStart.ToString("MM/dd/yyyy"));
                    SqlParameter date_End1 = new SqlParameter("f", DateEnd.ToString("MM/dd/yyyy"));
                    SqlParameter sum_contract1 = new SqlParameter("h", Sum_Contract);
                    SqlParameter currency1 = new SqlParameter("i", Currency);
                    SqlParameter order_num1 = new SqlParameter("j", Order_Num);
                    SqlParameter comment1 = new SqlParameter("k", Comment);
                    SqlParameter ensurance_Date1;
                    if (Ensurance_Date == DateTime.MinValue)
                        ensurance_Date1 = new SqlParameter("l", DBNull.Value);
                    else ensurance_Date1 = new SqlParameter("l", Ensurance_Date.ToString("MM/dd/yyyy"));
                    SqlParameter status1 = new SqlParameter("m", Status);
                    SqlParameter contract_num1 = new SqlParameter("n", Contract_Num);
                    SqlParameter days_alert1 = new SqlParameter("o", Days_Alert);
                    SqlParameter days_cancel1 = new SqlParameter("p", Days_Cancel);
                    SqlParameter frequency_alert1 = new SqlParameter("q", Frequency_Alert);
                    SqlParameter months_renew1 = new SqlParameter("r", Months_Renew);
                    SqlParameter bonus_period1 = new SqlParameter("s", Bonus_Preiod);
                    SqlParameter last_update1 = new SqlParameter("w", Last_Update.ToString("yyyy-MM-dd HH:mm:ss"));
                    SqlParameter stop_alert1 = new SqlParameter("s", Stop_Alert);
                    SqlParameter email1 = new SqlParameter("e", Email);
                    int row_change = sqlcombo.ExecuteQuery(for_update, CommandType.Text, name_supplier1, initator_name1, subject1, date_Start1, date_End1, sum_contract1, currency1, order_num1, comment1, ensurance_Date1,
                                          status1, contract_num1, days_alert1, days_cancel1, frequency_alert1, months_renew1, bonus_period1, last_update1, stop_alert1,email1); 
                    Copy_file();
                    if (row_change != 0)  MessageBox.Show("חוזה עודכן בהצלחה");
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }



            }

        }


        //get data from existing contract include pdf file
        /// <summary>
        /// זאת ריצה מול ה DB
        /// </summary>
        /// <returns></returns>
        public bool GetData()
        {
            string qry = $@"SELECT *
                            FROM contract t 
                            WHERE Cnum = {Contract_id}";
            DbServiceSQL For_Update = new DbServiceSQL();
            DataTable One_Row = new DataTable();

            if (One_Row.Rows.Count >= 0)
            {
                One_Row = For_Update.executeSelectQueryNoParam(qry);
                Name_supplier = One_Row.Rows[0]["supplier_code"].ToString();
                Initator_name = One_Row.Rows[0]["Initiator_Name"].ToString();
                Subject = One_Row.Rows[0]["Subject"].ToString();
                DateStart = DateTime.Parse( One_Row.Rows[0]["Date_Start"].ToString());
                DateEnd = DateTime.Parse( One_Row.Rows[0]["Date_End"].ToString());
                Sum_Contract = double.Parse(One_Row.Rows[0]["Sum_Contract"].ToString());
                Currency = One_Row.Rows[0]["Currency"].ToString();
                Order_Num = One_Row.Rows[0]["Order_Num"].ToString();
                Comment = One_Row.Rows[0]["Comments"].ToString();
                Email= One_Row.Rows[0]["Email_User"].ToString();
                if (One_Row.Rows[0]["Ensurance_Date"].ToString() != "")
                    Ensurance_Date = DateTime.Parse(One_Row.Rows[0]["Ensurance_Date"].ToString());
                else Ensurance_Date = DateTime.Today;
                Status = One_Row.Rows[0]["Status"].ToString();
                Contract_Num = int.Parse(One_Row.Rows[0]["Contract_Num"].ToString());
                Days_Alert = int.Parse(One_Row.Rows[0]["Days_Alert"].ToString());
                Days_Cancel = int.Parse(One_Row.Rows[0]["Days_For_Cancel"].ToString());
                Frequency_Alert = int.Parse(One_Row.Rows[0]["Frequency_Alert"].ToString());
                Months_Renew = int.Parse(One_Row.Rows[0]["Months_Renew"].ToString());
                Last_Update =DateTime.Parse( One_Row.Rows[0]["Last_Update"].ToString());
                Stop_Alert = One_Row.Rows[0]["Stop_Alert"].ToString();
                Bonus_Preiod = int.Parse(One_Row.Rows[0]["Bonus_Period"].ToString());
                //if (One_Row.Rows[0]["Bonus_Period"].ToString() != "")
                //{
                //    DateTime d1 = DateTime.Parse(One_Row.Rows[0]["Bonus_Period"].ToString());
                //    DateTime d2 = DateTime.Parse(DateEnd);
                //    Bonus_Preiod = ((d1.Month - d2.Month) + 12 * (d1.Year - d2.Year)).ToString();
                //    if (int.Parse(Bonus_Preiod) < 0) Bonus_Preiod = "0";
                //}
            }

            string pdf_file = "T:\\Contracts\\" + "חוזה " + Contract_id + ".pdf";
            string wrd_file = "T:\\Contracts\\" + "חוזה " + Contract_id + ".docx";
            if (File.Exists(pdf_file))
            {
                File_Name = pdf_file;
                return true;
            }
            if(File.Exists(wrd_file))
            {
                File_Name = wrd_file;
                return true;
            }
            return false;
        }


        /// <summary>
        /// רץ על שורת נתונים אחת עבור שליחת המיילים
        /// </summary>
        /// <param name="One_Row"></param>
        /// <returns></returns>
        public bool GetDataFromDT(DataRow One_Row)
        {
            Name_supplier = One_Row["supplier_code"].ToString();
            Initator_name = One_Row["Initiator_Name"].ToString();
            Subject = One_Row["Subject"].ToString();
            DateStart =DateTime.Parse( One_Row["Date_Start"].ToString());
            DateEnd = DateTime.Parse(One_Row["Date_End"].ToString());
            Sum_Contract = double.Parse(One_Row["Sum_Contract"].ToString());
            Currency = One_Row["Currency"].ToString();
            Order_Num = One_Row["Order_Num"].ToString();
            Comment = One_Row["Comments"].ToString();
            Email = One_Row["Email_User"].ToString();
            Status = One_Row["Status"].ToString();
            Contract_Num = int.Parse(One_Row["Contract_Num"].ToString());
            Days_Alert = int.Parse(One_Row["Days_Alert"].ToString());
            Days_Cancel = int.Parse(One_Row["Days_For_Cancel"].ToString());
            Frequency_Alert = int.Parse(One_Row["Frequency_Alert"].ToString());
            Months_Renew = int.Parse(One_Row["Months_Renew"].ToString());
            Last_Update = DateTime.Parse(One_Row["Last_Update"].ToString());
            Stop_Alert = One_Row["Stop_Alert"].ToString();
            Bonus_Preiod = int.Parse(One_Row["Bonus_Period"].ToString());
            Last_Date_Alert= DateTime.Parse(One_Row["Last_Date_Alert"].ToString());

            string pdf_file = "T:\\Contracts\\" + "חוזה " + Contract_id + ".pdf";
            string wrd_file = "T:\\Contracts\\" + "חוזה " + Contract_id + ".docx";
            if (File.Exists(pdf_file))
            {
                File_Name = pdf_file;
                return true;
            }
            if (File.Exists(wrd_file))
            {
                File_Name = wrd_file;
                return true;
            }
            return false;
        }


        public void DeleteContract()
        {
            string qry = $@"DELETE
                            FROM contract  
                            WHERE Cnum = {Contract_id}";
            DbServiceSQL sqlcombo = new DbServiceSQL();
            DataTable DT = new DataTable();
            sqlcombo.ExecuteQuery(qry);
            Delete_File();
        }




        /// <summary>
        /// 
        /// </summary>upload file for contract
        public string SaveContract()
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    File_Name = ofd.FileName;
                    //MessageBox.Show(FileName);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return File_Name;
        }

        //שמירת החוזה לתיקיית קבצים
        public void Copy_file()
        {

            if (File_Name != null)
            {
                //לקבצי word
                string subS = File_Name.Substring(File_Name.IndexOf(".") + 1).TrimEnd();
                FileInfo File_Contract = new FileInfo(File_Name);
                if (subS == "docx")
                {      
                    string docxPath = "T:\\Contracts\\" + "חוזה " + Contract_id + ".docx";
                    File_Contract.CopyTo(docxPath);

                    //אם ארצה לשמור כפידיאף
                    //SautinSoft.PdfMetamorphosis p = new SautinSoft.PdfMetamorphosis();
                    //string pdfPath = "T:\\Contracts\\" + "חוזה " + Contract_id + ".pdf";
                    //p.DocxToPdfConvertFile(docxPath, pdfPath);
                    //System.IO.File.Delete(docxPath);

                }
                else
                {
                    string path = "T:\\Contracts\\" + "חוזה " + Contract_id + ".pdf";
                    File_Contract.CopyTo(path);
                }
            }

        }

        public void file_attachment()
        {
            File_Name = "T:\\Contracts\\" + "חוזה " + Contract_id + ".pdf";
        }

        public bool Delete_File()
        {
            DialogResult d = MessageBox.Show("?להסיר את הקובץ המצורף", "אישור מחיקה", MessageBoxButtons.YesNo);
            if (d == DialogResult.Yes)
            {
                File_Name = "T:\\Contracts\\" + "חוזה " + Contract_id + ".pdf";
                System.IO.File.Delete(File_Name);
                File_Name = "T:\\Contracts\\" + "חוזה " + Contract_id + ".docx";
                System.IO.File.Delete(File_Name);
                return true;
            }
                return false;
        }

        /// <summary>
        /// אם החוזה מתחדש מוסיף כמות חודשים לתאריך סיום כל שנה
        /// </summary>
        public void Update_Year_Renew_Contract()
        {
            try
            {
                DbServiceSQL sqlcombo = new DbServiceSQL();
                DataTable DT = new DataTable();
                string update_date = $@"UPDATE contract
                            SET Date_End=(?)
                            WHERE Cnum={Contract_id}";
                SqlParameter date_End1 = new SqlParameter("f", DateEnd);
                sqlcombo.ExecuteQuery(update_date, CommandType.Text, date_End1);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// update last alert date on database
        /// </summary>
        public void Last_Mail_Alert()
        {
            DbServiceSQL sqlcombo = new DbServiceSQL();
            string last_date_alert = $@"update dbo.contract set Last_Date_Alert='{ DateTime.Today.ToString("yyyy-MM-dd")}'  where Cnum={Contract_id}";
            sqlcombo.ExecuteQuery(last_date_alert);
        }
    }

}
