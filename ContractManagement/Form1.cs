using System;
using System.Data;
using System.Windows.Forms;
using System.Net.Mail;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
using System.DirectoryServices.AccountManagement;
using System.Diagnostics;
using System.Drawing;

namespace ContractManagement
{
    public partial class Form1 : Form
    {
        public static int count_data = 0;
        int Cnum;//חוזה נבחר מהדטה גריד
        int Last_Cnum;
        string eMail_User = UserPrincipal.Current.EmailAddress;
        int row_for_update;
        int num_delete_row;
        int Contract_to_Delete;
        bool tabpage2_clicked;
        bool Clicked_CellMouseDoubleClick;//חיווי על עדכון רשומה ולא העלאת חוזה חדש
        string NameUser = Environment.UserName;//מקבל שם יוזר
        ArrayList list_supplier = new ArrayList();//רשימת שמות
        ArrayList list_sapak = new ArrayList();//רשימת קודי ספקים
        string file_copy;//שם קובץ להעלאה
        string file_name_update;//נתיב קובץ לעדכון
        bool exist_file;//בודק אם קובץ קיים,אם רוצים לעדכן הקובץ הקודם יימחק
        public Form1()
        {

            InitializeComponent();
            SupplierOrder_List();
            tabControl1.SelectedTab = tabPage1;
            Startscreen();
            tabControl1.Selecting += new TabControlCancelEventHandler(tabControl1_Selecting);//עבור אירוע מחיקת חוזה
            ShowData();
            Mail_Func();//בדיקת התראה
            ResizeMachiInfo();
        }



        private void Form1_Load(object sender, EventArgs e)
        {
                initSelectTabs();//מטעין בפתיחת התוכנית את הטבים בשביל לטעון מהר רשימת ספקים והזמנות

            //this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.Fixed3D; ;
            this.WindowState = FormWindowState.Maximized;
        }


        private void ResizeMachiInfo()
        {
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;           
            //Action<Control.ControlCollection> func = null;                  
                if (screenWidth <= 1920 && screenWidth > 1600)
                {                    
                    foreach (Control c in tabPage2.Controls)
                {
                        if (c.AccessibleName == "GroupMove")
                        {
                            int locationY = c.Location.Y;
                            int locationX = c.Location.X;
                            if (locationY < 450)
                                c.Location = new System.Drawing.Point(locationX - 200, locationY);
                        }
                      
                    }
                }
             

         
        }

        public void ResizeLBL(int Size)
        {
            if (Size == 1)
            {
                this.Height = 60;
                this.Width = 270;
                //NameLBL.Height = 19;
                //SizeLBL.Height = 19;
                //TimeLBL.Height = 19;
                //NameLBL.Font = new Font("Tahoma", 11, FontStyle.Bold);
                //SizeLBL.Font = new Font("Tahoma", 11, FontStyle.Bold);
                //TimeLBL.Font = new Font("Tahoma", 11, FontStyle.Bold);
            }
            else if (Size == 2)
            {
                this.Height = 38;
                this.Width = 250;
                //NameLBL.Height = 13;
                //SizeLBL.Height = 13;
                //TimeLBL.Height = 13;
                //NameLBL.Font = new Font("Tahoma", 9, FontStyle.Bold);
                //SizeLBL.Font = new Font("Tahoma", 9, FontStyle.Bold);
                //TimeLBL.Font = new Font("Tahoma", 9, FontStyle.Bold);
            }
        }



        public int GetLineNumber(Exception ex)
        {
            var lineNumber = 0;
            const string lineSearch = ":line ";
            var index = ex.StackTrace.LastIndexOf(lineSearch);
            if (index != -1)
            {
                var lineNumberText = ex.StackTrace.Substring(index + lineSearch.Length);
                if (int.TryParse(lineNumberText, out lineNumber))
                {
                }
            }
            return lineNumber;
        }


        private void initSelectTabs()
        {
            foreach (TabPage p in tabControl1.TabPages)
            {
                tabControl1.SelectedTab = p;
            }
            tabControl1.SelectedTab = tabPage1;

        }


        private void SupplierOrder_List()
        {
         contract For_SupplierOrder_List = new contract();
         DataTable s = For_SupplierOrder_List.SupplierList();//רשימת ספקים
      
         cbo_sapak.DataSource = s;
         cbo__supplier.DataSource = s;
         cbo_sapak.DisplayMember = "num";
         cbo__supplier.DisplayMember = "name";
      
         DataTable o = For_SupplierOrder_List.OrderList();//רשימת הזמנות
         cbo_ordernum.DataSource = o;
         cbo_ordernum.DisplayMember = "OrderNum";         
            
        }

        private DataTable  ShowData()//הצגת נתונים קיימים
        {
            DbServiceSQL sqlNow = new DbServiceSQL();//הצגה בגריד ויו של טבלה נוכחית
            DataTable dNow = new DataTable();
            string str;
            if (cb_show_active.Checked == false)
            {
                str = $@"SELECT t.Cnum as ' ',t.Supplier_Code as ' קוד ספק',t.Contract_Num 'מספר חוזה',case when t.Status=1 then 'פעיל' else 'לא פעיל' end 'סטטוס',
                          t.Initiator_Name 'שם יזם',t.Date_Start 'תאריך התחלה',t.Date_End 'תאריך סיום', t.Subject 'נושא', 
                          t.Months_Renew 'חודשי חידוש',t.Bonus_Period 'חודשי הטבה',t.Days_Alert 'מספר ימי התראה',t.Comments 'הערות'
                          FROM contract t";
            }
            else
            {
                 str = $@"SELECT t.Cnum as ' ',t.Supplier_Code as ' קוד ספק',t.Contract_Num 'מספר חוזה',case when t.Status=1 then 'פעיל' else 'לא פעיל' end 'סטטוס',
                          t.Initiator_Name 'שם יזם',t.Date_Start 'תאריך התחלה',t.Date_End 'תאריך סיום', t.Subject 'נושא', 
                          t.Months_Renew 'חודשי חידוש',t.Bonus_Period 'חודשי הטבה',t.Days_Alert 'מספר ימי התראה',t.Comments 'הערות'
                          FROM contract t
                          WHERE status=1";
            }
            dNow = sqlNow.executeSelectQueryNoParam(str);
            data_contract_view.DataSource = dNow;
            count_data = dNow.Rows.Count;//כמה רשומות יש בdb
            return dNow;
        }

        void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)//עבור לשונית מחיקת חוזה
        {
            TabPage current = (sender as TabControl).SelectedTab;
   
            if(e.TabPageIndex==2)
            {
                if (Contract_to_Delete != 0)
                {
                    DialogResult d = MessageBox.Show(" האם אתה בטוח שברצונך למחוק את החוזה"+ "\n                          ? " + data_contract_view.Rows[num_delete_row].Cells[7].Value.ToString(), "אישור מחיקה", MessageBoxButtons.YesNo);
                    if (d == DialogResult.Yes)
                    {
                        tabpage2_clicked = true;
                        contract c_delete = new contract(Contract_to_Delete);
                        c_delete.DeleteContract();              
                        ShowData();
                        MessageBox.Show("חוזה נמחק בהצלחה");     
                    }
                    tabControl1.SelectedTab = tabPage1;
                    if (d == DialogResult.No) Contract_to_Delete = 0;
                }
      
            }
            if (e.TabPageIndex == 1)
            {
                if(row_for_update==0)
                {
                    btn_save.Enabled = false;
                    btn_file.Enabled = false;
                }
            }
            if (e.TabPageIndex == 0 && tabpage2_clicked == true)//אם חוזרים לטב0 ולוחצים על רשומה ריקה ,שורת הקוד תמנע הודעת מחיקה(איפוס שורת דליט)
                Contract_to_Delete = 0;
            
        }

        private void Startscreen()
        {
            cbo_alarm.Items.Add("ימים");
            cbo_alarm.Items.Add("חודשים");
            cbo_cancel.Items.Add("ימים");
            cbo_cancel.Items.Add("חודשים");           
            cbo_frequency.Items.Add("פעם בשבוע");
            cbo_frequency.Items.Add("פעם בשבועיים");
            cbo_frequency.Items.Add("פעם בחודש");
            cbo_frequency.Items.Add("פעם בחצי שנה");
            cbo_monthrenew.Items.Add("שנה");
            cbo_monthrenew.Items.Add("שנתיים");
            cbo_monthrenew.Items.Add("שלוש שנים");
            cbo_currency.Items.Add("₪");
            cbo_currency.Items.Add("$");
            cActivate.Checked = true;
            cbo_currency.SelectedIndex = 0;
            cbo_cancel.SelectedIndex = 0;
            cbo_cancel.SelectedIndex = 0;
            cbo_alarm.SelectedIndex = 0;
            dateTimePicker_ensurance.Enabled = false;


        }




        private void button2_Click(object sender, EventArgs e)//save button
        {
            
            bool i = check_fields();
            if (i != true) return;
            bool email_ok = IsValidEmail(txt_mail.Text);
            if (email_ok != true) return;
            if (count_data != 0)
            {  Last_Cnum = int.Parse(data_contract_view.Rows[count_data - 1].Cells[0].Value.ToString()); }//משתנה עבור מספר חוזה אחרון בשביל לדעת מה הבא
            
                string activateOrNo;//שולח פעיל או לא פעיל
                if (cActivate.Checked)
                    activateOrNo = "true";
                else activateOrNo = "false";
  
            int bonus = 0;//בודק אם יש הטבה,מוסיף שנה או שמוציא ערך ריק
            if (txt_bonus.Text != "") bonus = int.Parse(txt_bonus.Text);


            DateTime ensurance;//אם היה שינוי בתאריך סימן שיש ביטוח אם לא שולח ערך ריק
            DateTime time_now= DateTime.Now.Date;
            DateTime ensurance_date = dateTimePicker_ensurance.Value.Date;
            if (time_now != ensurance_date)
                ensurance = dateTimePicker_ensurance.Value;
            else ensurance =DateTime.MinValue;

                int days_alert=0;//בודק אם סימן ימים או חודשים
                if (cbo_alarm.SelectedIndex == 1)
                    days_alert = int.Parse(txt_days.Text) * 30;
                else days_alert = int.Parse(txt_days.Text);

                int f = 0;//תדירות
                if (cbo_frequency.SelectedIndex == 0) f = 7;
                if (cbo_frequency.SelectedIndex == 1) f = 14;
                if (cbo_frequency.SelectedIndex == 2) f = 30;
                if (cbo_frequency.SelectedIndex == 3) f = 180;

                int month_renew = month_func();//כמה חודשי חידוש

                int days_cancel=0;
                if (cbo_cancel.SelectedIndex == 1)
                    days_cancel = int.Parse(txt_cancel.Text) * 30;
                if (cbo_cancel.SelectedIndex == 0)
                    if(txt_cancel.Text!="")
                    days_cancel = int.Parse(txt_cancel.Text);

                string stop_alert;
                if (cbo_stop_alert.Checked == false) stop_alert = "false";
                else stop_alert = "true";


                if (txt_num_contract.Text == "") txt_num_contract.Text = "0";
                if (txt_sumcontract.Text == "") txt_sumcontract.Text = "0";

            if (row_for_update > count_data)//אם לחץ על רשומה ריקה יהיה הוספת חוזה חדש
            {        
                Last_Cnum++;
                contract Contract = new contract        (Last_Cnum, cbo_sapak.Text, cbo_name.Text, txt_subject.Text, //שולח מספר חדש
                                                        date_start.Value, datetime_over.Value, double.Parse(txt_sumcontract.Text), cbo_currency.Text,
                                                        cbo_ordernum.Text, txt_comment.Text, ensurance, activateOrNo, int.Parse(txt_num_contract.Text), days_alert, days_cancel,
                                                        f, month_renew, bonus, date_start.Value, DateTime.Now,stop_alert, file_copy,txt_mail.Text);

                Contract.InsertContract();
                count_data++;
            }
            else //עדכון חוזה קיים
            {
                contract Contract = new contract        (Cnum, cbo_sapak.Text, cbo_name.Text, txt_subject.Text, //שולח מספר הרשומה
                                                        date_start.Value, datetime_over.Value, double.Parse(txt_sumcontract.Text), cbo_currency.Text,
                                                        cbo_ordernum.Text, txt_comment.Text, ensurance, activateOrNo, int.Parse(txt_num_contract.Text), days_alert, days_cancel,
                                                        f, month_renew, bonus, date_start.Value,DateTime.Now,stop_alert, file_copy, txt_mail.Text);


                Contract.UpdateContract();

            }

            tabControl1.SelectedTab = tabPage1;
            DbServiceSQL sqlcombo = new DbServiceSQL();
            DataTable DT = new DataTable();

            //הצגה מעודכנת לאחר הוספה
     
            ShowData();
            Clear_Data();
            
        }

        /// <summary>
        /// פונקציית שליחת התראות
        /// </summary>
        private void Mail_Func()
        {

            DbServiceSQL sqlcombo = new DbServiceSQL();
            DataTable for_mail_alert = new DataTable();
            string eMail_From = "Contract_Alert@atgtire.com";
            string str = $@"SELECT *
                                FROM contract t";
            for_mail_alert = sqlcombo.executeSelectQueryNoParam(str);
            DateTime date_alert;//תאריך התראה
            DateTime new_date_alert;//עבור התראות עם תדירות
            DateTime days_diff_cancel;//תאריך המקסימום שניתן לבטל חוזה מתחדש(אם לא מתחדש יהיה כמו תאריך סיום חוזה)

            for (int i = 0; i < for_mail_alert.Rows.Count; i++)
            {
                try
                {
                    contract mail_object = new contract(int.Parse(for_mail_alert.Rows[i]["Cnum"].ToString()));
                    mail_object.GetDataFromDT(for_mail_alert.Rows[i]);
                    string Email_For_Alert = mail_object.Email;
                    date_alert = mail_object.DateEnd.AddDays(-(double.Parse(mail_object.Days_Alert.ToString())));
                    days_diff_cancel = mail_object.DateEnd.AddDays(-(double.Parse(for_mail_alert.Rows[i]["Days_For_Cancel"].ToString())));

                    if (mail_object.Stop_Alert =="False")
                    {
                     
                        // התראת לפני סיום חוזה לפי מספר ימי התראה
                        if (DateTime.Today == date_alert && mail_object.DateEnd == days_diff_cancel && mail_object.Months_Renew == 0) //חוזה לא מתחדש,0 חודשי חידוש ולא הוספו ימים לdays diff cancel 
                        {
                            string subject = "התראת על סיום תקופת חוזה";
                            string body =
                                     $@"שים לב, 
                                        חוזה {mail_object.Subject} נגמר בתאריך {mail_object.DateEnd.ToString("dd/MM/yyyy")}
                                        לטיפולך. ";
                            Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From,subject,body);
                        }


                        ///התראה לפני סיום חוזה מתחדש לפי מספר ימי התראה
                        if (DateTime.Today == date_alert && mail_object.Months_Renew != 0) //רק אם התאריכים שונים מדובר בחוזה מתחדש
                        {
                            string subject = "התראה על סיום חוזה לקראת חידוש";
                            string body;
                            if (days_diff_cancel != mail_object.DateEnd)//המשתמש הזין ימי ביטול
                            {
                                body = $@"שים לב,
                                         חוזה {mail_object.Subject} מסתיים ב {mail_object.DateEnd.ToString("dd/MM/yyyy")}.
                                         החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.
                                         אם ברצונך לבטל את החוזה יש לבטלו עד התאריך { days_diff_cancel.ToString("dd/MM/yyyy")}.                       
                                         לטיפולך.";
                            }
                            else
                            {
                                body = $@"שים לב,
                                         חוזה {mail_object.Subject} מסתיים ב {mail_object.DateEnd.ToString("dd/MM/yyyy")}.
                                         החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.                      
                                         לטיפולך.";
                            }
                            Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From, subject, body);
                        }

                        ///-התראה לסיום חוזה לפי התדירות שנקבעה
                       //או חוזה מתחדש או חוזה מסתיים רגיל
                        new_date_alert = mail_object.Last_Date_Alert.AddDays(mail_object.Frequency_Alert);
                        if (mail_object.Last_Date_Alert < mail_object.DateEnd && DateTime.Today == new_date_alert && mail_object.Frequency_Alert!=0)//עדיין נצטרך להתריע
                        {
                            if (DateTime.Today >= mail_object.Last_Date_Alert && DateTime.Today < mail_object.DateEnd && days_diff_cancel == mail_object.DateEnd)//חוזה רגיל
                            {
                                string subject = "התראה על סיום חוזה-תזכורת נוספת";
                                string body= 
                                    $@"שים לב, 
                                    חוזה {mail_object.Subject} נגמר בתאריך {mail_object.DateEnd.ToString("dd/MM/yyyy")}
                                    לטיפולך. ";
                                Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From, subject, body);
                            }

                            if (DateTime.Today >= mail_object.Last_Date_Alert && DateTime.Today < mail_object.DateEnd && mail_object.Months_Renew != 0)//חוזה מתחדש
                            {
                                string subject= "התראה על התחדשות חוזה-תזכורת נוספת";
                                string body;
                                if (days_diff_cancel != mail_object.DateEnd)//המשתמש  הזין ימי ביטול
                                {
                                     body = $@"שים לב,
                                               חוזה {mail_object.Subject} מסתיים ב {mail_object.DateEnd.ToString("dd/MM/yyyy")}.
                                               החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.
                                               אם ברצונך לבטל את החוזה יש לבטלו עד התאריך { days_diff_cancel.ToString("dd/MM/yyyy")}.                       
                                               לטיפולך.";
                                }
                                else
                                {
                                    body= $@"שים לב,
                                               חוזה {mail_object.Subject} מסתיים ב {mail_object.DateEnd.ToString("dd/MM/yyyy")}.
                                               החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.            
                                               לטיפולך.";
                                }
                                Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From, subject, body);

                            }
                        }


                        ///התראה על חידוש חוזה ביום החידוש ועדכון תאריך סיום חוזה חדש 
                        int months = mail_object.Months_Renew;
                        int year = 0;
                        switch (months)
                        {
                            case 12:
                                year = 1;
                                break;
                            case 24:
                                year = 2;
                                break;
                            case 36:
                                year = 3;
                                break;
                        }
                        if (DateTime.Today == mail_object.DateEnd && mail_object.Months_Renew != 0 && mail_object.Status == "True")
                        {
                            mail_object.DateEnd = mail_object.DateEnd.AddYears(year);//להשלים את הוספת השנים
                            mail_object.Update_Year_Renew_Contract();
                            days_diff_cancel = days_diff_cancel.AddYears(year);
                            string subject = "התראה על התחדשות חוזה";
                            string body;
                            if (days_diff_cancel != mail_object.DateEnd)//המשתמש  הזין ימי ביטול
                            {
                                 body = $@"שים לב,
                                            חוזה {mail_object.Subject} התחדש היום {DateTime.Now.ToString("dd/MM/yyyy")}.
                                            החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.
                                            אם ברצונך לבטל בעתיד את החוזה יש לבטלו עד התאריך { days_diff_cancel.ToString("dd/MM/yyyy")}.                       
                                            לטיפולך.";
                            }
                            else
                            {
                                body = $@"שים לב,
                                            חוזה {mail_object.Subject} התחדש היום {DateTime.Now.ToString("dd/MM/yyyy")}.
                                            החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.                  
                                            לטיפולך.";
                            }
                            Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From, subject, body);
                        }

                        ///התראה על סיום חוזה ביום הסיום 
                        if (DateTime.Today ==mail_object.DateEnd && mail_object.Months_Renew == 0 && mail_object.Status == "True")
                        {
                            string subject= "התראה על סיום תקופת חוזה";
                            string body = $@"שים לב,
                                          חוזה {mail_object.Subject} הסתיים היום {DateTime.Now.ToString("dd/MM/yyyy")}.                           
                                          לידיעתך.";
                            Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From, subject, body);
                        }

                        ///התראה על יום ביטול אחרון
                        if (DateTime.Today == days_diff_cancel && mail_object.Months_Renew != 0 && mail_object.Status == "True")
                        {
                            string subject = "התראה על סיום ביטול חוזה";
                            string body = $@"שים לב,
                                          יום זה הינו האחרון שניתן לבטל את חוזה {mail_object.Subject} שמסתיים ב {mail_object.DateEnd.ToString("dd/MM/yyyy")}.
                                          החוזה הינו חוזה מתחדש לתקופה של {mail_object.Months_Renew} חודשים.                    
                                          לטיפולך.";
                            Mail_Alert_End_Contract(mail_object.Contract_id, Email_For_Alert, eMail_From, subject, body);
                        }
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        
        /// <summary>
        /// שליחת מייל עבור סיום חוזה
        private void Mail_Alert_End_Contract(int Cnum, string Email_For_Alert, string eMail_From, string sub,string body)
        {
            DbServiceSQL sqlcombo = new DbServiceSQL();
            contract mail_object = new contract(Cnum);
            mail_object.GetData();
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress(eMail_From);
            mail.To.Add(Email_For_Alert);
            mail.Subject = sub;
            mail.Body = body;
          

            if (mail_object.File_Name != null)
            {
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(mail_object.File_Name);
                mail.Attachments.Add(attachment);
            }

            SmtpClient client = new SmtpClient();
            client.Host = "almail";// ServerIP;
            client.Send(mail);
            client.Dispose();
            mail_object.Last_Mail_Alert();//שולח לפונקצייה ששומרת תאריך התראה אחרון


        }
        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                MessageBox.Show("כתובת המייל שהזנת לא נכונה");
                return false;
            }
        }

        private int month_func()
        {
            int month=0 ;
            if (cbo_monthrenew.SelectedIndex == 0)
                month = 12;
            if (cbo_monthrenew.SelectedIndex == 1)
                month = 24;
            if (cbo_monthrenew.SelectedIndex == 2)
                month = 36;
            return month;
                    

        }

        private bool check_fields()
        {
            if (cbo_name.Text == "") { MessageBox.Show("לא הכנסת שם יזם "); return false; }
            if (cbo_sapak.Text == ""){ MessageBox.Show("לא הכנסת קוד ספק "); return false; }
            if (txt_subject.Text == ""){ MessageBox.Show("לא הכנסת נושא "); return false; }
            if ( cActivate.Checked == false && cNoActivate.Checked == false) {MessageBox.Show("לא סימנת סטטוס חוזה "); return false; }
            if (cbo_alarm.Text ==""){ MessageBox.Show("לא בחרת מועד התראה רצוי במייל"); return false; }
            if (txt_days.Text == "") txt_days.Text = "0";
            if (date_start.Value >= datetime_over.Value) { MessageBox.Show("תאריך תחילת חוזה לא יכול להיות אחרי תאריך סיום חוזה"); return false; }
        
                try
                {
                    if(txt_cancel.Text!="")
                    int.Parse(txt_cancel.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("שדה "+label_cancel.Text+" יכול להכיל רק מספרים");
                    return false;
                }
                try
                {
                    if(txt_bonus.Text!="")
                    int.Parse(txt_bonus.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("שדה " + label12.Text + " יכול להכיל רק מספרים");
                    return false;
                }
                try
                {
                    if (txt_sumcontract.Text != "")
                        int.Parse(txt_sumcontract.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("שדה " + label_sum.Text + " יכול להכיל רק מספרים");
                    return false;
                }
                try
                {
                    if (cbo_ordernum.Text != "")
                        int.Parse(cbo_ordernum.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("שדה " + label_ordernum.Text + " יכול להכיל רק מספרים");
                    return false;
                }
                try
                {
                    if (txt_num_contract.Text != "")
                        int.Parse(txt_num_contract.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("שדה " + lbl_num_contract.Text + " יכול להכיל רק מספרים");
                    return false;
                }
            try
                {
                    if (cbo_cancel.Text != "ימים" && cbo_cancel.Text != "חודשים")
                       throw new Exception("הזן ימים/חודשים עבור"+label_cancel.Text);
                    if(cbo_alarm.Text != "ימים" && cbo_alarm.Text != "חודשים")
                       throw new Exception("הזן ימים/חודשים עבור" + label25.Text);
                    if (!cbo_monthrenew.Items.Contains(cbo_monthrenew.Text) && cbo_monthrenew.Text!="")
                        throw new Exception("לא הזנת כמות נכונה של תקופת חידוש");
                    if (!cbo_frequency.Items.Contains(cbo_frequency.Text) && cbo_frequency.Text!="")
                        throw new Exception("לא הזנת נכון תדירות הודעות");
                    if (!cbo_currency.Items.Contains(cbo_currency.Text) && cbo_currency.Text != "")
                        throw new Exception("בחר סימן שח או דולר");
                }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return false;
                    }

            return true;
            }


        private void cActivate_CheckedChanged(object sender, EventArgs e)
        {
            cNoActivate.Checked =false;
            if (cActivate.Checked == true) Enable_Fields();
        }

        private void cNoActivate_CheckedChanged(object sender, EventArgs e)
        {
            cActivate.Checked = false;
            if (cNoActivate.Checked == true) Enable_Fields();

        }

        private void Enable_Fields()
        {
            if (cNoActivate.Checked == true)
            {
                txt_bonus.Enabled = false;
                txt_cancel.Enabled = false;
                txt_comment.Enabled = false;
                txt_days.Enabled = false;
                txt_num_contract.Enabled = false;
                txt_subject.Enabled = false;
                txt_sumcontract.Enabled = false;
                cbo_name.Enabled = false;
                cbo_ordernum.Enabled = false;
                cbo_sapak.Enabled = false;
                cbo__supplier.Enabled = false;
                dateTimePicker_ensurance.Enabled = false;
                date_start.Enabled = false;
                datetime_over.Enabled = false;
                cbo_frequency.Enabled = false;
                cbo_currency.Enabled = false;
                cbo_monthrenew.Enabled = false;
                cbo_cancel.Enabled = false;
                cbo_alarm.Enabled = false;
                txt_mail.Enabled = false;
                dateTimePicker_ensurance.Enabled = false;
                cbo_ensurance.Enabled = false;
                
            }
            if (cActivate.Checked == true)
            {
                txt_bonus.Enabled = true;
                txt_cancel.Enabled = true;
                txt_comment.Enabled = true;
                txt_days.Enabled = true;
                txt_num_contract.Enabled = true;
                txt_subject.Enabled = true;
                txt_sumcontract.Enabled = true;
                cbo_name.Enabled = true;
                cbo_ordernum.Enabled = true;
                cbo_sapak.Enabled = true;
                cbo__supplier.Enabled = true;
                dateTimePicker_ensurance.Enabled = true;
                date_start.Enabled = true;
                datetime_over.Enabled = true;
                cbo_frequency.Enabled = true;
                cbo_currency.Enabled = true;
                cbo_monthrenew.Enabled = true;
                cbo_cancel.Enabled = true;
                cbo_alarm.Enabled = true;
                txt_mail.Enabled = true;
                dateTimePicker_ensurance.Enabled = true;
                cbo_ensurance.Enabled = true;
            }
            
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show(" לקראת סיום החוזה בחר כמות ימים\n" + "בה הינך מעוניין לקבל התראה במייל ");
        }

        private void btn_file_Click(object sender, EventArgs e)//צירוף קובץ
        {
            contract file_upload;//
            if (row_for_update > count_data)//יוסיף קובץ לחוזה חדש 
            {
                file_upload = new contract(int.Parse(data_contract_view.Rows[count_data - 1].Cells[0].Value.ToString()) + 1);
            }
            else//יוסיף קובץ לחוזה קיים
            {
                file_upload = new contract(int.Parse(data_contract_view.Rows[row_for_update-1].Cells[0].Value.ToString()));
            }
            file_copy = file_upload.SaveContract();
            string answer=Pdf_Or_Word(file_copy);
            if(answer=="pdf") 
                pbo_pdf.Visible = true;
            if (answer == "word")
                pbo_wrd.Visible = true;
            btn_DeleteFile.Visible = true;
            btn_DeleteFile.Enabled = true;
            btn_file.Enabled = false;//לא ניתן להעלות קובץ אחרי שהועלה



        }

        private string Pdf_Or_Word(string file_name)
        {
            string subS = file_name.Substring(file_name.IndexOf(".") + 1).TrimEnd();
            if (subS == "pdf")
                return "pdf";
            else return "word";
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" הזן מספר ימים מינימלי הדרוש להודעה  " +"\n "+" במידה ומעוניינים בביטול חידוש החוזה" );

        }



        private void button2_Click_2(object sender, EventArgs e)//כפתור מילוי שדות לבדיקות
        {
            //cbo_name.Text = "הדר שיפטן";
            //cbo_sapak.Text = "STAFILO INTERNATIONAL LTD";
            //txt_subject.Text = "סתם נושא";
            //txt_num_contract.Text = "2345";
            //date_start.Text = "29/08/2018 09:34";
            //datetime_over.Text = "29/08/2019 09:34";
            ////txt_ordernum.Text = "33333";
            //txt_sumcontract.Text = "1200";
            //cActivate.Checked = true;
            //txt_bonus.Text = "6";
            //txt_days.Text = "7";
            //cbo_frequency.SelectedIndex = 1;
            //txt_comment.Text = "מה קורה יגבר";

            cbo_name.Text = "איתמר סומר";
            cbo_sapak.Text = "06714324";
            txt_subject.Text = "שרות לרישיונות ה-Report Manager";
           
            date_start.Text = "01/05/2017 09:34";
            datetime_over.Text = "30/04/2018 09:34";
            //txt_ordernum.Text = "33333";
            //txt_sumcontract.Text = "8750";
            cActivate.Checked = true;
            txt_days.Text = "30";
            //cbo_frequency.SelectedIndex = 1;
            txt_comment.Text = "מוצר נלווה לקליק ויו. איתמר בדק באפריל 2017 חידוש או החלפת מוצר. החליט להשאיר - הועברה פניה לאיתמר בתחילת מאי 18";
            //cbo_ordernum.Text = "510155";



        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {        
                ShowData();    
        }


        private void cbo_sapak_SelectedIndexChanged(object sender, EventArgs e)//לאחר בחירת קוד ספק משלים את שם הספק
        {
            //cbo__supplier.Enabled = false;
            int num = cbo_sapak.SelectedIndex;
            for(int i=0;i<list_supplier.Count;i++)
            {
                if (num == i) cbo__supplier.Text = list_supplier[i].ToString();
            }
        }

        private void cbo__supplier_SelectedIndexChanged(object sender, EventArgs e)
        {
            int num = cbo__supplier.SelectedIndex;
            for (int i = 0; i < list_supplier.Count; i++)
            {
                if (num == i) cbo_sapak.Text = list_sapak[i].ToString();
            }
        }

        //פתיחת קובץ קיים
        private void pbo_pdf_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(file_name_update);
            
        }

        //ניקוי נתונים
        private void Clear_Data()
        {
            Action<Control.ControlCollection> func = null;

            func = (controls) =>
            {
                foreach (Control control in controls)
                {
                    if (control is TextBox)
                        (control as TextBox).Clear();
                    if (control is ComboBox)
                        (control as ComboBox).Text = "";
                    if (control is DateTimePicker)
                        (control as DateTimePicker).Text = "";
                    else
                        func(control.Controls);
                }
            };
            //pbo_pdf.Visible = true;
            btn_DeleteFile.Visible = false;
            cbo_stop_alert.Checked = false;
            cNoActivate.Checked = false;
            cActivate.Checked = true;
            dateTimePicker_ensurance.Enabled = false;

            func(Controls);
            cbo_currency.Text= "₪";
            cbo_cancel.Text="ימים";
            cbo_alarm.Text = "ימים";
            txt_mail.Text = eMail_User;//ממלא את המייל של היוזר במחשב
        }

        //delete file from existing contract
        private void btn_DeleteFile_Click(object sender, EventArgs e)
        {
            
            contract C = new contract(int.Parse(data_contract_view.Rows[row_for_update - 1].Cells[0].Value.ToString()));
            bool t=C.Delete_File();//t משתנה שבודק אם מחק קובץ
            if (t == true)
            {
                btn_file.Enabled = true;//ניתן להעלות קובץ חדש
                pbo_pdf.Visible = false;
                pbo_wrd.Visible = false;
                btn_DeleteFile.Enabled = false;
            }

        }


        /// <summary>
        /// סינון לפי תאריכים
        /// </summary>
        private void Filter()
        {
            if(cb_filter.Checked==true)
            {
                label13.Visible = true;
                label24.Visible = true;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                btn_show.Visible = true;
            }
            else
            {
                label13.Visible = false;
                label24.Visible = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                btn_show.Visible = false;
                ShowData();
            }
        }



        private void checkBox1_CheckedChanged_3(object sender, EventArgs e)
        {
            Filter();
        }

        private void btn_show_Click(object sender, EventArgs e)
        {
            DbServiceSQL sqlNow = new DbServiceSQL();//הצגה בגריד ויו של טבלה נוכחית
            DataTable dNow = new DataTable();
            string from_date = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            string until_date = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string str = $@"SELECT t.Cnum as ' ',t.Supplier_Code as ' קוד ספק',t.Contract_Num 'מספר חוזה',case when t.Status = 1 then 'פעיל' else 'לא פעיל' end 'סטטוס',
                          t.Initiator_Name 'שם יזם',t.Date_Start 'תאריך התחלה',t.Date_End 'תאריך סיום', t.Subject 'נושא', 
                          t.Months_Renew 'חודשי חידוש',t.Bonus_Period 'חודשי הטבה',t.Days_Alert 'מספר ימי התראה',t.Comments 'הערות'
                          FROM contract t
                          WHERE Date_End  between '{from_date}'and '{until_date}'";

            dNow = sqlNow.executeSelectQueryNoParam(str);
            data_contract_view.DataSource = dNow;
            data_contract_view.Sort(data_contract_view.Columns[2], System.ComponentModel.ListSortDirection.Descending);
            count_data = dNow.Rows.Count;//כמה רשומות יש בdb
        }

    

        private void btn_excel_Click(object sender, EventArgs e)
        {
         

        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }//שחרור פריטים באקסל

        private void data_contract_view_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.RowIndex < count_data && e.RowIndex >=0)
                    Contract_to_Delete = int.Parse((data_contract_view.Rows[e.RowIndex].Cells[0].Value.ToString()));
                num_delete_row = e.RowIndex;
            }
            catch
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void data_contract_view_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                tabControl1.SelectedTab = tabPage2;
                pbo_pdf.Visible = false;
                pbo_wrd.Visible = false;
                btn_DeleteFile.Visible = false;
                btn_save.Enabled = true;
                btn_file.Enabled = true;         
                row_for_update = e.RowIndex + 1;
                if (row_for_update <= count_data)//נלחץ על רשומה למטרת עדכון
                {
                    Clicked_CellMouseDoubleClick = true;//חיווי על עדכון רשומה ולא העלאת חוזה חדש
                    Cnum = int.Parse(data_contract_view.Rows[row_for_update - 1].Cells[0].Value.ToString());//מספר החוזה למציאה מתוך הdb
                    contract C = new contract(int.Parse(data_contract_view.Rows[row_for_update - 1].Cells[0].Value.ToString()));
                    exist_file = C.GetData();//עדכון רשומה+חוזה במידה וקיים

                    cbo_sapak.Text = C.Name_supplier;
                    cbo_name.Text = C.Initator_name;
                    txt_subject.Text = C.Subject;
                    date_start.Text = C.DateStart.ToShortDateString();
                    datetime_over.Text = C.DateEnd.ToShortDateString();
                    txt_sumcontract.Text = C.Sum_Contract.ToString();
                    cbo_currency.Text = C.Currency;
                    cbo_ordernum.Text = C.Order_Num;
                    txt_comment.Text = C.Comment;
                    txt_mail.Text = C.Email;
                    dateTimePicker_ensurance.Text = C.Ensurance_Date.ToShortDateString();
                    if (C.Status =="True") { cActivate.Checked = true;cActivate.Checked = true; Enable_Fields(); }
                    if (C.Status == "False") {cNoActivate.Checked = true; cNoActivate.Checked = true; }
                    txt_num_contract.Text = C.Contract_Num.ToString();
                    txt_days.Text = C.Days_Alert.ToString();
                    cbo_alarm.SelectedIndex = 0;
                    txt_cancel.Text = C.Days_Cancel.ToString();
                    cbo_cancel.SelectedIndex = 0;
                    if (C.Frequency_Alert == 7) cbo_frequency.Text = "פעם בשבוע";
                    else if (C.Frequency_Alert == 14) cbo_frequency.Text = "פעם בשבועיים";
                    else if (C.Frequency_Alert == 30) cbo_frequency.Text = "פעם בחודש";
                    else if(C.Frequency_Alert==180) cbo_frequency.Text = "פעם בחצי שנה";


                    if (C.Months_Renew == 12) cbo_monthrenew.SelectedIndex = 0;
                    else if (C.Months_Renew == 24) cbo_monthrenew.SelectedIndex = 1;
                    else if (C.Months_Renew == 36) cbo_monthrenew.SelectedIndex = 2;
                    if (C.Bonus_Preiod != 0) txt_bonus.Text = C.Bonus_Preiod.ToString();

                    //if (C.Bonus_Preiod != null) txt_bonus.Text = C.Bonus_Preiod;
                    lbl_LastUpdate.Text = C.Last_Update.ToLongDateString()+"  "+C.Last_Update.ToShortTimeString();
                    if (C.Stop_Alert == "True") cbo_stop_alert.Checked = true;
                    else { cbo_stop_alert.Checked = false; }

                    //אחרי העדכון יווצר אובייקט חדש לפי כפתור שמור

                    if (exist_file == true)
                    {
                       
                        btn_DeleteFile.Visible = true;
                        string answer = Pdf_Or_Word(C.File_Name);
                        if (answer == "pdf")
                        {
                            pbo_pdf.Visible = true;
                            pbo_pdf.Enabled = true;
                        }
                        if (answer == "word")
                        {
                            pbo_wrd.Visible = true;
                            pbo_wrd.Enabled = true;
                        }
                        btn_file.Enabled = false;
                        file_name_update = C.File_Name;

                    }
                }
                else
                {
                    Clear_Data();//ניקוי כל השדות במידה ולא ריקים 
                }
            }
        }

        private void data_contract_view_SortStringChanged(object sender, EventArgs e)
        {
            DataTable Data_Filter = new DataTable();
            Data_Filter = ShowData();
            Data_Filter.DefaultView.Sort = this.data_contract_view.SortString;
        }

        private void cbo_ensurance_CheckedChanged(object sender, EventArgs e)
        {
            if (cbo_ensurance.Checked == true) dateTimePicker_ensurance.Enabled = true;
            else dateTimePicker_ensurance.Enabled = false;
        }

        private void pbo_wrd_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(file_name_update);
        }

        private void pbo_wrd_Click_1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(file_name_update);

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void data_contract_view_FilterStringChanged(object sender, EventArgs e)
        {
            DataTable Data_Filter = new DataTable();
            Data_Filter = ShowData();
            Data_Filter.DefaultView.RowFilter = this.data_contract_view.FilterString;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Columns.AutoFit();



            xlWorkSheet.Cells[1, 1] = "Cnum";
            xlWorkSheet.Cells[1, 2] = "קוד ספק";
            xlWorkSheet.Cells[1, 3] = "מספר חוזה";
            xlWorkSheet.Cells[1, 4] = "סטטוס";
            xlWorkSheet.Cells[1, 5] = "שם יזם";
            xlWorkSheet.Cells[1, 6] = "תאריך התחלה";
            xlWorkSheet.Cells[1, 7] = "תאריך סיום";
            xlWorkSheet.Cells[1, 8] = "נושא";
            xlWorkSheet.Cells[1, 9] = "חודשי חידוש";
            xlWorkSheet.Cells[1, 10] = "חודשי הטבה";
            xlWorkSheet.Cells[1, 11] = "מספר ימי התראה";
            xlWorkSheet.Cells[1, 12] = "הערות";
            xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            xlWorkSheet.Range["A1:L1"].Interior.Color = System.Drawing.Color.Yellow;
            xlWorkSheet.Range["L1"].ColumnWidth = 120;
            xlWorkSheet.Range["E1:K1"].ColumnWidth = 11;
            xlWorkSheet.Range["H1"].ColumnWidth = 40;

            DataTable excel = new DataTable();
            DbServiceSQL sqlshow = new DbServiceSQL();
            string str = $@"SELECT t.Cnum as ' ',t.Supplier_Code as ' קוד ספק',t.Contract_Num 'מספר חוזה',t.Status 'סטטוס',   
                          t.Initiator_Name 'שם יזם',t.Date_Start 'תאריך התחלה',t.Date_End 'תאריך סיום', t.Subject 'נושא', 
                          t.Months_Renew 'חודשי חידוש',t.Bonus_Period 'חודשי הטבה',t.Days_Alert 'מספר ימי התראה',t.Comments 'הערות'
                          FROM contract t";
            excel = sqlshow.executeSelectQueryNoParam(str);
            for (int i = 0; i < excel.Rows.Count; i++)//רישום לאקסל מתוך sql
            {
                xlWorkSheet.Cells[i + 2, 1] = excel.Rows[i].ItemArray[0];
                xlWorkSheet.Cells[i + 2, 2] = excel.Rows[i].ItemArray[1];
                xlWorkSheet.Cells[i + 2, 3] = excel.Rows[i].ItemArray[2];
                xlWorkSheet.Cells[i + 2, 4] = excel.Rows[i].ItemArray[3];
                xlWorkSheet.Cells[i + 2, 5] = excel.Rows[i].ItemArray[4];
                xlWorkSheet.Cells[i + 2, 6] = excel.Rows[i].ItemArray[5];
                xlWorkSheet.Cells[i + 2, 7] = excel.Rows[i].ItemArray[6];
                xlWorkSheet.Cells[i + 2, 8] = excel.Rows[i].ItemArray[7];
                xlWorkSheet.Cells[i + 2, 9] = excel.Rows[i].ItemArray[8];
                xlWorkSheet.Cells[i + 2, 10] = excel.Rows[i].ItemArray[9];
                xlWorkSheet.Cells[i + 2, 11] = excel.Rows[i].ItemArray[10];
                xlWorkSheet.Cells[i + 2, 12] = excel.Rows[i].ItemArray[11];
            }

            xlApp.Visible = true;
            //xlWorkBook.SaveAs("C:\\Users\\hshiftan\\source\\repos\\ContractManagement\\.xls", Type.Missing, Type.Missing, Type.Missing,//שמירת קובץ
            //Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            //Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //Type.Missing);

            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();
            releaseObject(xlWorkSheet);//מחיקה
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void txt_comment_TextChanged(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void cbo_alarm_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txt_mail_TextChanged(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void txt_bonus_TextChanged(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void cbo_stop_alert_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cbo_monthrenew_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void txt_cancel_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker_ensurance_ValueChanged(object sender, EventArgs e)
        {

        }

        private void cbo_cancel_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void cbo_frequency_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void mail_Click(object sender, EventArgs e)
        {

        }

        private void label_cancel_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void txt_days_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
