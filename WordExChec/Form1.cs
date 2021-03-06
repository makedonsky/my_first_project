﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;
using iniLibrary;
using System.Data.SqlClient;
using System.Globalization;
using System.Diagnostics;
using Translate;


namespace WordExChec
{
    
    public enum Month { Января = 1, Февраля = 2, Марта = 3, Апреля = 4, Мая = 5, Июня = 6, Июля = 7, Августа = 8, Сентября = 9, Октября = 10, Ноября = 11, Декабря = 12 };
    public partial class Form1 : Form
    {
        string Basepath=Path.GetFullPath("./");
        DogovorInfo dogovor = new DogovorInfo();
        //Arg arguments = new Arg();
        Arg Predarguments = new Arg();
        Dictionary<string, string> clientsSerarch=new Dictionary<string,string>();
        Arg Mainarguments = new Arg();
        Dataview databox = new Dataview();
        SenderObj sendobject=new SenderObj();
        SortedDictionary<string, int> dict = new SortedDictionary<string, int>();
        string agent_key = "";
        public Form1()
        {
            InitializeComponent();
        }

       /* private void button1_Click(object sender, EventArgs e)
        {

                 object obj_App;
				 object obj_Doc;
				 object obj_Bookmarks;
                 object obj_Bookmark;
                 object obj_Selection;
                 object obj_Range;
                 object obj_tables;
            object[] Param;
            object[] ExcelParam;
            object[] Cells;
				// Nullable^ n;
				 //n=null;

				 Type obj_Class=Type.GetTypeFromProgID("Word.Application");
				 object word=Activator.CreateInstance(obj_Class);
                 Type obj_ClassExcel = Type.GetTypeFromProgID("Excel.Application");
                 object Excel = Activator.CreateInstance(obj_ClassExcel);
				 
				 obj_App=word.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, word, null);
                 obj_Doc = obj_App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj_App, null);
                 Param = new object[1];
                 Param[0] = @"C:\1.doc" ;
                 
                 object Doc = obj_Doc.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_Doc, Param);
                 Param[0] = "true";
                 obj_Doc = obj_App.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, obj_App, Param);
                 obj_Selection = obj_App.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, obj_App, null);
                 obj_Bookmarks=Doc.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, Doc, null);
                 Param[0] = "n1";
                 //obj_Bookmark = obj_Bookmarks.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, obj_Bookmarks, Param);
                 //obj_Bookmark.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, obj_Bookmark,null);
                 //obj_Range = obj_App.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, obj_App, null);
                 Param[0]="12312";
                 //obj_Range.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, obj_Range, Param);
                 //object text=obj_Range.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, obj_Range, null);
                 //Param[0] = @"C:\Doc2.doc";
                 //Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, Param);
                 Param[0] = 1;
                 obj_tables = Doc.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, Doc, null);
                 object table = obj_tables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null,obj_tables, Param);
                 Cells = new object[2];
                 Cells[0] = 1;
                 Cells[1] = 2;
                 object cell = table.GetType().InvokeMember("Cell", BindingFlags.InvokeMethod,null,table,Cells);
                 object Range = cell.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cell, null);
                 object text = Range.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, Range, null);
            //Doc.GetType().InvokeMember("Save", BindingFlags.InvokeMethod,
            //obj_Doc1 = obj_Doc.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, obj_Doc, null);
            
                 //obj_Bookmark = obj_Doc.GetType().InvokeMember("Selections", BindingFlags.GetProperty, null, obj_Doc, null);
            //Excel
                 ExcelParam = new object[1];
                 ExcelParam[0] = "true";
                 Excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null,Excel,ExcelParam);
                 object obj_wboks = Excel.GetType().InvokeMember("workbooks", BindingFlags.GetProperty,null,Excel,null);
                 object obj_wbok = obj_wboks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, obj_wboks, null);
                 object obj_worksheets = obj_wbok.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_wbok, null);
                 ExcelParam[0] = 1;
                 object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
                 ExcelParam[0] = "A1";
                 object targetcell = obj_worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,null,obj_worksheet,ExcelParam);
                 ExcelParam[0] = text;
                 targetcell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null,targetcell,ExcelParam);
        }*/

        private void Form1_Load(object sender, EventArgs e)
        {
            
            try
            {
                DirClear("Temp");
            }
            catch
            {
                this.richTextBox1.AppendText("ошибка очистки каталога \n\r");
            }
            //group1
            // Date
            calendarInitialize();
            inicializedict();
            boxaviaCompanyInitialize();
            //country

            //country
            //additem("[A+A Prague's Apartament's 3* Praha 8]Podlipneho 810/14, Praha 8, +420-602-322-600");
            //managers_load
            this.comboBox27.Items.Add("Все");
            this.comboBox30.Items.Add("Все");
           // this.comboBox27.Items.AddRange(manag);
            
            List<string> m_list = getmanagerList("Client");
            if (m_list.Count != 0)
            {
                foreach (string manager in m_list)
                {
                    this.comboBox5.Items.Add(manager);
                    this.comboBox14.Items.Add(manager);
                    this.comboBox27.Items.Add(manager);
                    this.comboBox30.Items.Add(manager);
                }
            }
            else
            {
             object[] manag = new object[] {
            "Семенова Н.А",
            "Дулебова Е.В",
            "Тищенко Е.С",
            "Малий Е.В",
            "Алхутова К.Г",
            "Пономарцева К.Д",
            "Кирилюк К.В",
            "Саяпина Н.Н",
            "Бахтуридзе В.В",
            "Чистякова А.В",
            "Бровко Л.Ю",
            "Ходокина Е.В",
            "Елисеева Л.В",
            "Пономарцева К.Д",
            "Семыкина Ю.С",
            "Пащинская Т.Е",
            "Дьякова Е.Е",
            "Яковлева И.С"};

                this.comboBox5.Items.AddRange(manag);
                this.comboBox14.Items.AddRange(manag);
                this.comboBox27.Items.AddRange(manag);
                this.comboBox30.Items.AddRange(manag);
            }
            /*object[] managsearch = new object[] {
            "Все",
            "Зелинская Е.И",
            "Дулебова Е.В",
            "Тищенко Е.С",
            "Малий Е.В",
            "Данчук Н.Н",
            "Бровко Л.Ю",
            "Ходокина Е.В",
            "Елисеева Л.В",
            "Пономарцева К.Д",
            "Пащинская Т.Е",
            "Дьякова Е.Е",
            "Буренко М.М"};*/
            this.comboBox27.SelectedItem = "Все";
            this.comboBox30.SelectedItem = "Все";
            //avia
            m_list = getmanagerList("Avia");
            if (m_list.Count != 0)
            {
                foreach (string manager in m_list)
                {
                    this.comboBox31.Items.Add(manager);
                    
                }
            }
            else
            {
                object[] manag = new object[] {
            "Батычко К.В"
             };
                this.comboBox31.Items.Add(manag);
            }
            //avia
            //managers_load
            //turoperators_load
            List<string> t_list = getToperatorsList();
            if (t_list.Count != 0)
            {
                //this.comboBox3.Items.Add("Росинтур");
                //this.comboBox16.Items.Add("Росинтур");
                foreach (string to in t_list)
                {
                    this.comboBox3.Items.Add(to);
                    this.comboBox16.Items.Add(to);
                    //this.comboBox27.Items.Add(to);
                    //this.comboBox30.Items.Add(to);
                }
                this.comboBox3.Items.Add("Магазин Путешествий");
                this.comboBox16.Items.Add("Магазин Путешествий");
                if (!comboBox3.Items.Contains("Росинтур"))
                {
                    this.comboBox3.Items.Add("Росинтур");
                    this.comboBox16.Items.Add("Росинтур");
                }
            }
            else
            {
                object[] turoper = new object[] {
            "Anextour",
            "Coral",
            "Pegas",
            "TezTour",
            "Росинтур",
            "Магазин Путешествий",
            "Intourist",
            "Labirinth",
            "Натали",
            "Ланта_тур_вояж",
            "Тур_Транс_Вояж",
            "Дельфин",
            "Аврора_Интур",
            "АлеанСПА",
            "ОРИЕНТ",
            "Таис",
            "Робинзон_Турс",
            "Мондиаль",
            "Нордлайн",
            "Sunmar_Tour",
            "СпутникЮг",
            "Круиз",
            "Альбион_Тур",
            "Панорама_21_век",
            "Здоровый_мир_Сочи",
            "КРИПТОН_ЮГ",
            "АКВА_Абаза",
            "КвайтДон",
            "МУЛЬТИТУР",
            "ПАКТУР",
            "Чайна_Трэвел",
            "Ривьера_Сочи"};
                this.comboBox3.Items.AddRange(turoper);
                //turoperators__load
                //pred_teuroper-load
                this.comboBox16.Items.AddRange(turoper);
                //pred_teuroper-load
            }
            //turoperators__load
            //country_load
            List<string> c_list = getcountryList();
            if (c_list.Count != 0)
            {
                this.comboBox26.Items.Add("Все");
                foreach (string c in c_list)
                {
                    this.comboBox26.Items.Add(c);
                    this.comboBox28.Items.Add(c);
                    this.comboBox29.Items.Add(c);
                    this.comboBox37.Items.Add(c);
                    //this.comboBox27.Items.Add(to);
                    //this.comboBox30.Items.Add(to);
                }
                comboBox26.SelectedItem = "Все";
                comboBox28.SelectedItem = "Россия";
                comboBox29.SelectedItem = "Россия";
                comboBox37.SelectedItem = "Италия";
            }
            else
            {
                object[] countrys = new object[] {
            "Все",
            "Россия",
            "Италия",
            "Чехия",
            "Греция",
            "Франция",
            "Болгария",
            "Египет",
            "Венгрия",
            "Испания",
            "ОАЭ",
            "Великобритания",
            "Турция",
            "Тайланд",
            "Израиль",
            "Доминикана",
            "Индия",
            "Индонезия",
            "Мальдивы"
            };
                object[] countrystwoV = new object[] {
            "Россия",
            "Италия",
            "Чехия",
            "Греция",
            "Франция",
            "Болгария",
            "Египет",
            "Венгрия",
            "Испания",
            "ОАЭ",
            "Великобритания",
            "Турция",
            "Тайланд",
            "Израиль",
            "Доминикана",
            "Индия",
            "Индонезия",
            "Мальдивы"
            };
                this.comboBox26.Items.AddRange(countrys);
                this.comboBox26.SelectedItem = "Все";
                this.comboBox28.Items.AddRange(countrystwoV);
                this.comboBox29.Items.AddRange(countrystwoV);
                //this.comboBox37.Items.AddRange(countrystwoV);
                // comboBox37.SelectedItem = "Италия";
                // comboBox37.
                comboBox28.SelectedItem = "Россия";
                comboBox29.SelectedItem = "Россия";
                object[] countrystwoA = new object[] {
            "Италия",
            "Чехия",
            "Греция",
            "Франция",
            "Болгария",
            "Египет",
            "Венгрия",
            "Испания",
            "ОАЭ",
            "Великобритания",
            "Турция",
            "Тайланд",
            "Израиль",
            "Доминикана",
            "Индия",
            "Индонезия",
            "Мальдивы",
            "Россия"
            };
                this.comboBox37.Items.AddRange(countrystwoA);
                comboBox37.SelectedItem = "Италия";
            }
            //country_load
            if (File.Exists("Dog1.ico"))
            {
                Icon ico = new Icon("Dog1.ico");
                this.Icon = ico;
            }
            DateTime date = DateTime.Now.Date;
            this.comboBox1.SelectedItem = date.Day.ToString();
            this.comboBox2.SelectedItem = ((Month)(date.Month)).ToString();
            this.comboBox3.SelectedItem = "Росинтур";
            this.comboBox4.SelectedItem = "РосинтурЮг";
            this.numericUpDown1.Value = date.Year;
            //this.numericUpDown3.Value = 100;
            //this.numericUpDown4.Value = date.Year;
            this.dataGridView8.RowCount = 1;
            this.dataGridView8.Rows[0].HeaderCell.Value = "Всего по заявке";
            //this.dataGridView8.Rows[1].HeaderCell.Value = "Отметки о платежах";
            this.dataGridView8.RowHeadersWidth = 180;
            this.comboBox19.SelectedItem = "Наличный";
            dataGridView8.ClearSelection();
            //group2

            //DataGridViewColumn col = this.dataGridView1.Columns[5];
            // DataGridViewComboBoxColumn c = (DataGridViewComboBoxColumn)this.dataGridView1.Columns[4];
            //c.
            // DataGridViewComboBoxCell tr = (DataGridViewComboBoxCell)this.dataGridView1.Rows[0].Cells[4];
            //tr.se
            //pred
            //
            //pasp_checkbox
            checkBox21.Checked = true;

            //pasp_checkbox
            dataGridView9.Rows[0].Cells[0].Value = false;
            dataGridView9.Rows[0].Cells[1].Value = false;
            dataGridView9.Rows[0].Cells[2].Value = false;
            ///
            this.numericUpDown5.Value = date.Year;
            this.numericUpDown6.Value = date.Year;
            this.comboBox11.SelectedItem = date.Day.ToString();
            this.comboBox10.SelectedItem = ((Month)(date.Month)).ToString();
            this.dataGridView15.RowCount = 1;
            this.dataGridView15.Rows[0].HeaderCell.Value = "Всего по заявке";
            //this.dataGridView8.Rows[1].HeaderCell.Value = "Отметки о платежах";
            this.dataGridView15.RowHeadersWidth = 180;
            dataGridView15.ClearSelection();
            //group2
            this.comboBox16.SelectedItem = "Росинтур";
            this.comboBox15.SelectedItem = "РосинтурЮг";
            this.comboBox17.SelectedItem = "Наличный";
            dataGridView5.RowCount = 2;
            dataGridView9.RowCount = 2;
            //
            //konsulDogovor
            this.numericUpDown3.Value = date.Year;
            this.comboBox8.SelectedItem = date.Day.ToString();
            this.comboBox7.SelectedItem = ((Month)(date.Month)).ToString();
            comboBox25.SelectedItem = "Клиентский";
            comboBox24.SelectedItem = "ДА, групповой(аэропорт-отель-аэропорт)";
            comboBox20.SelectedIndex = 0;
            textBox56.Text = "Ростов-Прага-Ростов";
            //konsulDogovor
            //aviaDogovor
            this.checkBox59.Checked = true;
            this.comboBox34.SelectedItem = date.Day.ToString();
            this.comboBox33.SelectedItem = ((Month)(date.Month)).ToString();
            this.numericUpDown7.Value = date.Year;
            this.comboBox31.SelectedItem = "Никонорова К.В";
            this.comboBox35.SelectedItem = "Росинтур";
            //textBox167.Text = "А/Б ";
            textBox168.Text = "невозвратный и не меняемый";
            //
            //Form1.ActiveForm.VerticalScroll.Value = 0;
            //this.AutoScroll = false;
            //this.VerticalScroll.Value = 0;
            //this.VScroll = true;
            //this.AutoScroll = true;
            //aviaDogovor

            //confirmation
            this.numericUpDown9.Value = date.Year;
            this.comboBox44.SelectedItem = date.Day.ToString();
            this.comboBox43.SelectedItem = ((Month)(date.Month)).ToString();
            comboBox52.SelectedItem = "Росинтур";
            //this.textBox184.Text = "15";
           //this.textBox216.Text = "3";
            
            /*object[] Amanag = new object[] {
            "Малий Е.В",
            "Бровко Л.Ю",
            "Дьякова Е.Е",
            "Чумакова О.В"};
            this.comboBox41.Items.AddRange(Amanag);*/
            List<string> managers = getmanagerList("Agent");
            if (managers.Count != 0)
            {
                foreach (string manager in managers)
                {
                    this.comboBox41.Items.Add(manager);
                }
            }
            numericUpDown14.Value = 20;
            dataGridView27.RowCount = 1;
            dataGridView29.RowCount = 2;
            this.dataGridView27.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView27_CellEndEdit);
            dataGridView30.RowCount = 2;
            this.dataGridView30.Rows[0].HeaderCell.Value = "НДС";
            this.dataGridView30.Rows[1].HeaderCell.Value = "Всего к оплате";
            dataGridView30.ClearSelection();
            //this.dataGridView30.Rows[1].HeaderCell.Value = "Всего к оплате";
            //EventArgs e=new EventArgs();
            numericUpDown12_ValueChanged((object)numericUpDown12, e);
            //confirmation
            
        }
        private void Getmanagers_db()
        {
           object[] a = new object[5];//a.add
        }

        private void calendarInitialize()
        {
            ///calendar
            // cal;
            this.m_calendar = new System.Windows.Forms.MonthCalendar();
            this.m_calendar.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.m_click);

            //calPanel
            this.calPanel = new System.Windows.Forms.Panel();
            if (!this.Controls.Contains(this.calPanel))
            {
                this.Controls.Add(this.calPanel);
            }
            this.calPanel.Controls.Add(this.m_calendar);
            this.calPanel.BringToFront();
            this.calPanel.Size = this.m_calendar.Size;
            this.calPanel.Hide();
            //
            /////calendar
        }
        private List<string> getToperatorsList()
        {
            List<string> operators = new List<string>();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "select id,to_name,to_shortname from touroperators";
            try
            {
                SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);

                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            operators.Add(reader["to_shortname"].ToString());
                        }
                    }
                }
                connect.Close();
            }
            catch
            {

            }
            return operators;
        }
        private List<string> getmanagerList(string type)
        {
            List<string> managers = new List<string>();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "";
            if (type=="Agent")
            {
                query = "select name from managers where ManagerType='agent'";
            }
            else if (type=="Client")
            {
                query = "select name from managers where ManagerType<>'avia'";
            }
            else if (type == "Avia")
            {
                query = "select name from managers where ManagerType='avia'";
            }
            try
            {
                SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);

                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            managers.Add(reader["name"].ToString());
                        }
                    }
                }
                connect.Close();
            }
            catch
            {

            }
            return managers;
        }
        private List<string> getcountryList()
        {
            List<string> countrys = new List<string>();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "select id,Runame from country";
            try
            {
                SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);

                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            countrys.Add(reader["Runame"].ToString());
                        }
                    }
                }
                connect.Close();
            }
            catch
            {

            }
            return countrys;
        }
        private void boxaviaCompanyInitialize()
        {
            ///calendar
            // cal;
            this.m_avia_company = new System.Windows.Forms.ComboBox();
            this.m_avia_company.Size = new Size(168, 21);
            object[] aviacompany = new object[]{
            "Донавиа",
            "Сибирь",
            "Люфтганза",
            "Австрийские Авиалинии",
            "Аэрофлот",
            "Руслайн",
            "Чешские Авиалинии"
            };
            this.m_avia_company.SelectedIndexChanged += new EventHandler(m_avia_company_SelectedIndexChanged);
            this.m_avia_company.Items.AddRange(aviacompany);

            //calPanel
            this.aviaCPanel = new System.Windows.Forms.Panel();
            if (!this.Controls.Contains(this.aviaCPanel))
            {
                this.Controls.Add(this.aviaCPanel);
            }
            this.aviaCPanel.Controls.Add(this.m_avia_company);
            this.aviaCPanel.BringToFront();
            this.aviaCPanel.Size = this.m_avia_company.Size;
            this.aviaCPanel.Hide();
            //
            /////calendar
        }

        void m_avia_company_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataGridView data = (DataGridView)databox.owner;
            ComboBox c = (ComboBox)sender;
            data.Rows[databox.args.RowIndex].Cells[databox.args.ColumnIndex].Value = c.Text;
            this.aviaCPanel.Hide();
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void DirClear(string path)
        {
            if (Directory.Exists(Basepath + path))
            {
                string[] files = Directory.GetFiles(path);
                foreach (string f in files)
                {
                    File.Delete(Basepath+f);
                }
            }
        }
        private string GetTempTemlate(string l_path,  string file)
        {
            
            string result = ""; int i_flag = 0; int count = 0;
            if (Directory.Exists(Basepath + "Temp\\"))
            {
                while (i_flag != 1)
                {
                    if (!File.Exists(Basepath + "Temp\\"+ count+ "_" + file))
                    {
                        File.Copy((Basepath + l_path + "\\" + file), (Basepath + "Temp\\" + count + "_" + file));
                        if (File.Exists(Basepath + "Temp\\" + count + "_" + file))
                        {
                            result = Basepath + "Temp\\" + count + "_" + file;
                            i_flag = 1;
                        }
                    }
                    else
                    {
                        count++;
                    }
                }
                
            }
            else
            {
                result = Basepath + l_path + "\\" + file;
            }
            //MessageBox.Show(result);
            return result;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            button14.Visible = true;

            object obj_App;
            object obj_Doc;
            object obj_Bookmarks;
            //object obj_Bookmark;
            //object obj_Selection;
            //object obj_Range;
            object obj_Tables;

            string currency = "";
            double zRubSum=0;
            object[] Param;
            string check="";
            string transport="";
            Param = new object[1];
            string save_param = "";

                Type obj_Class = Type.GetTypeFromProgID("Word.Application");
                object Word = Activator.CreateInstance(obj_Class);

                obj_App = Word.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, Word, null);
                obj_Doc = obj_App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj_App, null);
                if ((comboBox3.Text != "Росинтур")&&(comboBox3.Text != "Магазин Путешествий"))
                {
                   /* if (comboBox4.Text == "РосинтурЮг")
                    {
                        Param[0] = Basepath + @"Template\shablonUgA.doc";
                    }
                    if (comboBox4.Text == "Магазин Путешествий")
                    {
                        Param[0] = Basepath + @"Template\shablonTravelMagA.doc";
                    }*/
                    if (comboBox4.Text == "РосинтурЮг")
                    {
                        Param[0] = GetTempTemlate("Template","shablonUgA.doc");
                    }
                    if (comboBox4.Text == "Магазин Путешествий")
                    {
                        Param[0] = GetTempTemlate("Template","shablonTravelMagA.doc");
                    }
                }
                else
                {
                    /*if (comboBox3.Text == "Росинтур")
                    {
                        Param[0] = Basepath + @"Template\shablonRosintourO.doc";
                    }
                    if (comboBox3.Text == "Магазин Путешествий")
                    {
                        Param[0] = Basepath + @"Template\shablonTravelMagO.doc";
                    }*/
                    if (comboBox3.Text == "Росинтур")
                    {
                        Param[0] = GetTempTemlate("Template","shablonRosintourO.doc");
                    }
                    if (comboBox3.Text == "Магазин Путешествий")
                    {
                        Param[0] = GetTempTemlate("Template","shablonTravelMagO.doc");
                    }
                }
                object Doc = obj_Doc.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_Doc, Param);
                obj_Bookmarks = Doc.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, Doc, null);
                SetBookMarkText("DayNow", obj_Bookmarks, obj_App, this.comboBox1.SelectedItem.ToString());
                SetBookMarkText("MonthNow", obj_Bookmarks, obj_App, this.comboBox2.SelectedItem.ToString());
                SetBookMarkText("YearNow", obj_Bookmarks, obj_App, this.numericUpDown1.Value.ToString());
                SetBookMarkText("FIO", obj_Bookmarks, obj_App, this.comboBox6.Text.ToString() + " ");
                if ((comboBox3.Text != "Росинтур") && (comboBox3.Text != "Магазин Путешествий"))
                {
                    SetBookMarkText("Tyroperator", obj_Bookmarks, obj_App, this.comboBox3.Text);
                }
                else if (comboBox3.Text == "Росинтур")
                {
                    Touroperator to_rosin = new Touroperator();
                    to_rosin.getinfo(GetConnectSTR(), "Росинтур");
                    if ((to_rosin.to_id != null) && (to_rosin.to_id != ""))
                    {
                        SetBookMarkText("rosin_addr", obj_Bookmarks, obj_App, to_rosin.to_adress);
                        SetBookMarkText("rosin_insaddr", obj_Bookmarks, obj_App, to_rosin.ins_adress);
                        SetBookMarkText("rosin_insfrom", obj_Bookmarks, obj_App, to_rosin.ins_d_sdate);
                        SetBookMarkText("rosin_insname", obj_Bookmarks, obj_App, to_rosin.ins_name);
                        SetBookMarkText("rosin_insnum", obj_Bookmarks, obj_App, to_rosin.ins_d_num);
                        SetBookMarkText("rosin_insnumdate", obj_Bookmarks, obj_App, to_rosin.ins_d_date);
                        SetBookMarkText("rosin_insto", obj_Bookmarks, obj_App, to_rosin.ins_d_edate);
                        SetBookMarkText("rosin_rnum", obj_Bookmarks, obj_App, to_rosin.to_rn);
                        SetBookMarkText("rosin_insaddr2", obj_Bookmarks, obj_App, to_rosin.ins_name + "(" + to_rosin.ins_adress+")");
                    }
                }
                SetBookMarkText("ManNum", obj_Bookmarks, obj_App, this.textBox3.Text);
                SetBookMarkText("TravelProgram", obj_Bookmarks, obj_App, this.textBox2.Text + " ");
                SetBookMarkText("Travelstart", obj_Bookmarks, obj_App, this.dateTimePicker3.Text);
                SetBookMarkText("TravelEnd", obj_Bookmarks, obj_App, this.dateTimePicker4.Text);
                SetBookMarkText("TravelPlace", obj_Bookmarks, obj_App, this.textBox6.Text);
                SetBookMarkText("VizaDays", obj_Bookmarks, obj_App, this.textBox48.Text);
                if (this.checkBox72.Checked != true)
                {
                    SetBookMarkText("Sp", obj_Bookmarks, obj_App, "");
                }
                if (this.checkBox1.Checked) { check = "Да"; } else { check = "Нет"; }
                SetBookMarkText("checkbox1", obj_Bookmarks, obj_App, check);
                if (this.checkBox2.Checked) { check = "Да"; } else { check = "Нет"; }
                SetBookMarkText("checkbox2", obj_Bookmarks, obj_App, check);
                if (this.checkBox3.Checked) { check = "Да"; } else { check = "Нет"; }
                SetBookMarkText("checkbox3", obj_Bookmarks, obj_App, check);
                obj_Tables = Doc.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, Doc, null);
                if (dataGridView14.RowCount > 3)
                {
                    TableSize(dataGridView14, obj_Tables, 1,3);
                }
                TableProcess(dataGridView14, obj_Tables, 1);
                if (dataGridView1.RowCount > 3)
                {
                    TableSize(dataGridView1, obj_Tables, 2,3);
                }
                TableProcess(dataGridView1, obj_Tables, 2);
                if (dataGridView2.RowCount > 3)
                {
                    TableSize(dataGridView2, obj_Tables, 3,3);
                }
                TableProcess(dataGridView2, obj_Tables, 3);

                if (this.checkBox4.Checked) { transport = "Авиа"; }
                if (this.checkBox5.Checked) { transport = "Ж\\д "; }
                if (this.checkBox6.Checked) { transport = "Авто"; }
                //SetBookMarkText("Transport", obj_Bookmarks, obj_App, waycheck);
                SetBookMarkText("Transport", obj_Bookmarks, obj_App, transport);
                if (dataGridView3.RowCount > 3)
                {
                    TableSize(dataGridView3, obj_Tables, 4,3);
                }
                TableProcess(dataGridView3, obj_Tables, 4);
                if (dataGridView4.RowCount > 2)
                {
                    TableSize(dataGridView4, obj_Tables, 5, 2);
                }
                TableProcess(dataGridView4, obj_Tables, 5);
                TableProcessCheck(dataGridView5, obj_Tables, 6);
                //SetTableItemText(obj_Tables, 5, 3, 1, "sdfsd");
                //reqvizits
                if ((comboBox3.Text != "Росинтур") && (comboBox3.Text != "Магазин Путешествий"))
                {
                    if (this.comboBox4.SelectedItem != null)
                    {
                        Touroperator to=new Touroperator();
                        to.getinfo(GetConnectSTR(),this.comboBox3.SelectedItem.ToString());
                        if ((to.to_id!=null)&&(to.to_id!=""))
                        {
                            SetBookMarkText("to_name", obj_Bookmarks, obj_App, to.to_name);
                            SetBookMarkText("to_reestr_num", obj_Bookmarks, obj_App, to.to_rn);
                            SetBookMarkText("to_adr", obj_Bookmarks, obj_App, to.to_adress);
                            SetBookMarkText("to_tel", obj_Bookmarks, obj_App, to.to_tel);
                            SetBookMarkText("to_fax", obj_Bookmarks, obj_App, to.to_fax);
                            SetBookMarkText("to_fin_cap", obj_Bookmarks, obj_App, to.ins_fin_cap);
                            SetBookMarkText("to_ins_adr", obj_Bookmarks, obj_App, to.ins_adress);
                            SetBookMarkText("to_ins_d_date", obj_Bookmarks, obj_App, to.ins_d_date);
                            SetBookMarkText("to_ins_edate", obj_Bookmarks, obj_App, to.ins_d_edate);
                            SetBookMarkText("to_ins_name", obj_Bookmarks, obj_App, to.ins_name);
                            SetBookMarkText("to_ins_num", obj_Bookmarks, obj_App, to.ins_d_num);
                            SetBookMarkText("to_ins_sdate", obj_Bookmarks, obj_App, to.ins_d_sdate);
                            SetBookMarkText("to_inn", obj_Bookmarks, obj_App, to. to_inn);
                            SetBookMarkText("to_ogrn", obj_Bookmarks, obj_App, to.to_ogrn);
                        }
                        else
                        {
                        SetOperatorReqvizit(Word, obj_Bookmarks, obj_App, Doc, this.comboBox3.SelectedItem.ToString());
                        }
                    }
                }
                //SetAgentReqvizit(Word, obj_Bookmarks, obj_App, Doc);
                if (this.comboBox5.SelectedItem != null) { SetBookMarkText("manager", obj_Bookmarks, obj_App, this.comboBox5.SelectedItem.ToString()); SetBookMarkText("manager1", obj_Bookmarks, obj_App, this.comboBox5.SelectedItem.ToString()); }
                if (this.textBox136.Text == "")
                {
                    SetBookMarkText("FIO1", obj_Bookmarks, obj_App, this.textBox8.Text);
                }
                else
                {
                    SetBookMarkText("FIO1", obj_Bookmarks, obj_App, this.textBox8.Text + "(номер карты - " + this.textBox136.Text+")");
                }
                string passportStr = "";
                if (checkBox33.Checked == true)
                {
                    passportStr = textBox9.Text + " № " + textBox110.Text + " дата выдачи " + textBox111.Text + " выдан " + textBox112.Text;
                }
                else if (checkBox34.Checked == true)
                {
                    passportStr = textBox117.Text + " № " + textBox116.Text + " дата выдачи " + textBox115.Text + " выдан " + textBox114.Text;
                }
                SetBookMarkText("Pasport", obj_Bookmarks, obj_App, passportStr);
                SetBookMarkText("Adress", obj_Bookmarks, obj_App, this.textBox10.Text);
                SetBookMarkText("Phone", obj_Bookmarks, obj_App, this.maskedTextBox2.Text);
                ManagerInfo manager = GetmanagerInfo(comboBox5.Text);
                SetBookMarkText("meneger_phone", obj_Bookmarks, obj_App, manager.phone);
                string SMS_yes = ""; string Email_yes=""; 
                if (checkBox67.Checked)
                {
                    SMS_yes="Да";
                }
                else
                {
                    SMS_yes="Нет";
                }
                if (textBox217.Text != "")
                {
                    SetBookMarkText("station_phone", obj_Bookmarks, obj_App, textBox11.Text + "(" + textBox118.Text + ")" + textBox217.Text);
                }
                SetBookMarkText("Email", obj_Bookmarks, obj_App, this.textBox12.Text);
                if (checkBox68.Checked)
                {
                   Email_yes="Да";
                }
                else
                {
                    Email_yes = "Нет";
                }
                SetBookMarkText("SMS_yes", obj_Bookmarks, obj_App, SMS_yes);
                SetBookMarkText("Email_yes", obj_Bookmarks, obj_App, Email_yes);
                SetBookMarkText("PredNum", obj_Bookmarks, obj_App, this.label195.Text);
                //price
                SetBookMarkText("RubSum", obj_Bookmarks, obj_App, this.textBox14.Text);
                SetBookMarkText("YESUM", obj_Bookmarks, obj_App, this.textBox13.Text);
                SetBookMarkText("Kurs", obj_Bookmarks, obj_App, this.textBox15.Text);
            //
                SetBookMarkText("AvansRubSum", obj_Bookmarks, obj_App, this.textBox43.Text);
                SetBookMarkText("AvansYESum", obj_Bookmarks, obj_App, this.textBox44.Text);
                SetBookMarkText("Kurs2", obj_Bookmarks, obj_App, this.textBox20.Text);
            //
                SetBookMarkText("DolgRubSum", obj_Bookmarks, obj_App, this.textBox46.Text);
                SetBookMarkText("DolgYESum", obj_Bookmarks, obj_App, this.textBox47.Text);
                SetBookMarkText("Kurs3", obj_Bookmarks, obj_App, this.textBox45.Text);
                //SetBookMarkText("PartSum", obj_Bookmarks, obj_App, this.numericUpDown3.Value.ToString());
                //SetBookMarkText("PayDay", obj_Bookmarks, obj_App, this.comboBox7.Text);
                //SetBookMarkText("PayMonth", obj_Bookmarks, obj_App, this.comboBox8.Text);
                //SetBookMarkText("PayYear", obj_Bookmarks, obj_App, this.numericUpDown4.Value.ToString());
                TableProcess(dataGridView6, obj_Tables, 7);
                SetBookMarkText("DayNow1", obj_Bookmarks, obj_App, this.comboBox1.SelectedItem.ToString());
                SetBookMarkText("MonthNow1", obj_Bookmarks, obj_App, this.comboBox2.SelectedItem.ToString());
                SetBookMarkText("YearNow1", obj_Bookmarks, obj_App, this.numericUpDown1.Value.ToString());
                SetBookMarkText("DayNow2", obj_Bookmarks, obj_App, this.comboBox1.SelectedItem.ToString());
                SetBookMarkText("MonthNow2", obj_Bookmarks, obj_App, this.comboBox2.SelectedItem.ToString());
                SetBookMarkText("YearNow2", obj_Bookmarks, obj_App, this.numericUpDown1.Value.ToString());
                //CultureInfo provider = CultureInfo.InvariantCulture;
                //DateTime d1 = DateTime.ParseExact(this.textBox5.Text,"dd-MM-yyyy", provider);
                //d1.dat
                //DateTime d1 = dogovordateend.Date;
                if ((this.dateTimePicker4.Text!=null)&&(this.dateTimePicker4.Text!=""))
                {
                    //DateTime dogovordateend = DateTime.Parse(this.dateTimePicker4.Text).AddDays(1);
                    //SetBookMarkText("DogovorEndTime", obj_Bookmarks, obj_App, dogovordateend.Date.ToShortDateString());
                    SetBookMarkText("DogovorEndTime", obj_Bookmarks, obj_App, this.dateTimePicker4.Text);
                }
                //DateTime d1 = dogovordateend.Date.ToShortDateString();
                //string ssts = dogovordateend.Date.ToShortDateString();
                
                if (checkBox3.Checked == true)
                {
                    SetBookMarkText("Zagranpasport", obj_Bookmarks, obj_App, ", загранпаспорт");
                }
                SetBookMarkText("DogovorNum", obj_Bookmarks, obj_App, textBox49.Text);
                if (checkBox26.Checked == true)
                {
                    MakeNullPredDogovor(Word, textBox49.Text,comboBox5.Text,comboBox6.Text);
                }
                Param[0] = "true";
                obj_App.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, obj_App, Param);
                //object m=System.Type.Missing;
                Mainarguments.setparam(Doc, Word, obj_App);
            
            
            //make zayvka
            object[] ExcelParam = new object[1];

            Type obj_excel=Type.GetTypeFromProgID("Excel.Application");
            object Excel = Activator.CreateInstance(obj_excel);
            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks",BindingFlags.GetProperty,null,Excel,null);
            ExcelParam[0] = GetTempTemlate("Template","zayavkaNaOlatyTyraNMain.xls");
            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
            ExcelParam[0] = 1;
            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets,ExcelParam);
           /* SetCellData(comboBox6.Text,"D3",obj_worksheet);
            SetCellData(textBox2.Text,"D4",obj_worksheet);
            string date = textBox4.Text + "-" + textBox5.Text;
            SetCellData(date,"D5",obj_worksheet);
            if (dataGridView1.Rows[0].Cells[1].Value != null) { SetCellData(dataGridView1.Rows[0].Cells[1].Value.ToString(), "D6", obj_worksheet); }*/
            if ((comboBox3.Text != "Росинтур") && (comboBox3.Text != "Магазин Путешествий"))
            {
                if (comboBox4.Text == "РосинтурЮг")
                {
                    SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "A1", obj_worksheet);
                }
                if (comboBox4.Text == "Магазин Путешествий")
                {
                    SetCellData("ООО ТК \"МАГАЗИН ПУТЕШЕСТВИЙ\"", "A1", obj_worksheet);
                }
            }
            else
            {
                if (comboBox3.Text == "Росинтур")
                {
                    SetCellData("ООО ТК \"РОСИНТУР\"", "A1", obj_worksheet);
                }
                if (comboBox3.Text == "Магазин Путешествий")
                {
                    SetCellData("ООО ТК \"МАГАЗИН ПУТЕШЕСТВИЙ\"", "A1", obj_worksheet);
                }
            }
            SetCellData(comboBox19.Text, "H1", obj_worksheet);
            SetCellData(textBox16.Text, "D4", obj_worksheet);
            SetCellData(textBox17.Text, "D5", obj_worksheet);
            SetCellData(textBox18.Text, "D6", obj_worksheet);
            SetCellData(textBox19.Text, "D7", obj_worksheet);
            SetCellData("Основной договор № "+textBox49.Text, "D8", obj_worksheet);
            string Discount="";
            for (int i = 0; i < dataGridView7.RowCount; i++)
            {
                if (dataGridView7.Rows[i].Cells[0].Value!=null) { SetCellData(dataGridView7.Rows[i].Cells[0].Value.ToString(), "A" + (12 + i), obj_worksheet); }
                if (dataGridView7.Rows[i].Cells[1].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[1].Value.ToString(), "B" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[2].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[2].Value.ToString(), "C" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[3].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[3].Value.ToString(), "D" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[4].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[4].Value.ToString(), "E" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[5].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[5].Value.ToString(), "F" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[6].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[6].Value.ToString(), "G" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[7].Value!=null) {SetCellData(dataGridView7.Rows[i].Cells[7].Value.ToString(), "H" + (12 + i), obj_worksheet);}
                if (dataGridView7.Rows[i].Cells[8].Value != null) { SetCellData(dataGridView7.Rows[i].Cells[8].Value.ToString(), "I" + (12 + i), obj_worksheet); }
                if (dataGridView7.Rows[i].Cells[9].Value != null) { SetCellData(dataGridView7.Rows[i].Cells[9].Value.ToString(), "J" + (12 + i), obj_worksheet); }
            }

                if (dataGridView8.Rows[0].Cells[0].Value!=null) { SetCellData(dataGridView8.Rows[0].Cells[0].Value.ToString(), "B18", obj_worksheet);}
                if (dataGridView8.Rows[0].Cells[1].Value!=null) {SetCellData(dataGridView8.Rows[0].Cells[1].Value.ToString(), "C18", obj_worksheet);}
                if (dataGridView8.Rows[0].Cells[2].Value != null) { SetCellData(dataGridView8.Rows[0].Cells[2].Value.ToString(), "D18", obj_worksheet); Discount = dataGridView8.Rows[0].Cells[2].Value.ToString(); }
                if (dataGridView8.Rows[0].Cells[3].Value!=null) {SetCellData(dataGridView8.Rows[0].Cells[3].Value.ToString(), "E18", obj_worksheet);}
                if (dataGridView8.Rows[0].Cells[4].Value!=null) {SetCellData(dataGridView8.Rows[0].Cells[4].Value.ToString(), "F18", obj_worksheet);}
                if (dataGridView8.Rows[0].Cells[5].Value!=null) {SetCellData(dataGridView8.Rows[0].Cells[5].Value.ToString(), "G18", obj_worksheet);}
                if (dataGridView8.Rows[0].Cells[6].Value!=null) {SetCellData(dataGridView8.Rows[0].Cells[6].Value.ToString(), "H18", obj_worksheet);}
                if (dataGridView8.Rows[0].Cells[7].Value != null) { SetCellData(dataGridView8.Rows[0].Cells[7].Value.ToString(), "I18", obj_worksheet); }
                if (dataGridView8.Rows[0].Cells[8].Value != null) { SetCellData(dataGridView8.Rows[0].Cells[8].Value.ToString(), "J18", obj_worksheet); }
                SetCellData("№ " + textBox49.Text, "B3", obj_worksheet);
                SetCellData("от " + comboBox1.Text + " " + comboBox2.Text + " " + numericUpDown1.Value.ToString(), "D3", obj_worksheet);
            SetCellData(comboBox5.Text, "B25", obj_worksheet);
            if (radioButton4.Checked == true)
            {
                SetCellData("RUR", "B19", obj_worksheet);
                currency = "RUR";
            }
            else if (radioButton5.Checked == true)
            {
                SetCellData("EUR", "B19", obj_worksheet);
                currency = "EUR";
            }
            else if (radioButton6.Checked == true)
            {
                SetCellData("USD", "B19", obj_worksheet);
                currency = "USD";
            }
            //
            SetCellData(textBox15.Text, "E19", obj_worksheet);
            if ((radioButton5.Checked == true) || (radioButton6.Checked == true))
            {
                if ((dataGridView8.Rows[0].Cells[8].Value != null)&&(textBox15.Text!=""))
                {
                    zRubSum = Convert.ToDouble(textBox15.Text) * Convert.ToDouble(dataGridView8.Rows[0].Cells[8].Value);
                }
            }
            else if (radioButton4.Checked == true)
            {
                if (dataGridView8.Rows[0].Cells[8].Value != null)
                {
                    zRubSum = Convert.ToDouble(dataGridView8.Rows[0].Cells[8].Value);
                }
            }
            //SetCellData(zRubSum.ToString(), "J18", obj_worksheet);
            SetCellData(Convert.ToInt32(zRubSum).ToString(), "J19", obj_worksheet);
            //
            /*if (comboBox3.Text != "Росинтур")
            {
                if (comboBox4.Text == "РосинтурЮг")
                {
                    SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "D24", obj_worksheet);
                }
                if (comboBox4.Text == "Магазин Путешествий")
                {
                    SetCellData("ООО ТК \"МАГАЗИН ПУТЕШЕСТВИЙ\"", "D24", obj_worksheet); ;
                }
                //SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "D24", obj_worksheet);
            }
            else if (comboBox3.Text == "Росинтур")
            {
                SetCellData("ООО ТК \"РОСИНТУР\"", "D24", obj_worksheet);
            }*/
            SetCellData(textBox44.Text, "C20", obj_worksheet);
            SetCellData(textBox43.Text, "J20", obj_worksheet);
            SetCellData(textBox20.Text, "F20", obj_worksheet);
            SetCellData(textBox47.Text, "C21", obj_worksheet);
            SetCellData(textBox46.Text, "J21", obj_worksheet);
            SetCellData(textBox45.Text, "F21", obj_worksheet);
            //make tyrpytevka
            if (comboBox3.Text == "Росинтур")
            {
                SetCellData("061300", "I31", obj_worksheet);
                SetCellData("061300", "B39", obj_worksheet);
            }
            else
            {
                SetCellData("061400", "I31", obj_worksheet);
                SetCellData("061400", "B39", obj_worksheet);
            }
            SetCellData(comboBox6.Text, "B33", obj_worksheet);
            SetCellData(comboBox6.Text, "B37", obj_worksheet);
            SetCellData(textBox49.Text, "E32", obj_worksheet);
            SetCellData(textBox49.Text, "C55", obj_worksheet);
            SetCellData(textBox6.Text, "D39", obj_worksheet);
            SetCellData(dateTimePicker3.Text + "-" + dateTimePicker4.Text, "C41", obj_worksheet);
            SetCellData(comboBox1.Text + " " + comboBox2.Text + " " + numericUpDown1.Value.ToString(), "G51", obj_worksheet);
            SetCellData(comboBox1.Text + " " + comboBox2.Text + " " + numericUpDown1.Value.ToString(), "B54", obj_worksheet);
            SetCellData(comboBox1.Text + " " + comboBox2.Text + " " + numericUpDown1.Value.ToString(), "I55", obj_worksheet);
            SetCellData(comboBox5.Text, "A51", obj_worksheet);
            string data1 = "";
            if (comboBox3.Text == "Росинтур")
            {
                //data1 = "ИНН 6164076708; ОКПО ; ООО ТК «Росинтур»; адрес: 344007, г.Ростов-на-Дону, ул.Пушкинская 104/32, 2 эт.";
                data1 = "Общество с ограниченной ответственностью туристическая компания \"Росинтур\" ООО ТК \"Росинтур\", адрес: 344006, г. Ростов-на-Дону, ул.Пушкинская, 104/32, 2 этаж.  344006, г. Ростов-на-Дону, ул.Пушкинская, 104/32, 2 этаж тел.244-22-42; 299-94-30; 299-98-29; 269-42-48; 269-42-49; ИНН 6164076708  ОКПО 49804626  МВТ №001663";
            }
            if ((comboBox3.Text != "Росинтур") && (comboBox4.Text == "РосинтурЮг"))
            {
                //data1 = "ИНН 6164209066; ОКПО ; ООО ТК «Росинтур-Юг»; адрес: 344007, г.Ростов-на-Дону, ул.Пушкинская 104/32, 2 эт.";
                data1 = "Общество с ограниченной ответственностью туристическая компания \"Росинтур-Юг\" ООО ТК \"Росинтур-Юг\", адрес: 344007, г.Ростов-на-Дону, ул.Пушкинская 104/32, 2 эт.  344006, г. Ростов-на-Дону, ул.Пушкинская, 104/32, 2 этаж тел. 244-13-78; 244-22-42,299-94-30,269-42-48; ИНН 6164209066  ОКПО 79215599";
            }
            if ((comboBox3.Text == "Магазин Путешествий") || (comboBox4.Text == "Магазин Путешествий"))
            {
                //data1 = "ИНН " + textBox4.Text + "; ОКПО " + textBox23.Text + "; " + textBox5.Text + "; адрес: " + textBox22.Text;
                data1 = "Общество с ограниченной ответственностью туристическая компания \"Магазин путешествий\" ООО ТК \"Магазин путешествий\", адрес: 344007, г.Ростов-на-Дону, ул.Пушкинская 104/32, 2 эт.  344006, г. Ростов-на-Дону, ул.Пушкинская, 104/32, 2 этаж тел. 244-13-78; 244-22-42,299-94-30,269-42-48; ИНН 6164090646  ОКПО 55510521";
            }
            SetCellData(data1, "A29", obj_worksheet);
            //string data1 = "ИНН " + textBox4.Text + "; ОКПО " + textBox23.Text + "; " + textBox5.Text + "; адрес: " + textBox22.Text;
            string orderdata = textBox9.Text + " адрес :" + textBox10.Text;
            SetCellData(orderdata, "A35", obj_worksheet);
            int turputSum = 0;
            try
            {
                if (checkBox26.Checked == true)
                {
                    if ((textBox14.Text != "") && (textBox14.Text != null))
                    {
                        turputSum = Convert.ToInt32(textBox14.Text);
                    }
                }
                else
                {
                    turputSum = Convert.ToInt32(textBox43.Text) + Convert.ToInt32(textBox46.Text);
                }
            }
            catch
            {

            }
            /*if ((textBox14.Text != "") && (textBox14.Text != null))
            {*/
                TranslateData t1 = new TranslateData();
                string trstring = t1.TranslateStr(turputSum.ToString());
                SetCellData(trstring, "B43", obj_worksheet);
                SetCellData(trstring, "B45", obj_worksheet);
           // }
            
            //make tyrpytevka
                if (checkBox72.Checked == true)
                {
                    save_param += "(РБ)";
                }
                DocumentsaveA(Doc, obj_workbook, textBox49.Text, comboBox5.Text, comboBox6.Text, save_param);
            ExcelParam[0]="True";
            Excel.GetType().InvokeMember("Visible",BindingFlags.SetProperty,null,Excel,ExcelParam);
            Mainarguments.setparamE(Excel, obj_workbook);

            //DatasaveSQL
            object id = "";
            try
            {
                Client ClientData = new Client(textBox8.Text, textBox117.Text, textBox116.Text, textBox115.Text, textBox114.Text, textBox9.Text, textBox110.Text, textBox111.Text, textBox112.Text, makeSQLdate(maskedTextBox1.Text,'.'), textBox12.Text, maskedTextBox2.Text, "", textBox10.Text,"",textBox11.Text+"("+textBox118.Text+")"+textBox217.Text);
                id = GetClientId(ClientData);
                if (id==null)
                {
                    id=ClientInsert(ClientData);
                }
                else
                {
                    ClientUpdate(ClientData, id.ToString());
                }
            }
            catch (Exception e1)
            {
                this.richTextBox1.AppendText("Ошибка при получении клиента в основном договоре \n\r");
                erorrFSave("error.txt", e1.ToString());
            }
           // try
            //{
                string dID = "";
                if (id == null)
                {
                    id="";
                }
                DogovorInfo dinfo = new DogovorInfo(textBox49.Text, DateTime.Today.ToShortDateString(), textBox2.Text, textBox6.Text, dateTimePicker3.Text, dateTimePicker4.Text, textBox19.Text, comboBox19.Text, currency, textBox15.Text, textBox14.Text, textBox13.Text, "Основной", comboBox5.Text, comboBox3.Text, id.ToString(), comboBox28.Text, Discount, SMS_yes, Email_yes,textBox136.Text);
  
                try
                {
                    dID = DogovorInfoSave(dinfo, dataGridView14, dataGridView1, dataGridView2, dataGridView3, dataGridView4, dataGridView5, dataGridView6, dataGridView7);
                }
                catch (Exception e2)
                {
                    this.richTextBox1.AppendText("Ошибка при сохранении основного договора \n\r");
                    erorrFSave("error.txt", e2.ToString());
                }
                //reklama
                try
                {
                    SqlConnectionStringBuilder connectStr = GetConnectSTR();
                    SqlDataReader reader; short excist = 0;
                    SqlCommand sqlcom = null;
                    SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                    string query = ""; string Queskey = "";
                    connect.Open();
                    {
                        query = "select count(DInfoKey) as c from reklamaanswers where DInfoKey='"+dID +"'";
                        sqlcom = new SqlCommand(query, connect);
                        reader = sqlcom.ExecuteReader();
                        if (reader.HasRows != false)
                        {
                            while (reader.Read())
                            {
                                if (reader["c"].ToString() != "0")
                                {
                                    excist = 1;
                                }
                            }
                        }
                        reader.Close();
                        if (excist != 1)
                        {
                            //first Question
                            Queskey = GetQuestionId(groupBox15.Text);
                            if (checkBox11.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox11.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox37.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox37.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox38.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox38.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox39.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox39.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox40.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox40.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox41.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox41.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox42.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox42.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox43.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox43.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox44.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox44.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox45.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox45.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox46.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox46.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox47.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox47.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox48.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox48.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox49.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox49.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox50.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox50.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox35.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox35.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox36.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox36.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            //first Question
                            //second Question
                            Queskey = GetQuestionId(groupBox16.Text);
                            if (checkBox51.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox51.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox52.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox52.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox53.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox53.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox54.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox54.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox55.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox55.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox56.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox56.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox57.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox57.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox58.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + checkBox58.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            if (checkBox62.Checked == true)
                            {
                                query = "insert into ReklamaAnswers values('" + dID + "','" + Queskey + "','" + textBox169.Text + "')";
                                sqlcom = new SqlCommand(query, connect);
                                sqlcom.ExecuteNonQuery();
                            }
                            //second Question
                        }
                    }
                    connect.Close();

                }
                catch
                {
                    richTextBox1.AppendText("Ошибка при добавлении информации по рекламе");
                }

                //reklama
                for (int id4 = 0; id4 < dataGridView4.ColumnCount; id4++)
                {
                    if (dataGridView4.Rows[0].Cells[id4].Value == null)
                    {
                        dataGridView4.Rows[0].Cells[id4].Value = "";
                    }
                }
                FlightInfo finfo = new FlightInfo(dID, dataGridView4.Rows[0].Cells[0].Value.ToString(), dataGridView4.Rows[0].Cells[1].Value.ToString(), dataGridView4.Rows[0].Cells[2].Value.ToString(), dataGridView4.Rows[0].Cells[3].Value.ToString(), dataGridView4.Rows[0].Cells[4].Value.ToString(), dataGridView4.Rows[0].Cells[5].Value.ToString(),dataGridView4.Rows[0].Cells[6].Value.ToString(), textBox19.Text, textBox8.Text, id.ToString());
                try
                {   
                    FlightInfoSave(finfo, dinfo.Manager);
                }
                catch
                {
                    this.richTextBox1.AppendText("Ошибка при получении сохранении полетных данных \n\r");
                }
            //DatasaveSQLEnd
            //number+
            try
            {
                if ((checkBox26.Checked == true) && (textBox49.Text != ""))
                {
                    if ((comboBox3.Text == "Росинтур") || (comboBox3.Text == "Магазин Путешествий"))
                    {
                        IncInINum(comboBox3.Text, textBox49.Text, "ClientDocCount");
                    }
                    else
                    {
                        IncInINum(comboBox4.Text, textBox49.Text, "ClientDocCount");
                    }
                    //IncInINum(comboBox16.Text, textBox7.Text);
                }
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка увеличения номера договора в основном договоре \n\r");
            }
            //clean W
            Marshal.ReleaseComObject(obj_Tables);
            // Marshal.ReleaseComObject(obj_Selection);
            //Marshal.ReleaseComObject(obj_Range);
            Marshal.ReleaseComObject(obj_Doc);
            Marshal.ReleaseComObject(obj_Bookmarks);
            //Marshal.ReleaseComObject(obj_Bookmark);
            Marshal.ReleaseComObject(obj_App);
            // Marshal.ReleaseComObject(Word);
            // GC.GetTotalMemory(true);
            //clean Ex
                     
            Marshal.ReleaseComObject(obj_workbooks);
            Marshal.ReleaseComObject(obj_worksheet);
            Marshal.ReleaseComObject(obj_workbook);
            Marshal.ReleaseComObject(obj_worksheets);
            //make zayvka
            button3.Enabled = true;
            button4.Enabled = true;
            button2.Enabled = true;
            button14.Visible = false;
        }
            
        private string GetQuestionId(string question)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;  short find = 0;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            string QuestionKey="";
            string  Quesquery="select id, Question from ReklamaQuestions where Question='"+question+"'";
                            
            connect.Open();
            {
                sqlcom = new SqlCommand(Quesquery, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        QuestionKey=reader["id"].ToString();
                        find = 1;
                    }
                   
                }
                reader.Close();
                if (find != 1)
                {
                    //SqlDataReader newreader;
                    sqlcom=new SqlCommand("insert into ReklamaQuestions values('"+question+"')",connect);
                    sqlcom.ExecuteNonQuery();
                    Quesquery="select id, Question from ReklamaQuestions where Question='"+question+"'";
                    sqlcom=new SqlCommand(Quesquery,connect);
                    reader=sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            QuestionKey = reader["id"].ToString();
                        }

                    }
                    reader.Close();
                }
               
            }
            connect.Close();
            return QuestionKey;
        }
        private void SetCellData(string str, string cell, object worksheet)
        {
            object[] param = new object[1];
            param[0] = cell;
            object range = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, param);
            param[0] = str;
            range.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, range, param);
        }
        private void TableProcess(DataGridView data, object Tables, int num)
        {
            string tempstr = ""; string s = null; int pos = -1;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                for (int j = 0; j < data.Rows[i].Cells.Count; j++)
                {
                    if (data.Rows[i].Cells[j].Value != null)
                    {
                        if (((data.Name == dataGridView10.Name) || (data.Name == dataGridView4.Name)) && (j == 6))
                        {
                            s = data.Rows[i].Cells[j].Value.ToString();

                            pos = s.IndexOf('*');
                            if (pos != -1)
                            {
                                tempstr = s.Substring(0, pos) + " а/б  оплачивает по ВПД ООО АПЛ \"Градиент\"";
                            }
                            else
                            {
                                tempstr = data.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                        else
                        {
                            tempstr = data.Rows[i].Cells[j].Value.ToString();
                        }
                        SetTableItemText(Tables, num, i + 2, j + 1, tempstr);
                    }
                }
            }
        }
        private void TableProcessCheck(DataGridView data, object Tables, int num)
        {
            for (int i = 0; i < data.Rows.Count-1; i++)
            {
                for (int j = 0; j < data.Rows[i].Cells.Count; j++)
                {
                    if (data.Rows[i].Cells[j].Value != null)
                    {
                        if ((bool)data.Rows[i].Cells[j].Value == true)
                        {
                            SetTableItemText(Tables, num, i + 2, j + 1, "Да");
                        }
                        else
                        {
                            SetTableItemText(Tables, num, i + 2, j + 1, "Отказался");
                        }
                    }
                    else if (data.Rows[i].Cells[j].Value==null)
                    {
                        SetTableItemText(Tables, num, i + 2, j + 1, "Отказался");
                    }
                }
            }
        }
        private void SetOperatorReqvizit(object W, object obj_Bookmarks,object obj_App, object obj_Doc, string tyroperator)
        {
            object[] Parametr = new object[1];
           /* Type obj_Class = Type.GetTypeFromProgID("Word.Application");
            object W = Activator.CreateInstance(obj_Class);*/

            object AppWord = W.GetType().InvokeMember("Application", BindingFlags.GetProperty,null,W,null);
            //AppWord.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, AppWord, Parametr);
            object Docs = AppWord.GetType().InvokeMember("Documents", BindingFlags.GetProperty,null,AppWord,null);
            Parametr[0] = Basepath + @"Template\TourOperatorRekvizits.doc";
            object DocWord = Docs.GetType().InvokeMember("Open",BindingFlags.InvokeMethod,null,Docs, Parametr);
            //Parametr[0] = "false";
            object Bookmarks = DocWord.GetType().InvokeMember("Bookmarks",BindingFlags.GetProperty,null,DocWord,null);
            Parametr[0] = tyroperator;
            object Bookmark =Bookmarks.GetType().InvokeMember("Item",BindingFlags.InvokeMethod,null, Bookmarks,Parametr);
            Bookmark.GetType().InvokeMember("Select",BindingFlags.InvokeMethod,null,Bookmark,null);
            object Selection = AppWord.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, AppWord, null);
            Selection.GetType().InvokeMember("Copy",BindingFlags.InvokeMethod,null,Selection,null);
            Parametr[0] = "OperatorRequizit";
            DocWord.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, DocWord, null);
            obj_Doc.GetType().InvokeMember("Activate",BindingFlags.InvokeMethod,null,obj_Doc,null);
            object obj_Bookmark = obj_Bookmarks.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, obj_Bookmarks, Parametr);
            obj_Bookmark.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, obj_Bookmark, null);
            object Selection1 = obj_App.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, obj_App, null);
            Selection1.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, Selection1, null);


        }
        private void MakeNullPredDogovor(object W, string name, string managername,string client)
        {
            Section sec = new Section();
           // string Path
            object[] Parametr = new object[1];
            object App = W.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, W, null);
            object Docs = App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, App, null);
            Parametr[0] = Basepath + @"Template\NullDogovor.doc";
            object Doc = Docs.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, Docs, Parametr);
            //path
            string path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini"));
            try
            {
                if (path != null)
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
            }
            catch
            {
                path = null;
            }
            if ((path != null) && (Directory.Exists(path)))
            {
                if (managername != "")
                {
                    path += "\\" + managername + "\\" + DateTime.Now.Year.ToString();
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                }

                //WordParam[0] = CheckFileName(path + "\\Договор " + num, ".doc");
                Parametr[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Договор " + GetPredFromMain(name) + "(не заключался)", ".doc");

            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\КлиентскиеДоговора"))
                {
                    Directory.CreateDirectory("c:\\КлиентскиеДоговора");
                    localpath = "c:\\КлиентскиеДоговора";
                }
                else
                {
                    localpath = "c:\\КлиентскиеДоговора";
                }

                Parametr[0] = CheckFileName(localpath + "\\" + "(" + client + ")" + "Договор " + GetPredFromMain(name) + "(не заключался)", ".doc");
            }
            //path
            //Parametr[0] = name+"(не заключался)";
            Doc.GetType().InvokeMember("SaveAs",BindingFlags.InvokeMethod,null,Doc,Parametr);
            Doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, Doc, null);
            Marshal.ReleaseComObject(Docs);
            Marshal.ReleaseComObject(App);
            
        }
        private string GetPredFromMain(string text)
        {
            string result = null;
            string[] str = text.Split('-');
            str[2] = "П";
            result = string.Join("-", str);
            return result;
        }
        private void SetBookMarkText(string name, object Bookmarks, object App,string text)
        {
            object[] Parametr;
            Parametr = new object[1];
            Parametr[0]=name;
            object Bookmark = Bookmarks.GetType().InvokeMember("Item",BindingFlags.InvokeMethod,null,Bookmarks,Parametr);
            Bookmark.GetType().InvokeMember("Select",BindingFlags.InvokeMethod,null,Bookmark,null);
            object Range = App.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, App, null);
            Parametr[0]=text;
            Range.GetType().InvokeMember("Text", BindingFlags.SetProperty,null,Range,Parametr);
        }
        private void SetTableItemText(object Tables, int num,int x, int y,string text)
        {
            object[] Parametr,Cells;
            Parametr = new object[1];
            Parametr[0]=num;
            Cells = new object[2] {x,y};
            object Table = Tables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, Tables, Parametr);
            object Cell = Table.GetType().InvokeMember("Cell",BindingFlags.InvokeMethod,null,Table,Cells);
            object Range = Cell.GetType().InvokeMember("Range",BindingFlags.GetProperty,null,Cell,null);
            Parametr[0] = text;
            Range.GetType().InvokeMember("Text", BindingFlags.SetProperty,null,Range,Parametr);

        }
        private void TableSize(DataGridView data, object Tables, int num, int numrows)
        {
            object[] Parametr, Parametr1, Cells;
            Parametr = new object[1];
            Parametr1 = new object[7] { Missing.Value, true, true, true, true, true, true };
            Parametr[0] = num;
            object Table = Tables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, Tables, Parametr);
            object Rows = Table.GetType().InvokeMember("Rows", BindingFlags.GetProperty, null, Table, null);
            //object rt = Rows.GetType().InvokeMember("Index", BindingFlags.GetProperty, null, Rows, null);
            object Row = null;//=Rows.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, Rows, null);
            for (int i = 0; i + numrows < data.RowCount - 1; i++)
            {
                Rows.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, Rows, null);
            }
            Parametr[0] = true;
            object Border1 = Rows.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, Rows, null);
            Border1.GetType().InvokeMember("Enable", BindingFlags.SetProperty, null, Border1, Parametr);
        }
        private void tabPage3_Click(object sender, EventArgs e)
        {
            //this.textBox8.Text = this.textBox1.Text;
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {
            
            //this.dataGridView4.Rows[0].Cells[0].Value = textBox4.Text;
            this.textBox8.Text = this.comboBox6.Text;
            this.textBox16.Text = this.comboBox6.Text;
            this.textBox17.Text = this.textBox2.Text;
            this.textBox18.Text = this.dateTimePicker3.Text+"-"+this.dateTimePicker4.Text;
            if (this.dataGridView14.Rows[0].Cells[1].Value != null)
            {
               this.textBox19.Text = this.dataGridView14.Rows[0].Cells[1].Value.ToString();
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
           // Section sec = new Section();
            object[] WordParam = new object[1];
            object[] ExcelParam = new object[1];
            object Doc = Mainarguments.Doc;
            object WordApp = Mainarguments.App;
            object Workbook = Mainarguments.WorkBook;
            /*string path =Path.GetFullPath(sec.readkey("SavePath", "Path", "app.ini"));
            if ((path != null)&&(Directory.Exists(path)))
            {
                ExcelParam[0] = CheckFileName(path + "\\Заявка " + textBox1.Text, ".xls");

                    //object Doc = arguments.Doc;
                    //object WordApp = arguments.App;
                    WordParam[0] = CheckFileName(path + "\\Договор №" + textBox1.Text, ".doc");
                    Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, WordParam);

                Workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Workbook, ExcelParam);
            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\КлиентскиеДоговора"))
                {
                    Directory.CreateDirectory("c:\\КлиентскиеДоговора");
                    localpath = "c:\\КлиентскиеДоговора";
                }
                else
                {
                     localpath= "c:\\КлиентскиеДоговора";
                }

                    //object Doc = arguments.Doc;
                    //object WordApp = arguments.App;
                WordParam[0] = localpath + "\\Договор " + textBox1.Text;
                    Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, WordParam);

                    ExcelParam[0] = localpath + "\\Заявка " + textBox1.Text;
                Workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Workbook, ExcelParam);

            }*/
            try
            {
                DocumentSave(Doc, Workbook);
                Doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, Doc, null);
                WordApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, WordApp, null);
                Workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, Workbook, null);
            }
            catch
            {

            }
            button4.Enabled = false;
            button3.Enabled = false;
            Marshal.ReleaseComObject(Mainarguments.Doc);
            Marshal.ReleaseComObject(Mainarguments.Word);
            Marshal.ReleaseComObject(Mainarguments.WorkBook);
            Marshal.ReleaseComObject(Mainarguments.Excel);
            Mainarguments.clean();
            GC.GetTotalMemory(true);
            /*if ((checkBox26.Checked == true) && (textBox49.Text != ""))
            {
                if ((comboBox3.Text == "Росинтур") || (comboBox3.Text == "Магазин Путешествий"))
                {
                    IncInINum(comboBox3.Text, textBox49.Text);
                }
                else
                {
                    IncInINum(comboBox4.Text, textBox49.Text);
                }
                //IncInINum(comboBox16.Text, textBox7.Text);
            }*/
            if (checkBox10.Checked == false)
            {
                FormResetM();
            }
            checkBox10.Checked = false;
        }
        private string CheckFileName(string str, string ext)
        {
            if (File.Exists(str+ext))
            {
                for (int i = 0; i < 99; i++)
                {
                    if (!File.Exists(str + "(" + i + ")"+ext))
                    {
                        str = str + "(" + i + ")";
                        break;
                    }
                }
            }
            return str;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            /*object[] printarg =new object[9]; 
            if (comboBox3.Text != "Росинтур")
            {
                printarg = new object[9] { Missing.Value, Missing.Value, 4, Missing.Value, Missing.Value, Missing.Value, Missing.Value, numericUpDown2.Value.ToString(), "1,4,8" };
            }
            else if (comboBox3.Text == "Росинтур")
            {
                printarg = new object[9] { Missing.Value, Missing.Value, 4, Missing.Value, Missing.Value, Missing.Value, Missing.Value, numericUpDown2.Value.ToString(), "1,3,6" };
            }
            if (checkBox11.Checked == false)
            {
                object Doc = arguments.Doc;
                Doc.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, Doc, printarg);
            }*/
            object Doc = Mainarguments.Doc;
            Doc.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, Doc, null);
            object Work = Mainarguments.WorkBook;
            Work.GetType().InvokeMember("PrintOut",BindingFlags.InvokeMethod,null,Work,null);
        }

        private void comboBox6_TextChanged(object sender, EventArgs e)
        {

        }
        private Dictionary<string, string> Getclients(string str)
        {
            Dictionary<string, string> result=new Dictionary<string,string>() ;//= new object[0];
            Section sec = new Section();
            if (File.Exists("app.ini"))
            {
                if (str != "")
                {
                    SqlConnectionStringBuilder connectStr = new SqlConnectionStringBuilder();
                    connectStr.DataSource = sec.readkey("SQL", "Server", "app.ini");
                    connectStr.UserID = sec.readkey("SQL", "User_ID", "app.ini");
                    connectStr.Password = sec.readkey("SQL", "Pass", "app.ini");
                    connectStr.InitialCatalog = sec.readkey("SQL", "DataBase", "app.ini");
                    string query = @"SELECT FIO,id FROM dbo.Clients where FIO like '" + str + "%' ORDER BY FIO ASC"; //,Birthday,ENpassportnumber,ENpasportStartDate,ENpasportEndDate,phone,email,Adress FROM dbo.Clients_view";

                    SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                    connect.Open();
                    if (connect.State == ConnectionState.Open)
                    {
                        SqlCommand sqlcom = new SqlCommand(query, connect);
                        SqlDataReader read = sqlcom.ExecuteReader();
                        //result = new object[read.FieldCount];

                        while (read.Read())
                        {
                           // for (int i = 0; i < result.Length; i++)
                            //{
                                result.Add(read["id"].ToString(),read["FIO"].ToString());//, = read[i];
                            //}
                        }
                        read.Close();
                        connect.Close();
                    }

                }
                clientsSerarch.Clear();
            }
            return result;
        }

        /*private object[] GetClientData(string str)
        {
            object[] result = new object[0];
            Section sec = new Section();
            if (File.Exists("app.ini"))
            {

                SqlConnectionStringBuilder connectStr = new SqlConnectionStringBuilder();
                connectStr.DataSource = sec.readkey("SQL", "Server", "app.ini");
                connectStr.UserID = sec.readkey("SQL", "User_ID", "app.ini");
                connectStr.Password = sec.readkey("SQL", "Pass", "app.ini");
                connectStr.InitialCatalog = sec.readkey("SQL", "DataBase", "app.ini");
                string query = @"SELECT ENpassportseriy,ENpassportnum,ENpassportStartDate,ENpassportOwn,RUPassportseriy,RUPassportNum,RUPassportStartDate,RUPassportOwn, Birthdate,  Adress , phone, email FROM Clients WHERE id= '" + str + "'"; //,Birthday,ENpassportnumber,ENpasportStartDate,ENpasportEndDate,phone,email,Adress FROM dbo.Clients_view";

                SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    SqlCommand sqlcom = new SqlCommand(query, connect);
                    SqlDataReader read = sqlcom.ExecuteReader();
                    result = new object[read.FieldCount];
                    
                    while (read.Read())
                    {
                        for (int i = 0; i < result.Length; i++)
                        {
                            result[i] = read[i];
                        }
                      
                    }
                    read.Close();
                    connect.Close();
                }

            }
            return result;
        }*/
        private void GetClientsData()
        {
            //object[] result = new object[0];
            int first=0;
            string query = "select FIO, birthdate, ENpassportseriy,ENpassportnum,RUPassportseriy,RUPassportNum, Adress, RUPassportOwn, ENpassportOwn, phone, email,ENpassportStartDate,RUPassportStartDate,skype,icq,id from Clients";
            if ((textBox130.Text != "") || (textBox131.Text != "") || (textBox132.Text != "") || (textBox133.Text != "") || (textBox134.Text != "") || (textBox135.Text != ""))
            {
                query += " where ";
                if (textBox130.Text != "")
                {
                    if (first == 0)
                    {
                        query += " FIO like'" + textBox130.Text + "%'";
                        first = 1;
                    }
                    else
                    {
                        query += " and FIO like'" + textBox130.Text + "%'";
                    }
                }
                if (textBox131.Text != "")
                {
                    if (first == 0)
                    {
                        query += " Birthdate like'" + textBox131.Text + "%'";
                        first = 1;
                    }
                    else
                    {
                        query += " and Birthdate like'" + textBox131.Text + "%'";
                    }
                }
                if (textBox132.Text != "")
                {
                    if (first == 0)
                    {
                        query += " ENpassportseriy like'" + textBox132.Text + "%'";
                        first = 1;
                    }
                    else
                    {
                        query += " and ENpassportseriy like'" + textBox132.Text + "%'";
                    }
                }
                if (textBox133.Text != "")
                {
                    if (first == 0)
                    {
                        query += " ENpassportnum like'" + textBox133.Text + "%'";
                        first = 1;
                    }
                    else
                    {
                        query += " and ENpassportnum like'" + textBox133.Text + "%'";
                    }
                }
                if (textBox134.Text != "")
                {
                    if (first == 0)
                    {
                        query += " RUPassportseriy like'" + textBox134.Text + "%'";
                        first = 1;
                    }
                    else
                    {
                        query += " and RUPassportseriy like'" + textBox134.Text + "%'";
                    }
                }
                if (textBox135.Text != "")
                {
                    if (first == 0)
                    {
                        query += " RUPassportNum like'" + textBox135.Text + "%'";
                        first = 1;
                    }
                    else
                    {
                        query += " and RUPassportNum like'" + textBox135.Text + "%'";
                    }
                }
            }
            query += " ORDER BY FIO ASC";
            SqlConnectionStringBuilder connectstr = GetConnectSTR();
            SqlCommand sqlcom = null; SqlDataReader reader; ;
            SqlConnection connect = new SqlConnection(connectstr.ConnectionString);
            connect.Open();
            if (connect.State==ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    dataGridView23.Rows.Clear();
                   //for (int i = 0; i < dataGridView23.RowCount; i++)
                    //{int 
                    int count = 0;
                    while (reader.Read())
                    {
                        
                            dataGridView23.Rows.Add();
                            if (reader["FIO"] != null) { dataGridView23.Rows[count].Cells[0].Value = reader["FIO"].ToString(); }
                            if (reader["birthdate"] != null) { dataGridView23.Rows[count].Cells[1].Value = reader["birthdate"].ToString(); }
                            if (reader["ENpassportseriy"] != null) { dataGridView23.Rows[count].Cells[2].Value = reader["ENpassportseriy"].ToString(); }
                            if (reader["ENpassportnum"] != null) { dataGridView23.Rows[count].Cells[3].Value =  reader["ENpassportnum"].ToString(); }
                            if (reader["RUPassportseriy"] != null) { dataGridView23.Rows[count].Cells[4].Value = reader["RUPassportseriy"].ToString(); }
                            if (reader["RUPassportNum"] != null) { dataGridView23.Rows[count].Cells[5].Value =  reader["RUPassportNum"].ToString(); }
                            if (reader["Adress"] != null) { dataGridView23.Rows[count].Cells[6].Value = reader["Adress"].ToString(); }
                            if (reader["RUPassportOwn"] != null) { dataGridView23.Rows[count].Cells[7].Value = reader["RUPassportOwn"].ToString(); }
                            if (reader["ENpassportOwn"] != null) { dataGridView23.Rows[count].Cells[8].Value = reader["ENpassportOwn"].ToString(); }
                            if (reader["phone"] != null) { dataGridView23.Rows[count].Cells[9].Value = reader["phone"].ToString(); }
                            if (reader["email"] != null) { dataGridView23.Rows[count].Cells[10].Value = reader["email"].ToString(); }
                            if (reader["ENpassportStartDate"] != null) { dataGridView23.Rows[count].Cells[11].Value = reader["ENpassportStartDate"].ToString(); }
                            if (reader["RUPassportStartDate"] != null) { dataGridView23.Rows[count].Cells[12].Value = reader["RUPassportStartDate"].ToString(); }
                            if (reader["skype"] != null) { dataGridView23.Rows[count].Cells[13].Value = reader["skype"].ToString(); }
                            if (reader["icq"] != null) { dataGridView23.Rows[count].Cells[14].Value = reader["icq"].ToString(); }
                            if (reader["id"] != null) { dataGridView23.Rows[count].Cells[15].Value = reader["id"].ToString(); }
                            count++;
                     }
                   // }
                    
                }
            }
            dataGridView23.Rows[0].Selected = true;
        }
        private void comboBox6_DropDown(object sender, EventArgs e)
        {
            /*object[] str;
            //this.comboBox6.Items.Clear();
            //  string strline = this.comboBox6.SelectedItem.ToString();
            str = Getclients(this.comboBox6.Text.ToString());
            if (str.Length != 0)
            {
                for (int i = 0; i < str.Length; i++)
                {
                    if (str[i] != null)
                    {
                        if (!this.comboBox6.Items.Contains(str[i]))
                        {
                            this.comboBox6.Items.Add(str[i]);
                        }
                    }
                }
                //this.comboBox6.Items.AddRange(str);
            }*/
            //string prevstr = this.comboBox6.Text;
            //this.comboBox6.AllowSelection = false;
            //this.comboBox6.SelectedIndex = -1;
            //this.comboBox6.SelectedText = prevstr; 
            Dictionary<string,string> str=null;
            ComboBox com = (ComboBox)sender;
            //this.comboBox6.Items.Clear();
            //  string strline = this.comboBox6.SelectedItem.ToString();
            str = Getclients(com.Text);
            if (str.Count != 0)
            {
                foreach (KeyValuePair<string, string> kvp in str)
                {
                    /*if (str[i] != null)
                    {
                        if (!com.Items.Contains(str[i]))
                        {*/
                   // com.Items.Add(kvp.Key);
                   // clientsSerarch.Add(kvp.Value, (com.Items.Count - 1).ToString());

                    /*     }
                     }*/
                }
            }
        }

        private void comboBox6_SelectionChangeCommitted(object sender, EventArgs e)
        {
           /* object[] str; ComboBox com = (ComboBox)sender;
            //this.comboBox6.Items.Clear();this.combo
            //  string strline = this.comboBox6.SelectedItem.ToString();
            //str = GetClientData(this.comboBox6.Text.ToString());
            //KeyValuePair<string,string>=clientsSerarch[
            //str = GetClientData(clientsSerarch[com.SelectedIndex.ToString()]);
            textBox9.Text = str[0].ToString();
            textBox110.Text = str[1].ToString();
            textBox111.Text = str[2].ToString();
            textBox112.Text = str[3].ToString();
            textBox117.Text = str[4].ToString();
            textBox116.Text = str[5].ToString();
            textBox115.Text = str[6].ToString();
            textBox114.Text = str[7].ToString();
            /*dataGridView6.Rows[0].Cells[0].Value = this.comboBox6.Text;
            dataGridView6.Rows[0].Cells[1].Value = str[4];
            dataGridView6.Rows[0].Cells[2].Value = str[0];
            dataGridView6.Rows[0].Cells[3].Value = str[5];
            dataGridView6.Rows[0].Cells[4].Value = str[6];*/
            /*textBox118.Text = str[8].ToString();
            textBox10.Text = str[9].ToString();
            textBox11.Text = str[10].ToString();
            textBox12.Text = str[11].ToString();*/

        }
        private void FormResetM()
        {
            //1
            textBox2.Text = "";
            textBox3.Text = "";
            //textBox4.Text = "";
            //textBox5.Text = "";
            textBox6.Text = "";
            comboBox6.Text = "";
            comboBox5.Text = "";
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            DataGridReset(dataGridView14);
            //2
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            DataGridReset(dataGridView1);
            DataGridReset(dataGridView2);
            DataGridReset(dataGridView3);
            DataGridReset(dataGridView4);
            for (int i = 0; i < dataGridView5.ColumnCount; i++)
            {
                dataGridView5.Rows[0].Cells[i].Value = false;
            }
            //3
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            maskedTextBox2.Text = "";
            textBox12.Text = "";
            textBox49.Text = "";
            textBox48.Text = "";
            textBox1.Text = "";
            checkBox26.Checked = false;

            //numericUpDown3.Value = 0;
            DataGridReset(dataGridView6);
            //4
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox44.Text = "";
            textBox43.Text = "";
            textBox20.Text = "";
            textBox47.Text = "";
            textBox46.Text = "";
            textBox45.Text = "";

            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            DataGridReset(dataGridView8);
            DataGridReset(dataGridView7);
            /*checkBox7.Checked = false;
            checkBox7.Enabled = true;
            checkBox8.Checked = false;
            checkBox8.Enabled = true;
            checkBox9.Checked = false;
            checkBox9.Enabled = true;*/
            //5
            //textBox4.Text = "";
            //textBox5.Text = "";
            //textBox22.Text = "";
            //textBox23.Text = "";

        }
        private void FormResetP()
        {
            //1
            textBox25.Text = "";
            //textBox24.Text = "";
            //textBox23.Text = "";
            //textBox22.Text = "";
            textBox21.Text = "";
            comboBox9.Text = "";
            checkBox14.Checked = false;
            checkBox22.Checked = false;
            //checkBox3.Checked = false;
            //2
            checkBox17.Checked = false;
            checkBox16.Checked = false;
            checkBox15.Checked = false;
            DataGridReset(dataGridView13);
            DataGridReset(dataGridView12);
            DataGridReset(dataGridView11);
            DataGridReset(dataGridView10);
            for (int i = 0; i < dataGridView9.ColumnCount; i++)
            {
                dataGridView9.Rows[0].Cells[i].Value = false;
            }
            //3
            textBox7.Text = "";
            textBox34.Text = "";
            textBox33.Text = "";
            textBox32.Text = "";
            maskedTextBox4.Text = "";
            textBox30.Text = "";
            textBox29.Text = "";
            textBox28.Text = "";
            textBox27.Text = "";
            textBox24.Text = "";
            textBox26.Text = "";
            //numericUpDown6.Value = 0;
            comboBox12.Text = "";
            comboBox13.Text = "";
            comboBox14.Text = "";
            //DataGridReset(dataGridView6);
            //4
            textBox38.Text = "";
            textBox37.Text = "";
            textBox36.Text = "";
            textBox35.Text = "";
            DataGridReset(dataGridView16);
            DataGridReset(dataGridView15);
            /*checkBox20.Checked = false;
            checkBox20.Enabled = true;
            checkBox18.Checked = false;
            checkBox18.Enabled = true;
            checkBox19.Checked = false;
            checkBox19.Enabled = true;*/


        }
        private string GetmangerId(string manager)
        {
            string result=null;
            if (manager == "Буренко М.М")
            {
                result = "01";
            }
            if (manager == "Зелинская Е.И")
            {
                result = "02";
            }
            if (manager == "Бровко Л.Ю")
            {
                result = "03";
            }
            if (manager == "Дулебова Е.В")
            {
                result = "04";
            }
            if (manager == "Елисеева Л.В")
            {
                result = "05";
            }
            if (manager == "Данчук Н.Н")
            {
                result = "06";
            }
            if (manager == "Пономарцева К.Д")
            {
                result = "07";
            }
            if (manager == "Ходокина Е.В")
            {
                result = "08";
            }
            if (manager == "Малий Е.В")
            {
                result = "09";
            }
            if (manager == "Пащинская Т.Е")
            {
                result = "15";
            }
            if (manager == "Дьякова Е.Е")
            {
                result = "16";
            }
            if (manager == "Никонорова К.")
            {
                result = "17";
            }
            return result;
        }
        private void DataGridReset(DataGridView obj)
        {
            for (int i = 0; i < obj.RowCount; i++)
            {
                for (int j = 0; j < obj.ColumnCount; j++)
                {
                    obj.Rows[i].Cells[j].Value = null;
                }
            }
            obj.RowCount = 1;
        }
        private void dataGridView7_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Regex r = new Regex("([0-9]+)");
            Match m; double dscount = 0;
            double sum = 0, sumcol = 0, sumrow = 0,sumdiscount=0 ;
            //if ((e.ColumnIndex == 1) && (e.ColumnIndex == 2) && (e.ColumnIndex != 9))
            if (e.RowIndex > -1)
            {
                if ((e.ColumnIndex != 0) && (e.ColumnIndex != 3) && (e.ColumnIndex != 9))
                {
                    if ((e.ColumnIndex == 2) || ((dataGridView7.Rows[e.RowIndex].Cells[1].Value != null) && (dataGridView7.Rows[e.RowIndex].Cells[2].Value != null)))
                    {
                        m = r.Match(dataGridView7.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                        if (m.ToString() != "")
                        {
                            if ((dataGridView7.Rows[e.RowIndex].Cells[1].Value.ToString() != "") && (dataGridView7.Rows[e.RowIndex].Cells[2].Value.ToString() != ""))
                            {
                                dscount = ((Convert.ToDouble(dataGridView7.Rows[e.RowIndex].Cells[1].Value) / 100) * (Convert.ToDouble(dataGridView7.Rows[e.RowIndex].Cells[2].Value)));
                            }
                            //double f1 = (Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[1].Value) / 100);

                            //double dscount = f1 * (Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[2].Value));
                            //dataGridView7.Rows[e.RowIndex].Cells[3].Value = ((Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[1].Value) / 100) * (Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[2].Value)));

                            if (radioButton4.Checked == true)
                            {
                                dataGridView7.Rows[e.RowIndex].Cells[3].Value = Convert.ToInt32(dscount);
                            }
                            else
                            {
                                dataGridView7.Rows[e.RowIndex].Cells[3].Value = dscount;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Вводите Цифры");
                            dataGridView7.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                            return;
                        }
                    }

                    if (e.ColumnIndex != 2)
                    {
                        // m = null;
                        for (int i = 0; i < dataGridView7.RowCount; i++)
                        {
                            if ((dataGridView7.Rows[i].Cells[e.ColumnIndex].Value != null)&&(dataGridView7.Rows[i].Cells[e.ColumnIndex].Value.ToString() != ""))
                            {
                                m = r.Match(dataGridView7.Rows[i].Cells[e.ColumnIndex].Value.ToString());
                                if (m.ToString() != "")
                                {
                                    try
                                    {
                                        sumrow = sumrow + Convert.ToDouble(dataGridView7.Rows[i].Cells[e.ColumnIndex].Value);
                                    }
                                    catch
                                    {
                                        dataGridView7.Rows[i].Cells[e.ColumnIndex].Value = "";
                                        MessageBox.Show("Вводите Цифры");
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Вводите Цифры");
                                    dataGridView7.Rows[i].Cells[e.ColumnIndex].Value = null;
                                    return;
                                }
                            }
                        }
                        dataGridView8.Rows[0].Cells[e.ColumnIndex - 1].Value = sumrow;
                    }

                }

                for (int i = 1; i < dataGridView7.ColumnCount - 1; i++)
                {
                    if ((i != 2) && (i != 3))
                    {
                        if ((dataGridView7.Rows[e.RowIndex].Cells[i].Value != null)&&(dataGridView7.Rows[e.RowIndex].Cells[i].Value.ToString() !=""))
                        {
                            m = r.Match(dataGridView7.Rows[e.RowIndex].Cells[i].Value.ToString());
                            if (m.ToString() != "")
                            {
                                try
                                {
                                    sumcol = sumcol + Convert.ToDouble(dataGridView7.Rows[e.RowIndex].Cells[i].Value);
                                }
                                catch
                                {
                                    // MessageBox.Show("Вводите Цифры");
                                }
                            }
                        }
                    }
                }
                dataGridView7.Rows[e.RowIndex].Cells[9].Value = (sumcol - Convert.ToDouble(dataGridView7.Rows[e.RowIndex].Cells[3].Value));
                if (radioButton4.Checked == true)
                {
                    dataGridView7.Rows[e.RowIndex].Cells[9].Value = Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[9].Value);
                }
                for (int i = 0; i < dataGridView7.RowCount; i++)
                {
                    //  if (m == r.Match(dataGridView7.Rows[i].Cells[9].Value.ToString()))
                    //{
                    if ((dataGridView7.Rows[i].Cells[9].Value != null) && (dataGridView7.Rows[i].Cells[9].Value.ToString() != ""))
                    {
                        sum = sum + Convert.ToDouble(dataGridView7.Rows[i].Cells[9].Value);
                    }


                    if ((dataGridView7.Rows[i].Cells[3].Value != null)&&(dataGridView7.Rows[i].Cells[3].Value.ToString() !=""))
                    {
                        sumdiscount = sumdiscount + Convert.ToDouble(dataGridView7.Rows[i].Cells[3].Value);
                    }
                    //  }
                }
                dataGridView8.Rows[0].Cells[8].Value = sum;
                dataGridView8.Rows[0].Cells[2].Value = sumdiscount;
                if (radioButton4.Checked == true)
                {
                    if (dataGridView8.Rows[0].Cells[8].Value != null)
                    {
                        textBox13.Text = "";
                        textBox14.Text = dataGridView8.Rows[0].Cells[8].Value.ToString();
                    }
                }
                else if ((radioButton5.Checked == true) || (radioButton6.Checked == true))
                {
                    if (dataGridView8.Rows[0].Cells[8].Value != null)
                    {
                        textBox14.Text = "";
                        textBox13.Text = dataGridView8.Rows[0].Cells[8].Value.ToString();
                    }
                }

            }
        }

       /* private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == false)
            {
                checkBox8.Enabled = true;
                checkBox9.Enabled = true;
            }
            if (checkBox7.Checked == true)
            {
                checkBox8.Enabled = false;
                checkBox9.Enabled = false;
            }
            dataGridView7.Enabled = true;
            dataGridView8.Enabled = true;

        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == false)
            {
                checkBox7.Enabled = true;
                checkBox9.Enabled = true;
            }
            if (checkBox8.Checked == true)
            {
                checkBox7.Enabled = false;
                checkBox9.Enabled = false;
            }
            dataGridView7.Enabled = true;
            dataGridView8.Enabled = true;
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == false)
            {
                checkBox8.Enabled = true;
                checkBox7.Enabled = true;
            }
            if (checkBox9.Checked == true)
            {
                checkBox8.Enabled = false;
                checkBox7.Enabled = false;
            }
            dataGridView7.Enabled = true;
            dataGridView8.Enabled = true;
        }*/

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            if (t.Text != null)
            {
                string str = t.Text;
                if (str.Length == 2)
                {
                    //str+= ".";
                    t.AppendText(".");
                    //textBox4.
                }
                else if (str.Length == 5)
                {
                    t.AppendText(".");
                }
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            /*if (checkBox11.Checked == true)
            {
               tabControl1.TabPages[0].Enabled = false;
               tabControl1.TabPages[1].Enabled = false;
              // tabControl1.TabPages[2].Enabled = false;
                tabControl1.TabPages[3].BringToFront();
                tabControl1.SelectedIndex=3;

           }
            else 
            {
                tabControl1.TabPages[0].Enabled = true;
                tabControl1.TabPages[1].Enabled = true;
                //tabControl1.TabPages[2].Enabled = true;
                tabControl1.SelectedIndex = 0;
            }*/
        }

        private void dataGridView6_Leave(object sender, EventArgs e)
        {
            dataGridView7.RowCount = dataGridView6.RowCount;
            for (int i = 0; i < dataGridView6.RowCount; i++)
            {
                
               // if (dataGridView6.Rows[i].Cells[0]!=null)
               // {
                    dataGridView7.Rows[i].Cells[0].Value = dataGridView6.Rows[i].Cells[0].Value;
               // }
            }
        }

        private void основнойToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox17.Visible = false;
        }

        private void предварительныйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = true;
            groupBox1.Visible = false;
            groupBox3.Visible = false;
            groupBox17.Visible = false;

        }
        private void dataGridView6_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridView7.RowCount = dataGridView6.RowCount;
                for (int i = 0; i < dataGridView6.RowCount; i++)
                {

                    // if (dataGridView6.Rows[i].Cells[0]!=null)
                    // {
                    dataGridView7.Rows[i].Cells[0].Value = dataGridView6.Rows[i].Cells[0].Value;
                    // }
                }
            }
        }
//pred_dogovors
        private void button6_Click(object sender, EventArgs e)
        {
            button13.Visible = true;
            button6.Enabled = false; 
            string currency = ""; 
            string Discount = ""; 
            string dID = "";
            if (dataGridView15.Rows[0].Cells[2].Value != null) 
            {
                Discount = dataGridView15.Rows[0].Cells[2].Value.ToString(); 
            }
            if (radioButton1.Checked == true)
            {
                currency = "RUR";
            }
            else if (radioButton2.Checked == true)
            {
                currency = "EUR";
            }
            else if (radioButton3.Checked == true)
            {
                currency = "USD";
            }
            //DatasaveSQL
            object id = "";

            try
            {
                Client ClientData = new Client(textBox34.Text, textBox122.Text, textBox121.Text, textBox120.Text, textBox119.Text, textBox33.Text, textBox107.Text, textBox108.Text, textBox109.Text, makeSQLdate(maskedTextBox3.Text,'.'), textBox30.Text, maskedTextBox4.Text, "", textBox32.Text,"",textBox218.Text+"("+textBox123.Text+")"+textBox31.Text);
                id = GetClientId(ClientData);
                if (id == null)
                {
                    id = ClientInsert(ClientData);
                }
                else
                {
                    ClientUpdate(ClientData,id.ToString());
                }
            }
            catch
            {
                id = "";
            }
            string SMS_yes = "";
            string Email_yes = "";
            if (checkBox70.Checked)
            {
                SMS_yes="Да";
            }
            else
            {
                SMS_yes="Нет";
            }
            if (checkBox69.Checked)
            {
                Email_yes = "Да";
            }
            else
            {
                Email_yes = "Нет";
            }
            DogovorInfo dinfo = new DogovorInfo(textBox7.Text, DateTime.Now.ToShortDateString(), textBox25.Text, textBox21.Text, dateTimePicker1.Text, dateTimePicker2.Text, textBox35.Text, comboBox17.Text, currency, textBox27.Text, textBox24.Text, textBox26.Text, "Предварительный", comboBox14.Text, comboBox16.Text, id.ToString(), comboBox29.Text, Discount, SMS_yes, Email_yes,"");
            try
            {
                dID = DogovorInfoSave(dinfo, dataGridView17, dataGridView13, dataGridView12, dataGridView11, dataGridView10, dataGridView9, dataGridView18, dataGridView16);
            }
            catch
            {

            }
            //DatasaveSQLEnd
            try
            {
                PredDogSave(dID);
            }
            catch
            {
                //MessageBox.Show("Ошибка сохранения предварительного договора");
                richTextBox1.AppendText("Ошибка сохранения предварительного договора\n\r");
            }
            object obj_App;
            object obj_Doc;
            object obj_Bookmarks;
            //object obj_Bookmark;
            //object obj_Selection;
            //object obj_Range;
            object obj_Tables;

             double zRubSum=0;
            object[] Param;
            string check="";
            string transport="";
            string save_param = "";
            Param = new object[1];
                Type obj_Class = Type.GetTypeFromProgID("Word.Application");
                object Word = Activator.CreateInstance(obj_Class);

                obj_App = Word.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, Word, null);
                obj_Doc = obj_App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj_App, null);
                if ((comboBox16.Text == "Росинтур") || (comboBox16.Text == "Магазин Путешествий"))
                {
                    //Param[0] = Basepath + @"Template\Pred_Rosintour.doc";GetTempTemlate("Template",

                    /*if (comboBox16.Text == "Росинтур")
                    {
                        Param[0] = Basepath + @"Template\Pred_Rosintour.doc"; ;
                    }
                    if (comboBox16.Text == "Магазин Путешествий")
                    {
                        Param[0] = Basepath + @"Template\Pred_TravelMagO.doc";
                    }*/
                    if (comboBox16.Text == "Росинтур")
                    {
                        Param[0] = GetTempTemlate("Template","Pred_Rosintour.doc") ;
                    }
                    if (comboBox16.Text == "Магазин Путешествий")
                    {
                        Param[0] =  GetTempTemlate("Template","Pred_TravelMagO.doc");
                    }
                }
                else// if (comboBox16.Text != "Росинтур")
                {
                    /*if (comboBox15.Text == "РосинтурЮг")
                    {
                        Param[0] = Basepath + @"Template\Pred_RosintourUg.doc";
                    }
                    if (comboBox15.Text == "Магазин Путешествий")
                    {
                        Param[0] = Basepath + @"Template\Pred_TravelMag.doc";
                    }*/
                    if (comboBox15.Text == "РосинтурЮг")
                    {
                        Param[0] = GetTempTemlate("Template","Pred_RosintourUg.doc");
                    }
                    if (comboBox15.Text == "Магазин Путешествий")
                    {
                        Param[0] = GetTempTemlate("Template","Pred_TravelMag.doc");
                    }
                }

                object Doc = obj_Doc.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_Doc, Param);
                obj_Bookmarks = Doc.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, Doc, null);
                SetBookMarkText("DayNow", obj_Bookmarks, obj_App, this.comboBox11.SelectedItem.ToString());
                SetBookMarkText("MonthNow", obj_Bookmarks, obj_App, this.comboBox10.SelectedItem.ToString());
                SetBookMarkText("YearNow", obj_Bookmarks, obj_App, this.numericUpDown5.Value.ToString());
                SetBookMarkText("FIO", obj_Bookmarks, obj_App, this.comboBox9.Text.ToString() + " ");
                /*if (comboBox16.Text != "Росинтур")
                {
                    SetBookMarkText("Tyroperator", obj_Bookmarks, obj_App, this.comboBox16.Text);
                }*/
                //SetBookMarkText("ManNum", obj_Bookmarks, obj_App, this.textBox3.Text);
                SetBookMarkText("TravelProgram", obj_Bookmarks, obj_App, this.textBox25.Text + " ");
                SetBookMarkText("Travelstart", obj_Bookmarks, obj_App, this.dateTimePicker1.Text);
                SetBookMarkText("TravelEnd", obj_Bookmarks, obj_App, this.dateTimePicker2.Text);
                SetBookMarkText("TravelPlace", obj_Bookmarks, obj_App, this.textBox21.Text);
                if (this.checkBox73.Checked!=true)
                {
                    SetBookMarkText("Sp", obj_Bookmarks, obj_App, "");
                }
                if (this.checkBox14.Checked) { check = "Да"; } else { check = "Нет"; }
                SetBookMarkText("checkbox1", obj_Bookmarks, obj_App, check);
                //if (this.checkBox2.Checked) { check = "Да"; } else { check = "Нет"; }
                //SetBookMarkText("checkbox2", obj_Bookmarks, obj_App, check);
                if (this.checkBox12.Checked) { check = "Да"; } else { check = "Нет"; }
                SetBookMarkText("checkbox3", obj_Bookmarks, obj_App, check);
                obj_Tables = Doc.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, Doc, null);
                if (dataGridView17.RowCount > 3)
                {
                    TableSize(dataGridView17, obj_Tables, 1,3);
                }
                TableProcess(dataGridView17, obj_Tables, 1);
                if (dataGridView13.RowCount > 3)
                {
                    TableSize(dataGridView13, obj_Tables, 2,3);
                }
                TableProcess(dataGridView13, obj_Tables, 2);
                if (dataGridView12.RowCount > 3)
                {
                    TableSize(dataGridView12, obj_Tables, 3,3);
                }
                TableProcess(dataGridView12, obj_Tables, 3);
                if (dataGridView11.RowCount > 3)
                {
                    TableSize(dataGridView11, obj_Tables, 4,3);
                }
                TableProcess(dataGridView11, obj_Tables, 4);
                //TableProcess(dataGridView10, obj_Tables, 5);

                if (this.checkBox17.Checked) { transport = "Авиа"; }
                if (this.checkBox16.Checked) { transport = "Ж\\д "; }
                if (this.checkBox15.Checked) { transport = "Авто"; }
                //SetBookMarkText("Transport", obj_Bookmarks, obj_App, waycheck);
                SetBookMarkText("Transport", obj_Bookmarks, obj_App, transport);
                if (dataGridView10.RowCount > 2)
                {
                    TableSize(dataGridView10, obj_Tables, 5, 2);
                }
                TableProcess(dataGridView10, obj_Tables, 5);
                TableProcessCheck(dataGridView9, obj_Tables, 6);
                //SetTableItemText(obj_Tables, 5, 3, 1, "sdfsd");
                //reqvizits
                if ((comboBox16.Text != "Росинтур") && (comboBox16.Text != "Магазин Путешествий"))
                {
                    Touroperator to = new Touroperator();
                    to.getinfo(GetConnectSTR(), this.comboBox16.SelectedItem.ToString());
                    if ((to.to_id != null) && (to.to_id != ""))
                    {
                        SetBookMarkText("to_name", obj_Bookmarks, obj_App, to.to_name);
                        SetBookMarkText("to_reestr_num", obj_Bookmarks, obj_App, to.to_rn);
                        SetBookMarkText("to_adr", obj_Bookmarks, obj_App, to.to_adress);
                        SetBookMarkText("to_tel", obj_Bookmarks, obj_App, to.to_tel);
                        SetBookMarkText("to_fax", obj_Bookmarks, obj_App, to.to_fax);
                        SetBookMarkText("to_fin_cap", obj_Bookmarks, obj_App, to.ins_fin_cap);
                        SetBookMarkText("to_ins_adr", obj_Bookmarks, obj_App, to.ins_adress);
                        SetBookMarkText("to_ins_d_date", obj_Bookmarks, obj_App, to.ins_d_date);
                        SetBookMarkText("to_ins_edate", obj_Bookmarks, obj_App, to.ins_d_edate);
                        SetBookMarkText("to_ins_name", obj_Bookmarks, obj_App, to.ins_name);
                        SetBookMarkText("to_ins_num", obj_Bookmarks, obj_App, to.ins_d_num);
                        SetBookMarkText("to_ins_sdate", obj_Bookmarks, obj_App, to.ins_d_sdate);
                        SetBookMarkText("to_inn", obj_Bookmarks, obj_App, to.to_inn);
                        SetBookMarkText("to_ogrn", obj_Bookmarks, obj_App, to.to_ogrn);
                    }
                    else
                    {
                        SetOperatorReqvizit(Word, obj_Bookmarks, obj_App, Doc, this.comboBox16.SelectedItem.ToString());
                    }
                }
                //SetAgentReqvizit(Word, obj_Bookmarks, obj_App, Doc);
                if (this.comboBox14.SelectedItem != null) { SetBookMarkText("manager", obj_Bookmarks, obj_App, this.comboBox14.SelectedItem.ToString()); SetBookMarkText("manager1", obj_Bookmarks, obj_App, this.comboBox14.SelectedItem.ToString()); }
                SetBookMarkText("FIO1", obj_Bookmarks, obj_App, this.textBox34.Text);
                string passportStr="";
                if (checkBox21.Checked == true)
                {
                    passportStr = textBox33.Text + " № " + textBox107.Text + " дата выдачи " + textBox108.Text + " выдан " + textBox109.Text;
                }
                else if (checkBox32.Checked == true)
                {
                    passportStr = textBox122.Text +" № "+ textBox121.Text +" дата выдачи "+ textBox120.Text +" выдан "+ textBox119.Text;
                }
                SetBookMarkText("Pasport", obj_Bookmarks, obj_App, passportStr);
                SetBookMarkText("Adress", obj_Bookmarks, obj_App, this.textBox32.Text);
                SetBookMarkText("Phone", obj_Bookmarks, obj_App, maskedTextBox4.Text);       
                ManagerInfo manager = GetmanagerInfo(comboBox14.Text);
                SetBookMarkText("meneger_phone", obj_Bookmarks, obj_App, manager.phone);
                SetBookMarkText("SMS_yes", obj_Bookmarks, obj_App, SMS_yes);
                if (textBox218.Text != "")
                {
                    SetBookMarkText("station_phone", obj_Bookmarks, obj_App, textBox218.Text + "(" + textBox123.Text + ")" + textBox31.Text);
                }
                SetBookMarkText("Email", obj_Bookmarks, obj_App, this.textBox30.Text);
                SetBookMarkText("Email_yes", obj_Bookmarks, obj_App, Email_yes);
                //price
                SetBookMarkText("RubSum", obj_Bookmarks, obj_App, this.textBox28.Text);
                SetBookMarkText("YESUM", obj_Bookmarks, obj_App, this.textBox29.Text);
                SetBookMarkText("Kurs2", obj_Bookmarks, obj_App, this.textBox27.Text);
                //full
                SetBookMarkText("AllSumRub", obj_Bookmarks, obj_App, this.textBox24.Text);
                SetBookMarkText("AllSumYE", obj_Bookmarks, obj_App, this.textBox26.Text);
                SetBookMarkText("Kurs1", obj_Bookmarks, obj_App, this.textBox27.Text);
                //SetBookMarkText("PartSum", obj_Bookmarks, obj_App, this.numericUpDown3.Value.ToString());
                SetBookMarkText("PayDay", obj_Bookmarks, obj_App, this.comboBox13.Text);
                SetBookMarkText("PayMonth", obj_Bookmarks, obj_App, this.comboBox12.Text);
                SetBookMarkText("PayYear", obj_Bookmarks, obj_App, this.numericUpDown6.Value.ToString());
                //TableProcess(dataGridView6, obj_Tables, 6);
                SetBookMarkText("DayNow1", obj_Bookmarks, obj_App, this.comboBox11.SelectedItem.ToString());
                SetBookMarkText("MonthNow1", obj_Bookmarks, obj_App, this.comboBox10.SelectedItem.ToString());
                SetBookMarkText("YearNow1", obj_Bookmarks, obj_App, this.numericUpDown5.Value.ToString());
                SetBookMarkText("DayNow2", obj_Bookmarks, obj_App, this.comboBox11.SelectedItem.ToString());
                SetBookMarkText("MonthNow2", obj_Bookmarks, obj_App, this.comboBox10.SelectedItem.ToString());
                SetBookMarkText("YearNow2", obj_Bookmarks, obj_App, this.numericUpDown5.Value.ToString());
                TableProcess(dataGridView18, obj_Tables, 7);
                //CultureInfo provider = CultureInfo.InvariantCulture;
                //DateTime d1 = DateTime.ParseExact(this.textBox5.Text,"dd-MM-yyyy", provider);
                //d1.dat
                //DateTime d1 = dogovordateend.Date;
                //DateTime dogovordateend = DateTime.Parse(this.textBox5.Text).AddDays(1);
                //DateTime d1 = dogovordateend.Date.ToShortDateString();
                //string ssts = dogovordateend.Date.ToShortDateString();
                //SetBookMarkText("DogovorEndTime", obj_Bookmarks, obj_App, dogovordateend.Date.ToShortDateString());
                /*if (checkBox3.Checked = true)
                {
                    SetBookMarkText("Zagranpasport", obj_Bookmarks, obj_App, ", загранпаспорт");
                }*/
                SetBookMarkText("DogovorNum", obj_Bookmarks, obj_App, textBox7.Text);
                Param[0] = "true";
                obj_App.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, obj_App, Param);
                //object m=System.Type.Missing;
                Predarguments.setparam(Doc, Word, obj_App);
                
         //make zayvka
                object[] ExcelParam = new object[1];
                
                Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
                object Excel = Activator.CreateInstance(obj_excel);

                object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
                //ExcelParam[0] = Basepath + @"Template\zayavkaNaOlatyTyraNPred.xls";
                ExcelParam[0] = GetTempTemlate("Template", "zayavkaNaOlatyTyraNPred.xls");
                object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
                object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
                ExcelParam[0] = 1;
                object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
                /* SetCellData(comboBox6.Text,"D3",obj_worksheet);
                 SetCellData(textBox2.Text,"D4",obj_worksheet);
                 string date = textBox4.Text + "-" + textBox5.Text;
                 SetCellData(date,"D5",obj_worksheet);
                 if (dataGridView1.Rows[0].Cells[1].Value != null) { SetCellData(dataGridView1.Rows[0].Cells[1].Value.ToString(), "D6", obj_worksheet); }*/
                /*if (comboBox16.Text != "Росинтур")
                {
                    if (comboBox15.Text == "РосинтурЮг")
                    {
                        SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "A2", obj_worksheet);
                    }
                    if (comboBox15.Text == "Магазин Путешествий")
                    {
                        SetCellData("ООО ТК \"МАГАЗИН ПУТЕШЕСТВИЙ\"", "A2", obj_worksheet); ;
                    }
                    //SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "D24", obj_worksheet);
                }
                else if (comboBox16.Text == "Росинтур")
                {
                    SetCellData("ООО ТК \"РОСИНТУР\"", "A2", obj_worksheet);
                }*/
                if (comboBox16.Text == "Росинтур")
                {
                    SetCellData("ООО ТК \"РОСИНТУР\"", "A2", obj_worksheet);
                }
                if (comboBox16.Text == "РосинтурЮг")
                {
                    SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "A2", obj_worksheet);
                }
                if (comboBox16.Text == "Магазин Путешествий")
                {
                    SetCellData("ООО ТК \"МАГАЗИН ПУТЕШЕСТВИЙ\"", "A2", obj_worksheet);
                }
                SetCellData(comboBox17.Text, "H2", obj_worksheet);
                SetCellData(textBox38.Text, "D5", obj_worksheet);
                SetCellData(textBox37.Text, "D6", obj_worksheet);
                SetCellData(textBox36.Text, "D7", obj_worksheet);
                SetCellData(textBox35.Text, "D8", obj_worksheet);
                SetCellData("Предварительный договор № " + textBox7.Text, "D9", obj_worksheet);
                
                for (int i = 0; i < dataGridView16.RowCount; i++)
                {
                    if (dataGridView16.Rows[i].Cells[0].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[0].Value.ToString(), "A" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[1].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[1].Value.ToString(), "B" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[2].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[2].Value.ToString(), "C" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[3].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[3].Value.ToString(), "D" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[4].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[4].Value.ToString(), "E" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[5].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[5].Value.ToString(), "F" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[6].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[6].Value.ToString(), "G" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[7].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[7].Value.ToString(), "H" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[8].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[8].Value.ToString(), "I" + (13 + i), obj_worksheet); }
                    if (dataGridView16.Rows[i].Cells[9].Value != null) { SetCellData(dataGridView16.Rows[i].Cells[9].Value.ToString(), "J" + (13 + i), obj_worksheet); }
                }

                if (dataGridView15.Rows[0].Cells[0].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[0].Value.ToString(), "B19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[1].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[1].Value.ToString(), "C19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[2].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[2].Value.ToString(), "D19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[3].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[3].Value.ToString(), "E19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[4].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[4].Value.ToString(), "F19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[5].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[5].Value.ToString(), "G19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[6].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[6].Value.ToString(), "H19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[7].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[7].Value.ToString(), "I19", obj_worksheet); }
                if (dataGridView15.Rows[0].Cells[8].Value != null) { SetCellData(dataGridView15.Rows[0].Cells[8].Value.ToString(), "J19", obj_worksheet); }
                SetCellData("№ "+textBox7.Text, "B4", obj_worksheet);
                SetCellData("от " + comboBox11.Text + " " + comboBox10.Text + " "+numericUpDown5.Value.ToString(), "D4", obj_worksheet);
                SetCellData(comboBox14.Text, "B26", obj_worksheet);
                SetCellData(currency, "B20", obj_worksheet);
                SetCellData(textBox27.Text, "E20", obj_worksheet);
                if ((radioButton2.Checked == true) || (radioButton3.Checked == true))
                {
                    if ((dataGridView15.Rows[0].Cells[8].Value != null)&&(textBox27.Text!=""))
                    {
                        zRubSum = Convert.ToDouble(textBox27.Text) * Convert.ToDouble(dataGridView15.Rows[0].Cells[8].Value);
                    }
                }
                else
                {
                    if (dataGridView15.Rows[0].Cells[8].Value != null)
                    {
                        zRubSum = Convert.ToDouble(dataGridView15.Rows[0].Cells[8].Value);
                    }
                }
                SetCellData(Convert.ToInt32(zRubSum).ToString(), "J20", obj_worksheet);
            //razn
                //SetCellData("в " + currency, "B21", obj_worksheet);
                //SetCellData("в " + currency, "B22", obj_worksheet);
                double yedolg=0;
                double rusdolg=0;
                if ((textBox24.Text != "") && (textBox28.Text != ""))
                {
                    rusdolg = Convert.ToDouble(textBox24.Text) - Convert.ToDouble(textBox28.Text);
                }
                if ((textBox26.Text != "") && (textBox29.Text != ""))
                {
                    yedolg = Convert.ToDouble(textBox26.Text) - Convert.ToDouble(textBox29.Text);
                }
                SetCellData(yedolg.ToString(), "C22", obj_worksheet);
                SetCellData(rusdolg.ToString(), "J22", obj_worksheet);
                SetCellData(textBox28.Text, "J21", obj_worksheet);
                SetCellData(textBox29.Text, "C21", obj_worksheet);
            //razn
                SetCellData(Convert.ToInt32(zRubSum).ToString(), "J20", obj_worksheet);
                //SetCellData("Пердварите льный договор", "D7", obj_worksheet);
               /* if (comboBox16.Text == "Росинтур")
                {
                    SetCellData("ООО ТК \"РОСИНТУР\"", "D24", obj_worksheet);
                }
                else if (comboBox16.Text == "РосинтурЮг")
                {
                    SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "D24", obj_worksheet);
                }
                if (comboBox16.Text == "Магазин Путешествий")
                {
                    SetCellData("ООО ТК \"МАГАЗИН ПУТЕШЕСТВИЙ\"", "D24", obj_worksheet);
                }*/
                if (checkBox73.Checked == true)
                {
                    save_param += "(РБ)";
                }
                DocumentsaveA(Doc, obj_workbook, textBox7.Text, comboBox14.Text, comboBox9.Text, save_param);
                ExcelParam[0] = "True";
                Excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, Excel, ExcelParam);
                Predarguments.setparamE(Excel, obj_workbook);
                //make zayvka

                
                //DogNumber+
                try
                {
                    if (textBox7.Text != "")
                    {
                        if (comboBox16.Text == "Росинтур")
                        {
                            IncInINum(comboBox16.Text, textBox7.Text, "ClientDocCount");
                        }
                        else
                        {
                            IncInINum(comboBox15.Text, textBox7.Text, "ClientDocCount");
                        }
                    }
                }
                catch
                {

                }
                //DogNumber+


            //clean word
                Marshal.ReleaseComObject(obj_Tables);
                // Marshal.ReleaseComObject(obj_Selection);
                //Marshal.ReleaseComObject(obj_Range);
                Marshal.ReleaseComObject(obj_Doc);
                Marshal.ReleaseComObject(obj_Bookmarks);
                //Marshal.ReleaseComObject(obj_Bookmark);
                Marshal.ReleaseComObject(obj_App);
                // Marshal.ReleaseComObject(Word);
                  
                     
            //clean excel
                
                Marshal.ReleaseComObject(obj_workbooks);
                Marshal.ReleaseComObject(obj_workbook);
                Marshal.ReleaseComObject(obj_worksheet);
                Marshal.ReleaseComObject(obj_worksheets);

                //GC.GetTotalMemory(true);  
                button1.Enabled = true;
                button5.Enabled = true;
                //button6.Enabled = true;
                button13.Visible = false;

        }
//pred_dogovor
        private void tabControl2_Click(object sender, EventArgs e)
        {
            
            //this.dataGridView10.Rows[0].Cells[0].Value = textBox23.Text;
            this.textBox38.Text = this.comboBox9.Text;
            this.textBox34.Text = this.comboBox9.Text;
            this.textBox37.Text = this.textBox25.Text;
            this.textBox36.Text = this.dateTimePicker1.Text + "-" + this.dateTimePicker2.Text;
            if (this.dataGridView17.Rows[0].Cells[1].Value != null)
            {
                this.textBox35.Text = this.dataGridView17.Rows[0].Cells[1].Value.ToString();
            }
        }
        private string GetCellData(object worksheet, string cell)
        {
            try
            {
                object[] param = new object[1];
                param[0] = cell;
                string result = null;
                object range = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, param);
                result = range.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, range, null).ToString();
                return result;
            }
            catch
            {
                //MessageBox.Show("Ошибка получения данных с ячейки");
                richTextBox1.AppendText("Ошибка получения данных с ячейки\n\r");
                return "0";
            }
        }
        /*private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox20.Checked == false)
            {
                checkBox19.Enabled = true;
                checkBox18.Enabled = true;
            }
            if (checkBox20.Checked == true)
            {
                checkBox19.Enabled = false;
                checkBox18.Enabled = false;
                dataGridView16.Enabled = true;
                dataGridView15.Enabled = true;
            }
        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox19.Checked == false)
            {
                checkBox20.Enabled = true;
                checkBox18.Enabled = true;
              
            }
            if (checkBox19.Checked == true)
            {
                checkBox20.Enabled = false;
                checkBox18.Enabled = false;
                dataGridView16.Enabled = true;
                dataGridView15.Enabled = true;
            }
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked == false)
            {
                checkBox20.Enabled = true;
                checkBox19.Enabled = true;
               
            }
            if (checkBox18.Checked == true)
            {
                checkBox20.Enabled = false;
                checkBox19.Enabled = false;
                dataGridView16.Enabled = true;
                dataGridView15.Enabled = true;
            }
        }*/

        private void dataGridView16_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Regex r = new Regex("([0-9]+)");
            Match m=null;
            double sum = 0, sumcol = 0, sumrow = 0, sumdiscount = 0;
            //if ((e.ColumnIndex == 1) && (e.ColumnIndex == 2) && (e.ColumnIndex != 9))

            if ((e.ColumnIndex != 0) && (e.ColumnIndex != 3) && (e.ColumnIndex != 9))
            {
                if ((e.ColumnIndex == 2) || ((dataGridView16.Rows[e.RowIndex].Cells[1].Value != null) && (dataGridView16.Rows[e.RowIndex].Cells[2].Value != null)))
                {
                    //m = null;
                    m = r.Match(dataGridView16.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                    if (m.ToString() != "")
                    {

                        double dscount = ((Convert.ToDouble(dataGridView16.Rows[e.RowIndex].Cells[1].Value) / 100) * (Convert.ToDouble(dataGridView16.Rows[e.RowIndex].Cells[2].Value)));
                        //double f1 = (Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[1].Value) / 100);

                        //double dscount = f1 * (Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[2].Value));
                        //dataGridView7.Rows[e.RowIndex].Cells[3].Value = ((Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[1].Value) / 100) * (Convert.ToInt32(dataGridView7.Rows[e.RowIndex].Cells[2].Value)));

                        if (radioButton1.Checked == true)
                        {
                            dataGridView16.Rows[e.RowIndex].Cells[3].Value = Convert.ToInt32(dscount);
                        }
                        else
                        {
                            dataGridView16.Rows[e.RowIndex].Cells[3].Value = dscount;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Вводите Цифры");
                        dataGridView16.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                        return;
                    }
                }

                if (e.ColumnIndex != 2)
                {
                    //m = null;
                    for (int i = 0; i < dataGridView16.RowCount; i++)
                    {
                        if (dataGridView16.Rows[i].Cells[e.ColumnIndex].Value != null)
                        {
                            m = r.Match(dataGridView16.Rows[i].Cells[e.ColumnIndex].Value.ToString());
                            if (m.ToString() != "")
                            {
                                try
                                {
                                    sumrow = sumrow + Convert.ToDouble(dataGridView16.Rows[i].Cells[e.ColumnIndex].Value);
                                }
                                catch
                                {
                                    dataGridView16.Rows[i].Cells[e.ColumnIndex].Value = "";
                                    MessageBox.Show("Вводите Цифры");
                                }

                            }
                            else
                            {
                                MessageBox.Show("Вводите Цифры");
                                dataGridView16.Rows[i].Cells[e.ColumnIndex].Value = null;
                                return;
                            }
                        }
                    }
                    dataGridView15.Rows[0].Cells[e.ColumnIndex - 1].Value = sumrow;
                }

            }

            for (int i = 1; i < dataGridView16.ColumnCount - 1; i++)
            {
                if ((i != 2) && (i != 3))
                {
                    if (dataGridView16.Rows[e.RowIndex].Cells[i].Value != null)
                    {
                        m = null;
                        m = r.Match(dataGridView16.Rows[e.RowIndex].Cells[i].Value.ToString());
                        if (m.ToString() != "")
                        {
                            try
                            {
                                sumcol = sumcol + Convert.ToDouble(dataGridView16.Rows[e.RowIndex].Cells[i].Value);
                            }
                            catch
                            {
                                // MessageBox.Show("Вводите Цифры");
                            }
                        }
                        
                    }
                }
            }
            dataGridView16.Rows[e.RowIndex].Cells[9].Value = (sumcol - Convert.ToDouble(dataGridView16.Rows[e.RowIndex].Cells[3].Value));
            if (radioButton1.Checked == true)
            {
                dataGridView16.Rows[e.RowIndex].Cells[9].Value = Convert.ToInt32(dataGridView16.Rows[e.RowIndex].Cells[9].Value);
            }
            for (int i = 0; i < dataGridView16.RowCount; i++)
            {
                //  if (m == r.Match(dataGridView7.Rows[i].Cells[9].Value.ToString()))
                //{
                if ((dataGridView16.Rows[i].Cells[9].Value != null)&&(dataGridView16.Rows[i].Cells[9].Value.ToString() != ""))
                {
                    sum = sum + Convert.ToDouble(dataGridView16.Rows[i].Cells[9].Value);
                }


                if ((dataGridView16.Rows[i].Cells[3].Value != null) && (dataGridView16.Rows[i].Cells[3].Value.ToString() != ""))
                {
                    sumdiscount = sumdiscount + Convert.ToDouble(dataGridView16.Rows[i].Cells[3].Value);
                }
                //  }
            }
            dataGridView15.Rows[0].Cells[8].Value = sum;
            dataGridView15.Rows[0].Cells[2].Value = sumdiscount;
            if (radioButton1.Checked == true)
            {
                if (dataGridView15.Rows[0].Cells[8].Value != null)
                {
                    textBox26.Text = "";
                    textBox24.Text = dataGridView15.Rows[0].Cells[8].Value.ToString();
                }
            }
            else if ((radioButton2.Checked == true) || (radioButton3.Checked == true))
            {
                if (dataGridView15.Rows[0].Cells[8].Value != null)
                {
                    textBox24.Text = "";
                    textBox26.Text = dataGridView15.Rows[0].Cells[8].Value.ToString();
                }
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*object[] printarg = new object[9];
            printarg = new object[9] { Missing.Value, Missing.Value, 0, Missing.Value, Missing.Value, Missing.Value, Missing.Value, numericUpDown8.Value.ToString(), Missing.Value };
             */
            object Doc = Predarguments.Doc;
            Doc.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, Doc, null);
            object Work = Predarguments.WorkBook;
            Work.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, Work, null);
        }
        private void DocumentsaveA(object ODoc, object OWorkBook, string num, string managername, string client,string param)
        {
            Section sec = new Section();

            object[] WordParam = new object[1];
            object[] ExcelParam = new object[1];
            string path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini"));
            try
            {
                if (path != null)
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
            }
            catch
            {
                path = null;
            }
            if ((path != null) && (Directory.Exists(path)))
            {
                if (managername != "")
                {
                    path += "\\" + managername + "\\" + DateTime.Now.Year.ToString();
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                }
                /*if (checkBox11.Checked == true)
                {
                    ExcelParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Заявка " + num + "(проект)", ".xls");

                    WordParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Заявка " + num + "(проект)", ".doc");
                }
                else
                {
                    ExcelParam[0] = CheckFileName(path + "\\"+ "(" + client + ")"+"Заявка " + num+"("+DateTime.Now.Day+" " +((Month)DateTime.Now.Month).ToString()+" "+DateTime.Now.Year+")", ".xls");

                    WordParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Договор " + num + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".doc");
                }*/
                ExcelParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + param + "Заявка " + num + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".xls");

                WordParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + param + "Договор " + num + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".doc");
                ODoc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ODoc, WordParam);
                OWorkBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, OWorkBook, ExcelParam);

            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\КлиентскиеДоговора"))
                {
                    Directory.CreateDirectory("c:\\КлиентскиеДоговора");
                    localpath = "c:\\КлиентскиеДоговора";
                }
                else
                {
                    localpath = "c:\\КлиентскиеДоговора";
                }

                //object Doc = arguments.Doc;
                //object WordApp = arguments.App;
                WordParam[0] = CheckFileName(localpath + "\\" + "(" + client + ")" + "Договор " + num + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".doc");
                ODoc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ODoc, WordParam);
                ExcelParam[0] = CheckFileName(localpath + "\\" + "(" + client + ")" + "Заявка " + num + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".xls");
                OWorkBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, OWorkBook, ExcelParam);

            }
        }
        private void DocumentSave(object ODoc, object OWorkbook)
        {
            ODoc.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, ODoc, null);
            //ExcelParam[0] = localpath + "\\Заявка " + textBox7.Text;
            OWorkbook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, OWorkbook, null);
        }
        private string getsavepath()
        {
            Section sec = new Section();
            string path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini"));
            if ((path != null) && (Directory.Exists(path)))
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\КлиентскиеДоговора"))
                {
                    Directory.CreateDirectory("c:\\КлиентскиеДоговора");
                    localpath = "c:\\КлиентскиеДоговора";
                }
                else
                {
                    localpath = "c:\\КлиентскиеДоговора";
                }

            }
            return path;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            //Section sec = new Section();
            object[] WordParam = new object[1];
            object[] ExcelParam = new object[1];
            object Doc = Predarguments.Doc;
            object WordApp = Predarguments.App;
            object Workbook = Predarguments.WorkBook;
            //Marshal.GetActiveObject(Marshal.GenerateProgIdForType(GetType(Doc)));
            ////string path = Path.GetFullPath(sec.readkey("SavePath", "Path", "app.ini"));
            /*if ((path != null) && (Directory.Exists(path)))
            {
                ExcelParam[0] = CheckFileName(path + "\\Заявка " + textBox7.Text, ".xls");

                    //object Doc = arguments.Doc;
                    //object WordApp = arguments.App;
                WordParam[0] = CheckFileName(path + "\\Договор " + textBox7.Text, ".doc");
                    Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, WordParam);
                    Workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Workbook, ExcelParam);
            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\КлиентскиеДоговора"))
                {
                    Directory.CreateDirectory("c:\\КлиентскиеДоговора");
                    localpath = "c:\\КлиентскиеДоговора";
                }
                else
                {
                    localpath = "c:\\КлиентскиеДоговора";
                }

                    //object Doc = arguments.Doc;
                    //object WordApp = arguments.App;
                WordParam[0] = localpath + "\\Договор " + textBox7.Text;
                    Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, WordParam);
                    ExcelParam[0] = localpath + "\\Заявка " + textBox7.Text;
                    Workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Workbook, ExcelParam);

            }*/

            try
            {
                DocumentSave(Doc, Workbook);
                Doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, Doc, null);
                WordApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, WordApp, null);
                Workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, Workbook, null);
            }
            catch
            {

            }
            button1.Enabled = false;
            button5.Enabled = false;

            Marshal.ReleaseComObject(Predarguments.Doc);
            Marshal.ReleaseComObject(Predarguments.Word);
            Marshal.ReleaseComObject(Predarguments.WorkBook);
            Marshal.ReleaseComObject(Predarguments.Excel);
            Predarguments.clean();
            GC.GetTotalMemory(true);
            /*try
            {
                if (textBox7.Text != "")
                {
                    IncInINum(comboBox16.Text, textBox7.Text);
                }
            }
            catch
            {

            }*/

            if (checkBox22.Checked == false)
            {
                FormResetP();
            }
            checkBox22.Checked = false;
            button6.Enabled = false; ;
            //checkBox10.Checked = false;

            textBox7.Text = "";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox27.Text != "") && (textBox28.Text != ""))
                {
                    double yecena = Convert.ToDouble(textBox28.Text) / Convert.ToDouble(textBox27.Text);
                    textBox29.Text = formatDouble(yecena,3);
                }
                else
                {
                    MessageBox.Show("Введите курс и цену в рублях");
                }
            }
            catch
            {
                MessageBox.Show("Введите курс ввиде - 29,93 \r\n и цену в рублях ввиде - 1000,05");
                textBox27.Text = "";
            }
        }
        private string formatDouble(double obj,int num)
        {
            string[] str = new string[2] {"0","0"};
            string[] temp = obj.ToString().Split(',');
            if (temp.Length == 1)
            {
                str[0] = temp[0];
            }
            else
            {
                str = temp;
            }
            //int part1 = (int)(obj);
            //string part2 = Convert.ToString(obj -  (double)part1);
            //string temp = part2.ToString();
            if (str[1].Length >= num)
            {
                str[1] = str[1].Substring(0, num);
            }
            return string.Join(",",str);
            //str=String.Join(
            //return part1.ToString()+","+temp;
           // str[1]=str[1].su
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            button8.Enabled = true;
            button9.Enabled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox14.Text != "") && (textBox15.Text != ""))
                {
                    double sum = Convert.ToDouble(textBox14.Text) / Convert.ToDouble(textBox15.Text);
                    textBox13.Text = sum.ToString();
                }
                else
                {
                    MessageBox.Show("Введите курс и сумму в рублях");
                }
            }
            catch 
            {
                MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в рублях ввиде - 1000,05");
                textBox15.Text = "";
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox13.Text != "") && (textBox15.Text != ""))
                {
                    double sum = Convert.ToDouble(textBox13.Text) * Convert.ToDouble(textBox15.Text);
                    textBox14.Text = Convert.ToInt32(sum).ToString();
                }
                else
                {
                    MessageBox.Show("Введите курс и сумму в рублях");
                }
            }
            catch
            {
                MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в y.e ввиде - 1000,05");
                textBox15.Text = "";
            }
        }

        private void dataGridView14_Enter(object sender, EventArgs e)
        {
            this.dataGridView14.Rows[0].Cells[3].Value = dateTimePicker3.Text + "-" + dateTimePicker4.Text;
        }

        private void dataGridView17_Enter(object sender, EventArgs e)
        {
            this.dataGridView17.Rows[0].Cells[3].Value = dateTimePicker1.Text + "-" + dateTimePicker2.Text;
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {

            textBox7.Text = "";

            if ((comboBox16.Text == "Росинтур") || (comboBox16.Text == "Магазин Путешествий"))
            {
                textBox7.Text = MakeDogovorNum(comboBox14.Text, comboBox16.Text, 1);
            }
            else
            {
                textBox7.Text = MakeDogovorNum(comboBox14.Text, comboBox15.Text, 1);
            }
            /*
            if (comboBox16.Text == "РосинтурЮг")
            {
                textBox7.Text = "Ю-";
            }
            if (comboBox16.Text == "Магазин Путешествий")
            {
                textBox7.Text = "МП-";
            }*/
            //textBox7.Text = GetCompanyID(comboBox16.Text);
            //textBox7.Text += MakeDogovorNum(comboBox14.Text, comboBox16.Text,1);
            button6.Enabled = true;
        }
        private string GetCompanyID(string name)
        {
            string id = null;
            if (name == "Росинтур")
            {
                id = "Р-";
            }
            if (name == "РосинтурЮг")
            {
                id = "Ю-";
            }
            if (name == "Магазин Путешествий")
            {
                id = "МП-";
            }
            return id;
        }
        private void IncInINum(string company,string num,string key)
        {
            Section sec = new Section();
            string[] part = num.Split('-');
            //int number = Convert.ToInt32(sec.readkey("ClientDocCount", "number_"+company, "app.ini"));
            int number=Convert.ToInt32(part[1]);
            number++;
            sec.writekey(key, "number_"+company, "app.ini", number.ToString());
        }
        private string MakeDogovorNum(string manager, string company, int type)
        {
            Section sec = new Section();
            string Dtype=null;
            string companyid = null;
            string result = null;
            string num = sec.readkey("ClientDocCount", "number_"+company, "app.ini");
            //string Id = GetmangerId(manager);
            companyid = GetCompanyID(company);
            if (type == 1)
            {
               Dtype = "П";
            }
            else if (type == 0)
            {
               Dtype = "О";
            }
            //return result = companyid + Id + "-" + num + "-" + Dtype;
            return result = companyid  + num + "-" + Dtype;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*if (checkBox26.Checked == false)
            {
                textBox1.Text = "";
                if (comboBox3.Text == "Росинтур")
                {
                    //textBox1.Text = "Р-";
                    label95.Text = "Р-" + textBox1.Text + "-П";
                }
                if ((comboBox3.Text != "Росинтур") && (comboBox4.Text == "РосинтурЮг"))
                {
                    //textBox1.Text = "Ю-";
                    label95.Text = "Ю-" + textBox1.Text + "-П";
                }
                if ((comboBox3.Text == "Магазин Путешествий") || (comboBox4.Text == "Магазин Путешествий"))
                {
                    //textBox1.Text = "МП-";
                    label95.Text = "МП-" + textBox1.Text + "-П";
                }
                //textBox1.Text += GetmangerId(comboBox5.Text) + "-";
            }*/
        }

        private void dataGridView18_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView16.RowCount = dataGridView18.RowCount;
            for (int i = 0; i < dataGridView18.RowCount; i++)
            {

                // if (dataGridView6.Rows[i].Cells[0]!=null)
                // {
                dataGridView16.Rows[i].Cells[0].Value = dataGridView18.Rows[i].Cells[0].Value;
                // }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {

                try
                {
                    if (radioButton4.Checked == true)
                    {
                        if ((textBox14.Text != "") && (textBox43.Text != ""))
                        {
                            double sum = Convert.ToDouble(textBox14.Text) - Convert.ToDouble(textBox43.Text);
                            textBox46.Text = Convert.ToInt32(sum).ToString();
                        }
                    }
                    else if ((radioButton5.Checked == true)||(radioButton6.Checked == true))
                    {
                        if ((textBox47.Text != "") && (textBox45.Text != ""))
                        {
                            double sum = Convert.ToDouble(textBox47.Text) * Convert.ToDouble(textBox45.Text);
                            textBox46.Text = Convert.ToInt32(sum).ToString();
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в y.e ввиде - 1000,05");
                    textBox45.Text = "";
                }
        }
        //get data
        private void button10_Click(object sender, EventArgs e)
        {
            
            Section sec = new Section();
            object[] ExcelParam = new object[1];

            try
            {
                if (comboBox5.Text != "")
                {
                    string path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini") + "\\" + comboBox5.Text);
                    if ((path != null) && (Directory.Exists(path)))
                    {

                        string filename = path + "\\Заявка " + textBox1.Text + "(" + comboBox6.Text + ")" + ".xls";
                        if (File.Exists(filename))
                        {
                            Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
                            object Excel = Activator.CreateInstance(obj_excel);

                            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
                            ExcelParam[0] = filename;
                            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
                            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
                            ExcelParam[0] = 1;
                            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
                            textBox44.Text = GetCellData(obj_worksheet, "C21");
                            textBox43.Text = GetCellData(obj_worksheet, "J21");
                            textBox20.Text = GetCellData(obj_worksheet, "E20");

                            obj_workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, obj_workbook, null);
                            Marshal.ReleaseComObject(obj_worksheet);
                            Marshal.ReleaseComObject(obj_worksheets);
                            Marshal.ReleaseComObject(obj_workbook);
                            Marshal.ReleaseComObject(obj_workbooks);
                            //textBox47.Text = GetCellData(obj_worksheet, "C22");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Выберите менеджера");
                }
            }
            catch
            {
                MessageBox.Show("Ошибка при открытии файла. Возможно не правильно указано имя");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {/*
            string m = null;
            string[] strobj;
            TextBox obj=(TextBox)sender;
            string result = null;

            Regex r=new Regex("[А-Я]+(-)[0-9][0-9]-[0-9]+(-П)");
            m=r.Match(obj.Text).ToString();
            if ((m.ToString() != "") && (m != null))
            {
                strobj = m.Split('-');
                strobj[3] = "О";
                result = string.Join("-", strobj);
                textBox49.Text = result;
                button10.Enabled = true;
                //button2.Enabled = true;
            }
            else
            {
                button10.Enabled = false;
            }
           */
            if ((checkBox26.Checked == false)&&(textBox1.Text!=""))
            {
                if (comboBox3.Text == "Росинтур")
                {
                    //textBox1.Text = "Р-";
                    label195.Text = "Р-" + textBox1.Text + "-П";
                }
                if ((comboBox3.Text != "Росинтур") && (comboBox4.Text == "РосинтурЮг"))
                {
                    //textBox1.Text = "Ю-";
                    label195.Text = "Ю-" + textBox1.Text + "-П";
                }
                if ((comboBox3.Text == "Магазин Путешествий") || (comboBox4.Text == "Магазин Путешествий"))
                {
                    //textBox1.Text = "МП-";
                    label195.Text = "МП-" + textBox1.Text + "-П";
                }
            }

        }

        private void textBox15_TextChanged_1(object sender, EventArgs e)
        {
            if ((radioButton5.Checked == true) || (radioButton6.Checked == true))
            {
                textBox45.Text = textBox15.Text;
            }
        }

        private void папкаСДоговорамиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Section sec = new Section();
            string path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini"));
            if ((path != null) && (Directory.Exists(path)))
            {
                Process.Start(path);
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox27.Text != "") && (textBox26.Text != ""))
                {
                    double sum = Convert.ToDouble(textBox27.Text) * Convert.ToDouble(textBox26.Text);
                    textBox24.Text = Convert.ToInt32(sum).ToString();
                }
                else
                {
                    MessageBox.Show("Введите курс и сумму в y.e.");
                }
            }
            catch
            {
                MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в y.e ввиде - 1000,05");
                textBox26.Text = "";
            }
        }

        private void m_click(object sender, DateRangeEventArgs e)
        {
            //this.dataget
            DataGridView d1 = (DataGridView)databox.owner;
            // d1.RowCount = d1.RowCount + 1;
            DataGridViewCell c1 = d1.Rows[databox.args.RowIndex].Cells[databox.args.ColumnIndex];
            c1.Value = e.End.ToShortDateString();
            System.Windows.Forms.MonthCalendar m = (MonthCalendar)sender;
            Control p = m.Parent;
            p.Hide();
            databox.clean();
        }

        private void dataGridViewCalendar_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                databox.Set(sender, e);
                DataGridView d1 = (DataGridView)sender;
                DataGridViewCell c1 = d1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                Point point = d1.Location;
                Rectangle rect = new Rectangle();
                rect.X = point.X + c1.Size.Width / 2;
                rect.Y = point.Y + c1.Size.Height / 2;
                this.calPanel.Location = rect.Location;
                //p1.Location.Y = rect.Y + 10;


                this.calPanel.Show();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            button6.Enabled = true;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            button2.Enabled = true;
        }

        private void checkBox26_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox26.Checked == true)
            {
                //button15.Enabled = true;
                textBox1.Text = "";
                textBox49.Text = "";
                button10.Enabled = false;
                if (comboBox5.Text != "")
                {
                    if ((comboBox3.Text == "Росинтур") || (comboBox3.Text == "Магазин Путешествий"))
                    {
                        textBox49.Text = MakeDogovorNum(comboBox5.Text, comboBox3.Text, 0);
                    }
                    else
                    {
                        textBox49.Text = MakeDogovorNum(comboBox5.Text, comboBox4.Text, 0);
                    }
                }
                else
                {
                    MessageBox.Show("Выберете менеджера");
                    checkBox26.Checked = false;
                }
            }
            else
            {
                textBox49.Text = "";
                //button15.Enabled = false;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            /*if (comboBox5.Text != "")
            {
                if ((comboBox3.Text == "Росинтур") || (comboBox3.Text == "Магазин Путешествий"))
                {
                    textBox49.Text = MakeDogovorNum(comboBox5.Text, comboBox3.Text, 0);
                }
                else
                {
                    textBox49.Text = MakeDogovorNum(comboBox5.Text, comboBox4.Text, 0);
                }
            }
            else
            {
                MessageBox.Show("Выберете менеджера");
            }*/
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //1
            comboBox6.Text = comboBox9.Text;
            textBox2.Text = textBox25.Text;
            dateTimePicker3.Text = dateTimePicker1.Text;
            dateTimePicker2.Text = dateTimePicker4.Text;
            textBox6.Text = textBox21.Text;
            checkBox1.CheckState = checkBox14.CheckState;
            checkBox3.CheckState = checkBox12.CheckState;
            DataGridDataCopy(dataGridView17, dataGridView14);
            //2
            DataGridDataCopy(dataGridView13, dataGridView1);
            DataGridDataCopy(dataGridView12, dataGridView2);
            DataGridDataCopy(dataGridView11, dataGridView3);
            //DataGridDataCopy(dataGridView17, dataGridView14);
            checkBox4.CheckState = checkBox17.CheckState;
            checkBox5.CheckState = checkBox16.CheckState;
            checkBox6.CheckState = checkBox15.CheckState;
            DataGridDataCopy(dataGridView10, dataGridView4);
            for (int j = 0; j < dataGridView9.ColumnCount; j++)
            {
                if (dataGridView9.Rows[0].Cells[j].Value != null)
                {
                    dataGridView5.Rows[0].Cells[j].Value = dataGridView9.Rows[0].Cells[j].Value;
                }
            }
            //3
            textBox8.Text = textBox34.Text;
            textBox113.Text = maskedTextBox3.Text;
            textBox9.Text = textBox33.Text;
            maskedTextBox1.Text = maskedTextBox3.Text;
            textBox110.Text = textBox107.Text;
            textBox111.Text = textBox108.Text;
            textBox112.Text = textBox109.Text;
            textBox117.Text = textBox122.Text;
            textBox116.Text = textBox121.Text;
            textBox115.Text = textBox120.Text;
            textBox114.Text = textBox119.Text;
            textBox11.Text = textBox218.Text;
            textBox118.Text = textBox123.Text;
            textBox217.Text = textBox31.Text;
            textBox10.Text = textBox32.Text;
            maskedTextBox2.Text = maskedTextBox4.Text;
            textBox12.Text = textBox30.Text;
            DataGridDataCopy(dataGridView18, dataGridView6);
            //4
            textBox16.Text = textBox38.Text;
            textBox17.Text = textBox37.Text;
            textBox18.Text = textBox36.Text;
            textBox19.Text = textBox35.Text;
            comboBox19.Text = comboBox17.Text;
            radioButton4.Checked = radioButton1.Checked;
            radioButton5.Checked = radioButton2.Checked;
            radioButton6.Checked = radioButton3.Checked;
            DataGridDataCopy(dataGridView16, dataGridView7);
            DataGridDataCopy(dataGridView15, dataGridView8);
            textBox15.Text = textBox27.Text;
            textBox20.Text = textBox27.Text;
            textBox13.Text = textBox26.Text;
            textBox14.Text = textBox24.Text;
            textBox27.Text = textBox15.Text;
            textBox44.Text = textBox29.Text;
            textBox43.Text = textBox28.Text;
            //
            groupBox1.Visible = true;
            groupBox2.Visible = false;
            groupBox3.Visible = false;

        }
        private void button36_Click(object sender, EventArgs e)
        {
            //1
            comboBox9.Text = comboBox6.Text;
            textBox25.Text = textBox2.Text;
            dateTimePicker1.Text = dateTimePicker3.Text;
            dateTimePicker4.Text = dateTimePicker2.Text;
            textBox21.Text = textBox6.Text;
            checkBox14.CheckState = checkBox1.CheckState;
            checkBox12.CheckState = checkBox3.CheckState;
            DataGridDataCopy(dataGridView14, dataGridView17);
            //2
            DataGridDataCopy(dataGridView1, dataGridView13);
            DataGridDataCopy(dataGridView2, dataGridView12);
            DataGridDataCopy(dataGridView3, dataGridView11);
            //DataGridDataCopy(dataGridView17, dataGridView14);
            checkBox17.CheckState = checkBox4.CheckState;
            checkBox16.CheckState = checkBox5.CheckState;
            checkBox15.CheckState = checkBox6.CheckState;
            DataGridDataCopy(dataGridView4, dataGridView10);
            for (int j = 0; j < dataGridView5.ColumnCount; j++)
            {
                if (dataGridView5.Rows[0].Cells[j].Value != null)
                {
                    dataGridView9.Rows[0].Cells[j].Value = dataGridView5.Rows[0].Cells[j].Value;
                }
            }
            //3
            textBox34.Text = textBox8.Text;
            maskedTextBox3.Text = textBox113.Text;
            textBox33.Text = textBox9.Text;
            maskedTextBox3.Text = maskedTextBox1.Text;
            textBox107.Text = textBox110.Text;
            textBox108.Text = textBox111.Text;
            textBox109.Text = textBox112.Text;
            textBox122.Text = textBox117.Text;
            textBox121.Text = textBox116.Text;
            textBox120.Text = textBox115.Text;
            textBox119.Text = textBox114.Text;
            textBox218.Text = textBox11.Text;
            textBox123.Text = textBox118.Text;
            textBox31.Text = textBox217.Text;
            textBox32.Text = textBox10.Text;
            maskedTextBox4.Text = maskedTextBox2.Text;
            textBox30.Text = textBox12.Text;
            DataGridDataCopy(dataGridView6, dataGridView18);
            //4
            textBox38.Text = textBox16.Text;
            textBox37.Text = textBox17.Text;
            textBox36.Text = textBox18.Text;
            textBox35.Text = textBox19.Text;
            comboBox17.Text = comboBox19.Text;
            radioButton1.Checked = radioButton4.Checked;
            radioButton2.Checked = radioButton5.Checked;
            radioButton3.Checked = radioButton6.Checked;
            DataGridDataCopy(dataGridView7, dataGridView16);
            DataGridDataCopy(dataGridView8, dataGridView15);
            textBox27.Text = textBox15.Text;
            textBox27.Text = textBox20.Text;
            textBox26.Text = textBox13.Text;
            textBox24.Text = textBox14.Text;
            textBox15.Text = textBox27.Text;
            textBox29.Text = textBox44.Text;
            textBox28.Text = textBox43.Text;
            //
            groupBox2.Visible = true;
            groupBox1.Visible = false;
            groupBox3.Visible = false;
        }
        private void DataGridDataCopy(DataGridView source,DataGridView dest)
        {
            dest.RowCount = source.RowCount;
            dest.ColumnCount = source.ColumnCount;
            for (int i = 0; i < source.RowCount; i++)
            {
                for (int j = 0; j < source.ColumnCount; j++)
                {
                    if (source.Rows[i].Cells[j].Value != null)
                    {
                        dest.Rows[i].Cells[j].Value = source.Rows[i].Cells[j].Value;
                    }
                }
            }
        }

        private void анкетаДляКонсультваToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = true;
            groupBox17.Visible = false;
        }

        private void textBox61_TextChanged(object sender, EventArgs e)
        {
            textBox62.Text = GetWordSum(textBox62.Text);
        }

        private string GetWordSum(string sourcestr)
        {
            string str = "";string result="";
            TranslateData t = new TranslateData();
            if (sourcestr != "")
            {
                try
                {
                    str = Convert.ToInt32(sourcestr).ToString();
                }
                catch
                {
                    string error = int.MaxValue.ToString();
                    MessageBox.Show("Введите значение меньше " + error);
                }

            }
            if (str != "")
            {
                result = t.TranslateStr(str);
            }
            return result;
        }
        //KonsulDogovor
        private void button20_Click(object sender, EventArgs e)
        {
            //string[] client1 = getclientdata(1, textBox23.Text, textBox50.Text, textBox51.Text, textBox52.Text, textBox55.Text, textBox53.Text, textBox54.Text);
            //button13.Visible = true;
           //button6.Enabled = false;
            object obj_App;
            object obj_Doc;
            object obj_Bookmarks;
            //object obj_Bookmark;
            //object obj_Selection;
            //object obj_Range;
            //object obj_Tables;
            object[] Param;
            Param = new object[1];
            Type obj_Class = Type.GetTypeFromProgID("Word.Application");
            object Word = Activator.CreateInstance(obj_Class);

            obj_App = Word.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, Word, null);
            obj_Doc = obj_App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj_App, null);
            if (comboBox25.Text == "Клиентский")
            {
                Param[0] = Basepath + @"Template\KonsulDog.doc";
            }
            else if (comboBox25.Text == "Агентский")
            {
                Param[0] = Basepath + @"Template\KonsulDogAgent.doc";
            }
            object Doc = obj_Doc.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_Doc, Param);
            obj_Bookmarks = Doc.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, Doc, null);
            SetBookMarkText("DayNow", obj_Bookmarks, obj_App, this.comboBox8.SelectedItem.ToString());
            SetBookMarkText("MonthNow", obj_Bookmarks, obj_App, this.comboBox7.SelectedItem.ToString());
            SetBookMarkText("YearNow", obj_Bookmarks, obj_App, this.numericUpDown3.Value.ToString());
            SetBookMarkText("number", obj_Bookmarks, obj_App, this.textBox22.Text);
            SetBookMarkText("HomePhone", obj_Bookmarks, obj_App, this.textBox104.Text);
            SetBookMarkText("WorkPhone", obj_Bookmarks, obj_App, this.textBox105.Text);
            SetBookMarkText("MobilePhone", obj_Bookmarks, obj_App, this.textBox106.Text);
            SetBookMarkText("client", obj_Bookmarks, obj_App, textBox23.Text);
            //string client1 = getclientdata(textBox23.Text, textBox50.Text, textBox51.Text,textBox52.Text,textBox55.Text,textBox53.Text,textBox54.Text);
            string[] client1 = getclientdata(1, textBox23.Text, textBox50.Text, textBox51.Text, textBox52.Text, textBox55.Text, textBox53.Text, textBox54.Text);
            //if ((client1[0]!="")&&(client1[1]!=""))
            //{
                SetBookMarkText("Client1s1", obj_Bookmarks, obj_App, client1[0]);
                SetBookMarkText("Client1s2", obj_Bookmarks, obj_App, client1[1]);
           // }
                SetBookMarkText("Adress", obj_Bookmarks, obj_App, textBox54.Text);
            string[] client2 = getclientdata(2, textBox83.Text, textBox82.Text, textBox81.Text, textBox80.Text, textBox77.Text, textBox79.Text, textBox78.Text);
            SetBookMarkText("Client2s1", obj_Bookmarks, obj_App, client2[0]);
            SetBookMarkText("Client2s2", obj_Bookmarks, obj_App, client2[1]);
            string[] client3 = getclientdata(3, textBox76.Text, textBox75.Text, textBox74.Text, textBox73.Text, textBox70.Text, textBox72.Text, textBox71.Text);
            SetBookMarkText("Client3s1", obj_Bookmarks, obj_App, client3[0]);
            SetBookMarkText("Client3s2", obj_Bookmarks, obj_App, client3[1]);
            string[] client4 = getclientdata(4, textBox69.Text, textBox68.Text, textBox67.Text, textBox66.Text, textBox63.Text, textBox65.Text, textBox64.Text);
            SetBookMarkText("Client4s1", obj_Bookmarks, obj_App, client4[0]);
            SetBookMarkText("Client4s2", obj_Bookmarks, obj_App, client4[1]);
            if (textBox23.Text != "")
            {
                SetBookMarkText("Client1Short", obj_Bookmarks, obj_App, getclientshortname(textBox23.Text));
            }
            if (textBox83.Text != "")
            {
                SetBookMarkText("Client2Short", obj_Bookmarks, obj_App, getclientshortname(textBox83.Text));
            }
            if (textBox76.Text != "")
            {
                SetBookMarkText("Client3Short", obj_Bookmarks, obj_App, getclientshortname(textBox76.Text));
            }
            if (textBox69.Text != "")
            {
                SetBookMarkText("Client4Short", obj_Bookmarks, obj_App, getclientshortname(textBox69.Text));
            }
            SetBookMarkText("TravelPlace", obj_Bookmarks, obj_App, this.textBox56.Text);
            SetBookMarkText("TravelProgram", obj_Bookmarks, obj_App, this.textBox57.Text);
            SetBookMarkText("TravelDate", obj_Bookmarks, obj_App, this.dateTimePicker5.Text + "-" + this.dateTimePicker6.Text);
            SetBookMarkText("TravelDoc", obj_Bookmarks, obj_App, this.textBox58.Text);
            SetBookMarkText("Hotel", obj_Bookmarks, obj_App, label119.Text);
            SetBookMarkText("RoomCategory", obj_Bookmarks, obj_App, comboBox21.Text);
            SetBookMarkText("NomerType", obj_Bookmarks, obj_App, comboBox22.Text);
            SetBookMarkText("FoodType", obj_Bookmarks, obj_App, comboBox23.Text);
            SetBookMarkText("Excursion", obj_Bookmarks, obj_App, comboBox23.Text);
            SetBookMarkText("DopService", obj_Bookmarks, obj_App, textBox60.Text);
            SetBookMarkText("Viza", obj_Bookmarks, obj_App, getCheckYesNo(checkBox27));
            SetBookMarkText("ExcursionGid", obj_Bookmarks, obj_App, getCheckYesNo(checkBox28));
            SetBookMarkText("Medstrahovka", obj_Bookmarks, obj_App, getCheckYesNo(checkBox29));
            SetBookMarkText("Strahovka", obj_Bookmarks, obj_App, getCheckYesNo(checkBox30));
            SetBookMarkText("Transfer", obj_Bookmarks, obj_App, comboBox24.Text);
            SetBookMarkText("Symma", obj_Bookmarks, obj_App, textBox61.Text + "(" + textBox62.Text+")");
            if (comboBox25.Text == "Агентский")
            {
                SetBookMarkText("Tyrcompany", obj_Bookmarks, obj_App, textBox5.Text+" "+textBox103.Text);
                SetBookMarkText("TyrcompanyName", obj_Bookmarks, obj_App, textBox5.Text);
                SetBookMarkText("GenDirect", obj_Bookmarks, obj_App, textBox4.Text);
                SetBookMarkText("GenDirectShort", obj_Bookmarks, obj_App, getclientshortname(textBox4.Text));
                SetBookMarkText("osnovanie", obj_Bookmarks, obj_App, textBox102.Text);
                SetBookMarkText("MailAdress", obj_Bookmarks, obj_App, textBox84.Text);
                SetBookMarkText("Phone", obj_Bookmarks, obj_App, textBox85.Text);
                SetBookMarkText("Fax", obj_Bookmarks, obj_App, textBox86.Text);
                SetBookMarkText("UrAdress", obj_Bookmarks, obj_App, textBox87.Text);
                SetBookMarkText("Email", obj_Bookmarks, obj_App, textBox88.Text);
                SetBookMarkText("INN", obj_Bookmarks, obj_App, textBox96.Text);
                SetBookMarkText("KPP", obj_Bookmarks, obj_App, textBox97.Text);
                SetBookMarkText("bank", obj_Bookmarks, obj_App, textBox98.Text);
                SetBookMarkText("rs", obj_Bookmarks, obj_App, textBox99.Text);
                SetBookMarkText("kors", obj_Bookmarks, obj_App, textBox100.Text);
                SetBookMarkText("BIK", obj_Bookmarks, obj_App, textBox101.Text);
            }

           

            //DocumentsaveA(Doc, obj_workbook, textBox7.Text, comboBox14.Text, comboBox9.Text);
            object[] WordParam = new object[1];
            string path = getsavepath();
            path += "\\Консульские Договора";
            if (Directory.Exists(path))
            {
                WordParam[0] = CheckFileName(path + "\\Консульский Договор " + "(" + textBox23.Text + ")", ".doc");
                Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, WordParam);
            }
            else
            {
                Directory.CreateDirectory(path);
                WordParam[0] = CheckFileName(path + "\\Консульский Договор " + "(" + textBox23.Text + ")", ".doc");
                Doc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, Doc, WordParam);
            }
            Param[0] = "true";
            obj_App.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, obj_App, Param);
            /*//DatasaveSQL
            Client ClientData = new Client (textBox34.Text,"","","","",textBox33.Text,textBox107.Text,textBox108.Text,textBox109.Text,"",textBox30.Text,textBox31.Text,"",textBox32.Text);
            object id = GetClientId(ClientData); 
            DogovorInfo dinfo=new DogovorInfo(DateTime.Now.Day.ToString()+"."+DateTime.Now.Month.ToString()+"."+DateTime.Now.Year.ToString(),textBox25.Text,textBox21.Text,dateTimePicker1.Text,dateTimePicker2.Text,textBox19.Text,comboBox19.Text,
            //DatasaveSQLEnd*/

            //Predarguments.setparamE(Excel, obj_workbook);
            //make zayvka

            /*
            try
            {
                if (textBox7.Text != "")
                {
                    IncInINum(comboBox16.Text, textBox7.Text);
                }
            }
            catch
            {

            }*/
            //clean word
            // Marshal.ReleaseComObject(obj_Selection);
            //Marshal.ReleaseComObject(obj_Range);
            Marshal.ReleaseComObject(obj_Doc);
            Marshal.ReleaseComObject(obj_Bookmarks);
            //Marshal.ReleaseComObject(obj_Bookmark);
            Marshal.ReleaseComObject(obj_App);
            // Marshal.ReleaseComObject(Word);
            // GC.GetTotalMemory(true);    
              
           
        }
        private string getCheckYesNo(CheckBox check)
        {
            string result = "";
            if (check.Checked == true)
            {
                result= "ДА";
            }
            else
            {
                result="НЕТ";
            }
            return result;
        }
        private string getclientshortname(string str)
        {
            string result="";
            //char[] c=new char[3]{' ', ',','/'};
            string[] temp = str.Split(' ');
            if (temp.Length == 3)
            {
                result = temp[0] + " " + temp[1].Substring(0, 1) + ". " + temp[2].Substring(0, 1) + ".";
            }
            if (temp.Length == 2)
            {
                result = temp[0] + " " + temp[1].Substring(0, 1) + ".";
            }
            return result;
        }
        private string[] getclientdata(int num,string FIO, string birthdate,string paspserie, string paspnum, string paspdate, string paspown, string adress)
        {
            string[] result= new string[2];
            string str = "";
            /*//ComponentCollection c = groupBox4.Container.Components;
            System.Collections.IEnumerator i = groupBox4.Controls.GetEnumerator();
            i.MoveNext();
           object o = i.Current;*/
            if ((FIO != "") && (birthdate != "") && (paspserie != "") && (paspnum!=""))
            {
                str += num + ". " + FIO + " " + birthdate + " " + paspserie + " " + paspnum  + " выдан: " + paspown + " " + paspdate +" Адрес:" + adress;
                /* for (int i = 0; i < 7;i++ )
                 {

                 }*/
                int pos = 0;
                if (str.Length < 110)
                {
                    result[0] = str;
                    result[1] = "";
                }
                else
                {
                    for (int i = 0; i < 110; i++)
                    {
                        if (str[i] == ' ')
                        {
                            pos = i;
                        }
                    }
                    result[0] = str.Substring(0, pos);
                    // int size = str.Length - 1;
                    result[1] = str.Substring(pos, str.Length - 1 - pos);
                }
            }
            else
            {
                result[0] = "";
                result[1] = "";
            }
            return result;
        }
        //fileHotelwork
        private void inicializedict()
        {
            if (File.Exists("Addr.t"))
            {
                FileStream fs = new FileStream("Addr.t", FileMode.Open, FileAccess.Read);
                byte[] readbuffer=new byte[fs.Length];
                fs.Read(readbuffer, 0, (int)fs.Length);
                //readbuffer = Encoding.Convert(Encoding.GetEncoding(65001), Encoding.GetEncoding(28595), readbuffer);
               // string s = GetStringfromByte(readbuffer);
                string s = Encoding.Default.GetString(readbuffer);
                int startindex=0;
                for (int i = 0; i < s.Length;i++)
                {
                    if (s[i] == '[') 
                    {
                        startindex=i;
                        while (s[i]!=']')
                        {
                            i++;
                        }
                        //string str = s.Substring(startindex + 1, i - startindex - 1);
                       //comboBox20.Items.Add(s.Substring(startindex+1, i - startindex-1));
                        dict.Add(s.Substring(startindex + 1, i - startindex - 1), startindex);
                    }
                }
                fs.Close();
            }
            
                //SortedDictionary<string, int>.KeyCollection key = new SortedDictionary<string, int>();
            foreach (string key in dict.Keys)
            {
                //dict.TryGetValue(key,out value);
                comboBox20.Items.Add(key);
            }
        }
        private string GetStringfromByte(byte[] source)
        {
            
            string result="";
            for (int i = 0; i < source.Length; i++)
            {
                 result+= Convert.ToChar(source[i]);
            }
            return result;
        }
        private string GetItem(int startindex)
        {
            Section sec = new Section();
            string result = "";
            FileStream fs = new FileStream("Addr.t", FileMode.Open, FileAccess.Read);
            byte[] readbuffer = new byte[fs.Length];
            fs.Read(readbuffer, 0, (int)fs.Length);
            fs.Close();
            int i = startindex+1;
            string s = Encoding.Default.GetString(readbuffer);
            while (s[i]!= 13)
            {
                if (s[i] != ']')
                {
                    result += Convert.ToChar(s[i]);
                }
                else
                {
                    result += ", ";
                }
                i++;

            }

            return result;
        }
        private void additem(string source)
        {
            Section sec = new Section();
            checkendline("Addr.t");
            FileStream fs = new FileStream("Addr.t", FileMode.Append, FileAccess.Write);
            byte[] writebuffer = sec.strTobyte(source);
            //writebuffer = sec.strTobyte(source);
            fs.Write(writebuffer, 0, writebuffer.Length);
            fs.Close();
        }
        private void checkendline(string str)
        {
            Section sec = new Section();
            byte[] temp=new byte[4];
            FileStream fs = new FileStream(str, FileMode.Open, FileAccess.Read);
            byte[] source=new byte[fs.Length];
            fs.Read(source,0,(int)fs.Length);
            fs.Close();
            for(int i=0;i<4;i++)
            {
                temp[i]=source[source.Length-5];
            }
            if (sec.byteTostr(temp) != "\r\n")
            {
                fs = new FileStream("Addr.t", FileMode.Append, FileAccess.Write);
                string addstr="\r\n";
                byte[] writebyte=sec.strTobyte(addstr);
                fs.Write(writebyte, 0, writebyte.Length);
                fs.Close();
            }
        }





        //fileHotelwork

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                if ((radioButton5.Checked == true) || (radioButton6.Checked == true))
                {
                    if ((textBox13.Text != "") && (textBox44.Text != ""))
                    {
                        double sum = Convert.ToDouble(textBox13.Text) - Convert.ToDouble(textBox44.Text);
                        textBox47.Text = sum.ToString();
                    }
                }
                else if (radioButton4.Checked == true)
                {
                    if ((textBox46.Text != "") && (textBox45.Text != ""))
                    {
                        double sum = Convert.ToDouble(textBox46.Text) / Convert.ToDouble(textBox45.Text);
                        textBox47.Text = formatDouble(sum, 3);
                    }
                    //MessageBox.Show("Введите курс и сумму в рублях");
                }
            }
            catch
            {
                MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в y.e ввиде - 1000,05");
                textBox45.Text = "";
                textBox47.Text = "";
            }
        }

        private void textBox96_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox25.Text == "Агентский")
            {
                tabControl3.Controls.Add(tabPage11);
            }
            else
            {
                tabControl3.Controls.Remove(tabPage11);
            }
        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox20.Text!="")
            {
                int value;
                //dict.
                bool b=dict.TryGetValue(comboBox20.Text,out value);
                string result=GetItem(value);
                label119.Text = result;
            }
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (panel1.Visible != true)
            {
                panel1.Visible = true;
                panel1.Parent = Form1.ActiveForm;
                panel1.BringToFront();
            }
            else
            {
                panel1.Visible = false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            Process.Start("mailto:nevsky2@yandex.ru");
        }

        private void button22_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }


        private void label195_TextChanged(object sender, EventArgs e)
        {
            string m = null;
            string[] strobj;
            Label obj = (Label)sender;
            string result = null;

            Regex r = new Regex("[А-Я]+(-)[0-9]+(-П)");
            m = r.Match(obj.Text).ToString();
            if ((m.ToString() != "") && (m != null))
            {
                strobj = m.Split('-');
                strobj[2] = "О";
                result = string.Join("-", strobj);
                textBox49.Text = result;
                button10.Enabled = true;
                //button2.Enabled = true;
            }
            else
            {
                button10.Enabled = false;
            }
        }
        private SqlConnectionStringBuilder GetConnectSTR()
        {
             SqlConnectionStringBuilder connectStr = new SqlConnectionStringBuilder();
             Section sec = new Section();
             try
             {
                 if (File.Exists("app.ini"))
                 {
                     //SqlConnectionStringBuilder connectStr = new SqlConnectionStringBuilder();
                     connectStr.DataSource = sec.readkey("SQL", "Server", "app.ini");
                     connectStr.UserID = sec.readkey("SQL", "User_ID", "app.ini");
                     connectStr.Password = sec.readkey("SQL", "Pass", "app.ini");
                     connectStr.InitialCatalog = sec.readkey("SQL", "DataBase", "app.ini");
                 }
                 return connectStr;
             }
             catch
             {
                 //MessageBox.Show("Ошибка файла настройки");
                 richTextBox1.AppendText("Ошибка файла настройки\n\r");
                 return null;
             }

        }

        //Transforma
        private void PredDogSave(string id)
        {
            //Section sec = new Section();
            string DogovorNum = textBox7.Text;
            //string id = "";


                SqlConnectionStringBuilder connectStr = GetConnectSTR();
                /* connectStr.DataSource = sec.readkey("SQL", "Server", "app.ini");
                 connectStr.UserID = sec.readkey("SQL", "User_ID", "app.ini");
                 connectStr.Password = sec.readkey("SQL", "Pass", "app.ini");
                 connectStr.InitialCatalog = sec.readkey("SQL", "DataBase", "app.ini");*/
                if (connectStr != null)
                {
                    string[] parametr = new string[2];
                    /*if (checkBox14.Checked==true)
                    {
                        parametr[0]="1";
                    }
                    else
                    {
                        parametr[0]="0";
                    }
                    if (checkBox12.Checked==true)
                    {
                        parametr[1]="1";
                    }
                    else
                    {
                        parametr[1]="0";
                    }*/
                    //transport
                    if (checkBox17.Checked == true)
                    {
                        parametr[0] = checkBox17.Text;
                    }
                    else if (checkBox16.Checked == true)
                    {
                        parametr[0] = checkBox16.Text;
                    }
                    else if (checkBox15.Checked == true)
                    {
                        parametr[0] = checkBox15.Text;
                    }
                    //curremcy
                    if (radioButton1.Checked == true)
                    {
                        parametr[1] = radioButton1.Text;
                    }
                    else if (radioButton2.Checked == true)
                    {
                        parametr[1] = radioButton2.Text;
                    }
                    else if (radioButton3.Checked == true)
                    {
                        parametr[1] = radioButton3.Text;
                    }

                    string query = "INSERT INTO tempdatadogovor VALUES('" + id + "'" + "," + "'" + comboBox9.Text + "'" + "," + "'" + textBox25.Text + "'" + "," + "'" + dateTimePicker1.Text + "'" + "," + "'" + dateTimePicker2.Text + "'" + "," + "'" + textBox21.Text + "'" + "," + "'" + Convert.ToInt32(checkBox14.Checked) + "'" + "," + "'" + Convert.ToInt32(checkBox12.Checked) + "'" + "," + "'" + parametr[0] + "'" + "," + "'" + comboBox16.Text + "'" + "," + "'" + DogovorNum + "'" + "," + "'" + comboBox14.Text + "'" + "," + "'" + maskedTextBox3.Text + "'" + "," + "'" + textBox33.Text + "'" + "," + "'" + textBox107.Text + "'" + "," + "'" + textBox108.Text + "'" + "," + "'" + textBox109.Text + "'" + "," + "'" + textBox122.Text + "'" + "," + "'" + textBox121.Text + "'" + "," + "'" + textBox120.Text + "'" + "," + "'" + textBox119.Text + "'" + "," + "'" + textBox32.Text + "'" + "," + "'" + maskedTextBox4.Text + "'" + "," + "'" + textBox30.Text + "'" + "," + "'" + parametr[1] + "'" + "," + "'" + comboBox13.Text + " " + comboBox12.Text + " " + numericUpDown6.Value.ToString() + " г." + "'" + "," + "'" + comboBox11.Text + " " + comboBox10.Text + " " + numericUpDown5.Value.ToString() + " г." + "'" + ",'" + textBox28.Text + "','" + textBox29.Text + "','" + textBox27.Text + "','" + comboBox29.Text + "')"; //,Birthday,ENpassportnumber,ENpasportStartDate,ENpasportEndDate,phone,email,Adress FROM dbo.Clients_view";

                    SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                    connect.Open();
                    if (connect.State == ConnectionState.Open)
                    {
                        SqlCommand sqlcom = new SqlCommand(query, connect);
                        sqlcom.ExecuteNonQuery();
                        /*//query = "select id, IDENT_CURRENT('Tempdatadogovor') as id1 from tempdatadogovor where dogovornum='" + DogovorNum + "' and manager='" + comboBox14.Text + "' UNION SELECT IDENT_CURRENT('Tempdatadogovor') as id1 ";
                        query = "select id from tempdatadogovor where dogovornum='" + DogovorNum + "' and manager='" + comboBox14.Text + "'";
                        sqlcom=new SqlCommand(query, connect);
                        SqlDataReader r=sqlcom.ExecuteReader();
                        r.Read();
                        id = r["id"].ToString(); ;
                        r.Close(); */                      
                    }
                    /*Datagridsave(dataGridView17, "Location", connect, id);
                    Datagridsave(dataGridView13, "LocationNote", connect, id);
                    Datagridsave(dataGridView12, "Transfer", connect, id);
                    Datagridsave(dataGridView11, "Excurtion", connect, id);
                    Datagridsave(dataGridView10, "Ticket", connect, id);
                    Datagridsave(dataGridView18, "TuristInfo", connect, id);
                    Datagridsave(dataGridView16, "zayvka", connect, id);*/
                    connect.Close();
                }
           /* }
            catch
            {
                MessageBox.Show("Ошибка связи с базой данных");
            }*/
        }
        private void PredDogRead(string DogovorNum)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            object[] TempDogovor=new object[0];
            string query="";string Dogovorid="";
            //try
            //{  //query="select FIO,travelprogram,startdate,enddate,travelroute, GidTranslate, VizaHelp, Transport, medstrach, troublestrach,canselstrach, tyroperator,DogovorNum,manager from tempdataDogovor whrere ";
                if (connectStr != null)
                {
                    query = "select * from tempdataDogovor where id='" + DogovorNum + "'";
                    //SqlConnectionStringBuilder connectStr = GetConnectSTR();

                    SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                    connect.Open();
                    if (connect.State == ConnectionState.Open)
                    {
                        sqlcom = new SqlCommand(query, connect);
                        reader = sqlcom.ExecuteReader();
                        while (reader.Read())
                        {
                            Dogovorid=reader["DInfoKey"].ToString();
                            comboBox6.Text = reader["FIO"].ToString();
                            textBox2.Text = reader["travelprogram"].ToString();
                            
                            maskedTextBox1.Text = reader["birthdate"].ToString();
                            comboBox28.Text = reader["Country"].ToString();
                            dateTimePicker3.Value = DateTime.Parse(reader["startdate"].ToString());
                            dateTimePicker4.Value = DateTime.Parse(reader["enddate"].ToString());
                            textBox6.Text = reader["travelroute"].ToString();
                            checkBox1.Checked = (bool)reader["GidTranslate"];
                            checkBox3.Checked = (bool)reader["VizaHelp"];
                            if (checkBox4.Text == reader["Transport"].ToString())
                            {
                                checkBox4.Checked = true;
                            }
                            else if (checkBox5.Text == reader["Transport"].ToString())
                            {
                                checkBox5.Checked = true;
                            }
                            else if (checkBox6.Text == reader["Transport"].ToString())
                            {
                                checkBox6.Checked = true;
                            }
                           /* dataGridView5.Rows[0].Cells[0].Value = reader["medstrach"];
                            dataGridView5.Rows[0].Cells[1].Value = reader["troublestrach"];
                            dataGridView5.Rows[0].Cells[2].Value = reader["canselstrach"];*/
                            if (reader["tyroperator"].ToString() == "Росинтур")
                            {
                                comboBox3.SelectedItem = reader["tyroperator"].ToString();
                            }
                            else
                            {
                                comboBox3.SelectedItem = reader["tyroperator"].ToString();
                                //string etert = reader["DogovorNum"].ToString().Substring(0, 1);
                                if (reader["DogovorNum"].ToString().Substring(0, 1) != "Ю")
                                {
                                    comboBox4.SelectedItem = "Магазин Путешествий";
                                }
                            }
                            comboBox5.SelectedItem = reader["manager"].ToString();
                            
                            if (reader["dogovornum"].ToString() != null)
                            {
                                string[] strarr=reader["dogovornum"].ToString().Split('-');
                                textBox1.Text = strarr[1];
                            }

                            textBox9.Text = reader["ENpassportseriy"].ToString();
                            textBox110.Text = reader["ENpassportnum"].ToString();
                            textBox111.Text = reader["ENpassportStartDate"].ToString();
                            textBox112.Text = reader["ENpassportOwn"].ToString();
                            textBox117.Text = reader["RUSPassportseriy"].ToString();
                            textBox116.Text = reader["RUSPassportNum"].ToString();
                            textBox115.Text = reader["RUSPassportStartDate"].ToString();
                            textBox114.Text = reader["RUSPassportOwn"].ToString();
                            textBox10.Text = reader["Adress"].ToString();
                            maskedTextBox2.Text = reader["phone"].ToString();
                            textBox12.Text = reader["Email"].ToString();
                            //checkBox67.Checked = (bool)reader["Sms_yes"];
                            //checkBox68.Checked = (bool)reader["Email_yes"];
                            textBox43.Text = reader["RUavans"].ToString();
                            textBox44.Text = reader["ENavans"].ToString();
                            textBox20.Text=reader["Course"].ToString();
                            if (radioButton4.Text == reader["Currency"].ToString())
                            {
                                radioButton4.Checked = true;
                            }
                            else if (radioButton5.Text == reader["Currency"].ToString())
                            {
                                radioButton5.Checked = true;
                            }
                            else if (radioButton6.Text == reader["Currency"].ToString())
                            {
                                radioButton6.Checked = true;
                            }
                        }
                        reader.Close();
                    }
                    DatagridRead(dataGridView14,"Location", Dogovorid, connect);
                    DatagridRead(dataGridView1, "LocationNote", Dogovorid, connect);
                    DatagridRead(dataGridView2, "Transfer", Dogovorid, connect);
                    DatagridRead(dataGridView3, "Excurtion", Dogovorid, connect);
                    DatagridRead(dataGridView4, "Ticket", Dogovorid, connect);
                    DatagridRead(dataGridView5, "Insurance", Dogovorid, connect);
                    DatagridRead(dataGridView6, "TuristInfo", Dogovorid, connect);
                    DatagridRead(dataGridView7, "zayvka", Dogovorid, connect);
                    connect.Close();
                }
                //EventArgs eq=new EventArgs();
                tabControl1_Click(tabControl1, new EventArgs());
           /* }
            catch
            {
                //reader.Close();
            }*/
        }

        private void DatagridRead(DataGridView data, string table, string DogovorNum, SqlConnection connect)
        {
            int i = 0; string query = "";
            data.RowCount = 1; object o = null;
            SqlCommand sqlcom = null; SqlDataReader reader = null;
            if (data.Name == "dataGridView6")
            {
                query = "select id,DInfoKey,FIO,convert(varchar,birthday,105),pasport,PaspBeginDAte,PaspEndDAte from " + table + " where DInfoKey='" + DogovorNum + "'";
            }
            else
            {
                query = "select * from " + table + " where DInfoKey='" + DogovorNum + "'";
            }
            try
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    while (reader.Read())
                    {
                        data.RowCount++;
                        for (int j = 0; j < reader.FieldCount - 2; j++)
                        {
                            o = reader[j + 2];
                            data.Rows[i].Cells[j].Value = reader[j + 2];
                        }
                        i++;
                    }
                    reader.Close();
                }
            }
            catch
            {
                //MessageBox.Show("Ошибка в заполнении таблицы");
                richTextBox1.AppendText("Ошибка в заполнении таблицы при чтении пред_договора\n\r");
            }
            
        }
        private void getDogovorList(string tyroperator, string manager, DataGridView data)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader=null;
            SqlCommand sqlcom = null;
            int count = 0; int first = 0;
            string query = "select id, dogovornum, dogovordate, fio, country, manager from TempDataDogovor";//tyroperator='" + tyroperator+"'";//"' and dogovornum='" + num + "'";// +client.ENpaspOwn + "," + client.RUpaspSeriy + "," + client.RUpaspnum + "," + client.RUpaspDate + "," + client.RUpaspOwn + "," + client.Phone + "," + client.Email;
            if ((comboBox27.Text != "Все")||(comboBox26.Text != "Все"))
            {
                query += " where ";
                if (comboBox27.Text != "Все")
                {
                    query +="manager='"+ comboBox27.Text+"'";
                    first = 1;
                }
                if (comboBox26.Text != "Все")
                {
                    if (first != 1)
                    {
                        query += "country='" + comboBox26.Text + "'";
                        first = 1;
                    }
                    else
                    {
                        query += "and country='" + comboBox26.Text + "'";
                    }
                }
            }
            query += " order by id DESC";
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        data.Rows.Add();
                        data.Rows[count].Cells[0].Value = reader["id"];
                        data.Rows[count].Cells[1].Value = reader["dogovornum"];
                        data.Rows[count].Cells[2].Value = reader["dogovordate"];
                        data.Rows[count].Cells[3].Value = reader["country"];
                        data.Rows[count].Cells[4].Value = reader["manager"];
                        data.Rows[count].Cells[5].Value = reader["fio"];
                        count++;
                        //data.Rows.Add();
                    }
                }
            }
            reader.Close();
            connect.Close();
        }
        private int getrowscount(SqlDataReader sqlr)
        {
            int result=0;
            SqlDataReader s = sqlr;
            if (s.HasRows != false)
            {
                while (s.Read())
                {
                    result++;
                }

            }
            s.Close();
            return result;
        }
        //Transforma
        //DataSQL
        private object GetClientId(Client client)
        {
            //object result = null;
            object id = null;
            //Section sec=new Section();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "select id from Clients where FIO='" + client.FIO + "' and birthdate='" + client.Birthdate + "'";// +client.ENpaspOwn + "," + client.RUpaspSeriy + "," + client.RUpaspnum + "," + client.RUpaspDate + "," + client.RUpaspOwn + "," + client.Phone + "," + client.Email;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
               
               if (reader.HasRows!=false)
                {
                    while (reader.Read())
                    {
                        id = reader["id"];
                    }
                }
                reader.Close();
            }
            connect.Close();
            return id;
                /*if (id != null)
                {
                    bool farg = false;
                    result = id;
                    if (action == 1)
                    {
                        query = "update Clients Set ";
                        if (client.Email != "")
                        {
                            query += "email='" + client.Email + "'";
                            farg = true;
                        }
                        if (client.ENpaspSeriy != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "ENpassportseriy='" + client.ENpaspSeriy + "'";
                        }
                        if (client.ENpaspnum != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "ENpassportnum='" + client.ENpaspnum + "'";
                        }
                        if (client.ENpaspDate != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "ENpassportStartDate='" + client.ENpaspDate + "'";
                        }
                        if (client.ENpaspOwn != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "ENpassportOwn='" + client.ENpaspOwn + "'";
                        }
                        if (client.RUpaspSeriy != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "RUpassportseriy='" + client.RUpaspSeriy + "'";
                        }
                        if (client.RUpaspnum != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "RUpassportnum='" + client.RUpaspnum + "'";
                        }
                        if (client.RUpaspDate != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "RUpassportStartDate='" + client.RUpaspDate + "'";
                        }
                        if (client.RUpaspOwn != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "RUpassportOwn='" + client.RUpaspOwn + "'";
                        }
                        if (client.Phone != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "phone='" + client.Phone + "'";
                        }
                        if (client.ICQ != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "icq='" + client.ICQ + "'";
                        }
                        if (client.Skype != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "skype='" + client.Skype + "'";
                        }
                        if (client.Adress != "")
                        {
                            if (farg == true)
                            {
                                query += ", ";
                            }
                            else
                            {
                                farg = true;
                            }
                            query += "Adress='" + client.Adress + "'";
                        }
                        query += " where id='" + id + "'";
                        sqlcom = new SqlCommand(query, connect);
                        sqlcom.ExecuteNonQuery();
                    }

                }*/
                    //upd_Client
        }
        private object ClientInsert(Client client)
        {
            object result = "";
            //id = GetClientId(client);
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "";
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                    query = "insert into Clients values('" + client.FIO + "', '" + client.Birthdate + "','" + client.ENpaspSeriy + "','" + client.ENpaspnum + "','" + client.ENpaspDate + "','" + "" + "','" + client.ENpaspOwn + "','" + client.RUpaspSeriy + "','" + client.RUpaspnum + "','" + client.RUpaspDate + "','" + client.RUpaspOwn + "','" + client.Phone + "','" + client.Email + "','" + client.ICQ + "','" + client.Skype + "','" + client.Adress  + "','" + client.State_phone+ "')";
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                    query = "select id from Clients where FIO='" + client.FIO + "' and birthdate='" + client.Birthdate+"'";
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    reader.Read();
                    if (reader["id"] != null)
                    {
                        result = reader["id"];
                    }
                    reader.Close();
            }
            connect.Close();
            return result;
        }
        private void ClientUpdate(Client client,string id)
        {
            object result = "";
            //object id = GetClientId(client); 
            bool farg = false;
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            //SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "";
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                query = "update Clients Set ";
                if (client.Email != "")
                {
                    query += "email='" + client.Email + "'";
                    farg = true;
                }
                if (client.ENpaspSeriy != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "ENpassportseriy='" + client.ENpaspSeriy + "'";
                }
                if (client.ENpaspnum != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "ENpassportnum='" + client.ENpaspnum + "'";
                }
                if (client.ENpaspDate != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "ENpassportStartDate='" + client.ENpaspDate + "'";
                }
                if (client.ENpaspOwn != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "ENpassportOwn='" + client.ENpaspOwn + "'";
                }
                if (client.RUpaspSeriy != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "RUpassportseriy='" + client.RUpaspSeriy + "'";
                }
                if (client.RUpaspnum != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "RUpassportnum='" + client.RUpaspnum + "'";
                }
                if (client.RUpaspDate != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "RUpassportStartDate='" + client.RUpaspDate + "'";
                }
                if (client.RUpaspOwn != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "RUpassportOwn='" + client.RUpaspOwn + "'";
                }
                if (client.Phone != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "phone='" + client.Phone + "'";
                }
                if (client.ICQ != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "icq='" + client.ICQ + "'";
                }
                if (client.Skype != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "skype='" + client.Skype + "'";
                }
                if (client.Adress != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "Adress='" + client.Adress + "'";
                }
                if (client.State_phone != "")
                {
                    if (farg == true)
                    {
                        query += ", ";
                    }
                    else
                    {
                        farg = true;
                    }
                    query += "state_phone='" + client.State_phone + "'";
                }
                query += " where id='" + id + "'";
                sqlcom = new SqlCommand(query, connect);
                sqlcom.ExecuteNonQuery();
            }
          
        }
        private string GetDogovorId(string dogovornum, string dogovordate, string dogovormanager, string table)
        {
            string result = null;
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader=null;
            SqlCommand sqlcom = null;
            string query = "select id from "+table+" where Dogovornum='" + dogovornum + "' and DogovorDate='" + makeSQLdate(dogovordate,'.') + "' and Manager='" + dogovormanager + "'";
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader=sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    reader.Read();
                    result = reader["id"].ToString();
                }
            }
            reader.Close();
            connect.Close();
            return result;
        }
        private string DogovorInfoSave(DogovorInfo dinfo, DataGridView dataGr1, DataGridView dataGr2, DataGridView dataGr3, DataGridView dataGr4, DataGridView dataGr5, DataGridView dataGr6, DataGridView dataGr7, DataGridView dataGr8)
            {
            string id = "";
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            //SqlDataReader reader;
            SqlCommand sqlcom = null;
            id = GetDogovorId(dinfo.Dogovornum, dinfo.DogovorDate, dinfo.Manager,"DogovorInfo");
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            if (id == null)
            {
                string query = "insert into DogovorInfo values('" + dinfo.Dogovornum + "','" + makeSQLdate(dinfo.DogovorDate, '.') + "','" + dinfo.clientID + "','" + dinfo.TyrName + "','" + dinfo.TravelPath + "','" + dinfo.StartDate + "','" + dinfo.EndDate + "','" + dinfo.Hotel + "','" + dinfo.PayType + "','" + dinfo.Currency + "','" + dinfo.Course + "','" + dinfo.RUPrice + "','" + dinfo.ENPrice + "','" + dinfo.DogovorType + "','" + dinfo.Manager + "','" + dinfo.Tyroperator + "','" + dinfo.Country + "','" + dinfo.Discount + "','" + dinfo.Sms_yes + "','" + dinfo.Email_yes + "','" + dinfo.CardNum + "')";
                erorrFSave("error.txt", query);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                }
                id = GetDogovorId(dinfo.Dogovornum, dinfo.DogovorDate, dinfo.Manager, "DogovorInfo");
                Datagridsave(dataGr1, "Location", connect, id);
                Datagridsave(dataGr2, "LocationNote", connect, id);
                Datagridsave(dataGr3, "Transfer", connect, id);
                Datagridsave(dataGr4, "Excurtion", connect, id);
                Datagridsave(dataGr5, "Ticket", connect, id);
                DatagridsaveCheck(dataGr6, "Insurance", connect, id);
                Datagridsave(dataGr7, "TuristInfo", connect, id);
                Datagridsave(dataGr8, "zayvka", connect, id);
            }
            else
            {
                string query = "update DogovorInfo Set Dogovornum='" + dinfo.Dogovornum + "', DogovorDate='" + makeSQLdate(dinfo.DogovorDate, '.') + "', Client='" + dinfo.clientID + "', TyrName='" + dinfo.TyrName + "', TravelPath='" + dinfo.TravelPath + "', StartDate='" + dinfo.StartDate + "', EndDate='" + dinfo.EndDate + "', Hotel='" + dinfo.Hotel + "', PayType='" + dinfo.PayType + "',Currency='" + dinfo.Currency + "', Course='" + dinfo.Course + "',RUPrice='" + dinfo.RUPrice + "', ENPrice='" + dinfo.ENPrice + "', DogovorType='" + dinfo.DogovorType + "', Manager='" + dinfo.Manager + "',Tyroperator='" + dinfo.Tyroperator + "', Country='" + dinfo.Country + "', Discount='" + dinfo.Discount + "', Sms_yes='" + dinfo.Sms_yes + "', Email_yes='" + dinfo.Email_yes + "', cardnum='" + dinfo.CardNum + "' where id='" + id + "'"; ;
                erorrFSave("error.txt", query);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                }
            }
            /*query = "select id from DogovorInfo where Dogovorum='" + dinfo.Dogovornum + "' and DogovorDate='" + dinfo.DogovorDate + "' and Manager='" + dinfo.Manager + "'";
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader=sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    reader.Read();
                    id = reader["id"].ToString();
                }
            }*/
            connect.Close();
            return id;
        }
        private string AviaDogovorSave(AviaDogovorInfo ainfo,DataGridView datagr1)
        {
            string id="";
            string query="insert into AviaDogovorInfo values('"+ainfo.Dogovornum+"','"+makeSQLdate(ainfo.DogovorDate,'.')+"','"+ainfo.clientID+"','"+ainfo.Manager+"','"+ainfo.Company+"','"+ainfo.AgentDogNum+"','"+ainfo.AgentDogDate+"','"+ainfo.FIO+"','"+ainfo.Adress+"','"+ainfo.Phone+"','"+ainfo.Country+"','"+ainfo.TravelPath+"')";
            SqlConnectionStringBuilder connectstr = GetConnectSTR();
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectstr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                sqlcom.ExecuteNonQuery();
                id = GetDogovorId(ainfo.Dogovornum, ainfo.DogovorDate, ainfo.Manager, "AviaDogovorInfo");
                Datagridsave(datagr1, "AviaSpecification", connect, id);
            }
            connect.Close();
            return id;
        }
        private string GetFlightInfoId(string dogovorid)
        {
            string result = null;
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader = null;
            SqlCommand sqlcom = null;
            string query = "select id from FlightInfo where DInfoKey='" + dogovorid + "'";
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    reader.Read();
                    result = reader["id"].ToString();
                }
            }
            reader.Close();
            connect.Close();
            return result;
        }
        private void FlightInfoSave(FlightInfo finfo,string manager)
        {
            string id = null;
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            //SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "";
            id = GetFlightInfoId(finfo.DogovorNum);
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (id == null)
            {
                query = "insert into FlightInfo values('" + finfo.DogovorNum + "','" + makeSQLdate(finfo.Date, '.') + "','" + finfo.Time + "','" + finfo.FlightNum + "','" + finfo.StartCity + "','" + finfo.EndCity + "','" + finfo.Mannum + "','" + finfo.Tariff + "','" + finfo.Hotel + "','" + finfo.FIO + "','" + finfo.clientID + "')";
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                }
                //id = GetFlightInfoId(finfo.DogovorNum);
            }
            else
            {
                query = "update FlightInfo SET DInfoKey='" + finfo.DogovorNum + "', Date='" + makeSQLdate(finfo.Date, '.') + "', Time='" + finfo.Time + "', FlightNum='" + finfo.FlightNum + "', StartCity='" + finfo.StartCity + "', EndCity='" + finfo.EndCity + "',Mannum='" + finfo.Mannum + "', Tariff='" + finfo.Mannum + "', Hotel='" + finfo.Hotel + "', FIO='" + finfo.FIO + "', FIOid='" + finfo.clientID + "' where id='"+id+"'";
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                }
            }
            connect.Close();

        }
        //error
        private void erorrFSave(string path, string e)
        {
           if (File.Exists(path))
           {
               using (StreamWriter sw = File.AppendText(path))
               {
                   sw.WriteLine(DateTime.Now.ToString() + e);
               }
           }
           else
           {
               using (StreamWriter sw = File.CreateText(path))
               {
                   sw.WriteLine(DateTime.Now.ToString() + e);
               }
           }
             
        }
        private void AddTextFs(FileStream fs, string value)
        {
            byte[] info = new UTF8Encoding(true).GetBytes(value);
            fs.Write(info, 0, info.Length);
        }
        ///error
        // //DataSQL End
        private void Datagridsave(DataGridView data, string table, SqlConnection connect, string  DIKey)
        {
            SqlCommand sqlcom;
            string str = "";
            for (int i = 0; i < data.RowCount-1; i++)               
            {
                str = "insert into " + table + " values('" + DIKey + "'";
                for (int j = 0; j < data.ColumnCount; j++)
                {
                   /* if (j == 0)
                    {
                        str += "'" + data.Rows[i].Cells[j].Value + "'";
                    }
                    else
                    {*///if (data)
                    if ((data.Name == "dataGridView6") || (data.Name == "dataGridView18"))
                    {
                        if (j == 1)
                        {
                            if (data.Rows[i].Cells[j].Value != null)
                            {
                                //char sym=data.Rows[i].Cells[j].Value.ToString()[2];
                                str += ", " + "'" + makeSQLdate(data.Rows[i].Cells[j].Value.ToString(), data.Rows[i].Cells[j].Value.ToString()[2]) + "'";
                            }
                            else
                            {
                                str += ", " + "'" + data.Rows[i].Cells[j].Value + "'";
                            }
                        }
                        else
                        {
                            str += ", " + "'" + data.Rows[i].Cells[j].Value + "'";
                        }
                    }
                    else
                    {
                        str += ", " + "'" + data.Rows[i].Cells[j].Value + "'";
                    }
                   // }
                }
                str += ")"; 
               // insertstr(str);
                try
                {
                    //connect.Open();
                    if (connect.State != ConnectionState.Open)
                    {
                        connect.Open();
                    }
                    if (connect.State == ConnectionState.Open)
                    {
                        sqlcom = new SqlCommand(str, connect);
                        sqlcom.ExecuteNonQuery();
                    }
                    
                }
                catch
                {
                    //MessageBox.Show("Ошибка вставки блока данных");
                    richTextBox1.AppendText("Ошибка вставки блока данных при сохранении грида\n\r");
                }
            }
        }
        private void DatagridsaveCheck(DataGridView data, string table, SqlConnection connect, string DIKey)
        {
            SqlCommand sqlcom;
            string str = "";
            for (int i = 0; i < data.RowCount - 1; i++)
            {
                str = "insert into " + table + " values('" + DIKey + "'";
                for (int j = 0; j < data.ColumnCount; j++)
                {
                    /* if (j == 0)
                     {
                         str += "'" + data.Rows[i].Cells[j].Value + "'";
                     }
                     else
                     {*/
                    //if (data)
                    str += ", " + "'" + Convert.ToInt32(data.Rows[i].Cells[j].Value) + "'";
                    // }
                }
                str += ")";
                // insertstr(str);
                try
                {
                    //connect.Open();
                    if (connect.State != ConnectionState.Open)
                    {
                        connect.Open();
                    }
                    if (connect.State == ConnectionState.Open)
                    {
                        sqlcom = new SqlCommand(str, connect);
                        sqlcom.ExecuteNonQuery();
                    }

                }
                catch
                {
                    //MessageBox.Show("Ошибка вставки блока данных");
                    richTextBox1.AppendText("Ошибка вставки блока данных при сохранении грида check\n\r");
                }
            }
        }
        private bool getbooldatagrid(DataGridView data, int row, int col)
        {
            bool result;
            if (data.Rows[row].Cells[col].Value != null)
            {
                result = (bool)data.Rows[row].Cells[col].Value;
            }
            else
            {
                result = false;
            }
            return result;
        }
        private void insertstring(string query)
        {

        }

        private void button23_Click(object sender, EventArgs e)
        {
           /* SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
           // DatagridRead("Р-16-П", connect);*/
            //PredDogRead("Р-22-П");
            if (panel2.Visible == false)
            {
                panel2.Visible = true;
                panel2.Parent = Form1.ActiveForm;
                panel2.BringToFront();
            }
            else
            {
                panel2.Visible = false;
                panel2.SendToBack();
            }
        }

        /*private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox26.Text == "Росинтур")
            {
                //textBox1.Text = "Р-";
                label199.Text = "Р-";
            }
            if  (comboBox26.Text == "РосинтурЮг")
            {
                //textBox1.Text = "Ю-";
                label199.Text = "Ю-" ;
            }
            if (comboBox26.Text == "Магазин Путешествий")
            {
                //textBox1.Text = "МП-";
                label199.Text = "МП-";
                if (textBox113.Text != "")
                {
                    System.Event(this.textBox113_TextChanged);
                }
            }
        }*/

        private void textBox113_TextChanged(object sender, EventArgs e)
        {
            //TextBox t=(TextBox)sender;

           /* if (comboBox26.Text == "Росинтур")
            {
                //textBox1.Text = "Р-";
                label199.Text = "Р-";
            }
            if (comboBox26.Text == "РосинтурЮг")
            {
                //textBox1.Text = "Ю-";
                label199.Text = "Ю-";
            }
            if (comboBox26.Text == "Магазин Путешествий")
            {
                //textBox1.Text = "МП-";
                label199.Text = "МП-";

            }
            label199.Text +=textBox113.Text+"-П";
            label199.Visible = true;
            button24.Enabled = true;*/
        }

        private void button24_Click(object sender, EventArgs e)
        {
            /*try
            {
                PredDogRead(label199.Text);
            }
            catch
            {
                MessageBox.Show("Ошибки при загрузке договора");
            }
            panel2.Visible = false;*/
          //  try
           // {
                // Size sP = new Size(panel2.Size.Width, 540);
                //Size sD = new Size(dataGridView21.Size.Width, 350);
                //panel2.Size=sP;
                //dataGridView21.Size=sD;
                dataGridView21.Visible = true;
                dataGridView21.Rows.Clear();
                //try
                //{
                    getDogovorList(comboBox26.Text, comboBox27.Text, dataGridView21);
                /*}
                catch
                {

                }*/

           /* }
            catch
            {
                MessageBox.Show("Ошибки при загрузке договора");
            }*/
            DataGridViewCellEventArgs earg=new DataGridViewCellEventArgs(0,0);
            //earg.ColumnIndex = 0; earg.RowIndex = 0;

            try
            {
                dataGridView7_CellEndEdit(dataGridView7, earg);
            }
            catch
            {
                MessageBox.Show("Ошибка при расчете заявки");
            }
        }

        private void вылетыТуристовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel3.Parent = Form1.ActiveForm;
            panel3.BringToFront();

            //string query = "select f.dogovrnum,f.Country, f.date,f.time,f.flightnum,f.startcity,f.endcity, f.returndate,f.hotel,f.fio,f.mannum, f.fioid, c.fio, c.ENpassportseriy, c.ENpassportnum,c.ENpassportStartDate,d.manager from FlightInfo as f, Clients as c, dogovorinfo as d where c.id=f.FIOid and d.id=f.dogovornum and f.date='"+dateTimePicker7.Text;
        }

        private void button27_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
        }
        private void makeFlightInformation(string query)
        {
            object[] ExcelParam = new object[1];
            Section sec = new Section();
            string path1 = sec.readkey("SavePath", "ClientPath", "app.ini") + @"\Список вылетов туристов\";

            Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
            object Excel = Activator.CreateInstance(obj_excel);

            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
            ExcelParam[0] = Basepath + @"Template\TouristList.xls";
            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
            ExcelParam[0] = 1;
            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
            /*object[] ExcelParamC = new object[2] { Missing.Value, obj_worksheet };
            obj_worksheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, obj_worksheet, ExcelParamC);
            if (File.Exists(path1 + @"\Список вылетов туристов.xls"))
            {
                ExcelParam[0] = Basepath + @"Template\TouristList.xls";
                object obj_workbook_t = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
                object obj_worksheets_t = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook_t, null);
                object shcount = obj_worksheets.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, obj_worksheets_t, null);
                ExcelParam[0] = shcount;
                object obj_worksheet_t = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets_t, ExcelParam);

                obj_worksheet_t.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, obj_worksheet_t, null);
            }
           
            //ExcelParam[0] = 1;
           // object[] ExcelParamC = new object[2] { obj_worksheet, Missing.Value };
            //obj_worksheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, obj_worksheet, ExcelParamC);
            //obj_worksheet.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, obj_worksheet, null);*/
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader; int mannum = 0;
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                int count = 2;
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        
                        SetCellData(reader["date"].ToString(), "A" + count.ToString(), obj_worksheet);
                        SetCellData(reader["Country"].ToString(), "B" + count.ToString(), obj_worksheet);
                        SetCellData(reader["time"].ToString(), "C" + count.ToString(), obj_worksheet);
                        SetCellData(reader["flightnum"].ToString(), "D" + count.ToString(), obj_worksheet);
                        SetCellData(reader["startcity"].ToString(), "E" + count.ToString(), obj_worksheet);
                        SetCellData(reader["endcity"].ToString(), "F" + count.ToString(), obj_worksheet);
                        SetCellData(reader["enddate"].ToString(), "G" + count.ToString(), obj_worksheet);
                        SetCellData(reader["hotel"].ToString(), "H" + count.ToString(), obj_worksheet);
                        if ((reader["fioid"].ToString() != null) && (reader["fioid"].ToString() != ""))
                        {
                            //if (Convert.ToInt32(reader["mannum"]) > 1)
                            //{
                                mannum = Convert.ToInt32(reader["mannum"]) - 1;
                            /*}
                            else
                            {
                                mannum = Convert.ToInt32(reader["mannum"]);
                            }*/
                            SetCellData(reader["FIO"].ToString() + "+" + mannum, "I" + count.ToString(), obj_worksheet);
                            SetCellData(reader["phone"].ToString(), "J" + count.ToString(), obj_worksheet);
                            SetCellData(reader["Tyroperator"].ToString(), "K" + count.ToString(), obj_worksheet);
                            //SetCellData(reader["ENpassportseriy"].ToString() + " " + reader["ENpassportnum"].ToString(), "J" + count.ToString(), obj_worksheet);
                        //SetCellData(reader["c.ENpassportnum"], "A" + count.ToString(), obj_worksheet);
                            //SetCellData(reader["ENpassportStartDate"].ToString(), "K" + count.ToString(), obj_worksheet);
                        }
                        SetCellData(reader["fio"].ToString() + "+" + reader["mannum"].ToString(), "I" + count.ToString(), obj_worksheet);
                        SetCellData(reader["manager"].ToString(), "L" + count.ToString(), obj_worksheet);
                        count++; mannum = 0;
                    }
                }
                reader.Close();
            }
            connect.Close();
            //ExcelParam[0] = 1;
            
            
                if (Directory.Exists(@"c:\Список вылетов туристов"))
                {
                    if (File.Exists(@"c:\Список вылетов туристов\Список вылетов туристов.xls"))
                    {
                        File.Delete(@"c:\Список вылетов туристов\Список вылетов туристов.xls");
                    }
                    ExcelParam[0] = @"c:\Список вылетов туристов\Список вылетов туристов";
                    obj_workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, obj_workbook, ExcelParam);
                }
                else
                {
                    Directory.CreateDirectory(@"c:\Список вылетов туристов");
                    ExcelParam[0] = @"c:\Список вылетов туристов\Список вылетов туристов";
                    obj_workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, obj_workbook, ExcelParam);
                }
                
           
            ExcelParam[0] = 1;
            Excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, Excel, ExcelParam);
            //clean
            Marshal.ReleaseComObject(obj_worksheet);
            Marshal.ReleaseComObject(obj_worksheets);
            Marshal.ReleaseComObject(obj_workbook);
            Marshal.ReleaseComObject(obj_workbooks);
            //Marshal.ReleaseComObject(Excel);
           
        }

        private void button25_Click(object sender, EventArgs e)
        {

            string query = "select f.DInfoKey, convert(varchar(15),f.date,104) as date,f.time,f.flightnum,f.startcity,f.endcity, f.hotel,f.fio,f.mannum, f.fioid, c.fio, c.phone,d.manager,d.EndDate, d.Country, d.Tyroperator from FlightInfo as f, Clients as c, dogovorinfo as d where c.id=f.FIOid and d.id=f.DInfoKey and f.date='" + makeSQLdate(dateTimePicker7.Text, '.') + "' order by  f.date,d.country";
            makeFlightInformation(query);

        }
        private string makeSQLdate(string str, char sep)
        {
            string result;
            string[] temp = str.Split(sep);
            string[] date = new string[temp.Length];
            for (int i = 0; i < temp.Length; i++)
            {
                date[i] = temp[temp.Length - i - 1];
            }
            result = string.Join(".", date);
            return result;
        }
        private void button26_Click(object sender, EventArgs e)
        {
            string query = "select f.DInfoKey, convert(varchar(15),f.date,104) as date,f.time,f.flightnum,f.startcity,f.endcity, f.hotel,f.fio,f.mannum, f.fioid, c.fio, c.phone,d.EndDate,d.manager,d.Country, d.Tyroperator from FlightInfo as f, Clients as c, dogovorinfo as d where c.id=f.FIOid and d.id=f.DInfoKey and f.date>='" + makeSQLdate(dateTimePicker8.Text, '.') + "' and f.date<='" + makeSQLdate(dateTimePicker9.Text, '.') + "' order by f.date, d.country";
            makeFlightInformation(query);

        }
        private void button28_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            //dataGridView21.Visible = false;
        }

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            button24.Enabled = true;
        }

        private void dataGridView21_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView data = (DataGridView)sender;
            PredDogRead(data.Rows[e.RowIndex].Cells[0].Value.ToString());
            //data.Visible = false;daa
            DataGridViewCellEventArgs earg=new DataGridViewCellEventArgs(1,0);
            dataGridView7_CellEndEdit(dataGridView7, earg);
            panel2.Visible = false;
        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            tracechekedstate(checkBox21, checkBox32);
        }

        private void checkBox32_CheckedChanged(object sender, EventArgs e)
        {
            tracechekedstate(checkBox32, checkBox21);
        }
        private void tracechekedstate(CheckBox chsource,CheckBox chtarget)
        {
            if (chsource.Checked == true)
            {
                chtarget.Checked = false;

            }
            else
            {
                chtarget.Checked = true;
            }
        }

        private void checkBox33_CheckedChanged(object sender, EventArgs e)
        {
            tracechekedstate(checkBox33, checkBox34);
        }

        private void checkBox34_CheckedChanged(object sender, EventArgs e)
        {
            tracechekedstate(checkBox34, checkBox33);
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*object[] str;
            ComboBox tempC = (ComboBox)sender;
            //this.comboBox6.Items.Clear();
            //  string strline = this.comboBox6.SelectedItem.ToString();
            //str = GetClientData(tempC.Text.ToString());
            textBox33.Text = str[0].ToString();
            textBox107.Text = str[1].ToString();
            textBox108.Text = str[2].ToString();
            textBox109.Text = str[3].ToString();
            textBox122.Text = str[4].ToString();
            textBox121.Text = str[5].ToString();
            textBox120.Text = str[6].ToString();
            textBox119.Text = str[7].ToString();
            /*dataGridView6.Rows[0].Cells[0].Value = this.comboBox6.Text;
            dataGridView6.Rows[0].Cells[1].Value = str[4];
            dataGridView6.Rows[0].Cells[2].Value = str[0];
            dataGridView6.Rows[0].Cells[3].Value = str[5];
            dataGridView6.Rows[0].Cells[4].Value = str[6];*/
            /*textBox123.Text = str[8].ToString();
            textBox32.Text = str[9].ToString();
            textBox31.Text = str[10].ToString();
            textBox30.Text = str[11].ToString();*/
        }
        private void PaspCH(TextBox tempT,int symcount, int shift)
        {
            
            if (tempT.Text.Length == symcount)
            {
                Control c = tempT.Parent;
                foreach (Control t in c.Controls)
                {
                    if (t.TabIndex == tempT.TabIndex + shift)
                    {
                        t.Select();
                       
                    }
                }
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 2, 2);
        }

        private void textBox110_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 7, 2);
        }

        private void textBox117_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 4, 2);
        }

        private void textBox116_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 6, 2);
        }
        private void dateCH(TextBox t)
        {
            string str = "";
            if ((t.TextLength == 3) || (t.TextLength == 6))
            {
                if ((t.Text[t.TextLength - 1] == '0') || (t.Text[t.TextLength - 1] == '1') || (t.Text[t.TextLength - 1] == '2') || (t.Text[t.TextLength - 1] == '3') || (t.Text[t.TextLength - 1] == '4') || (t.Text[t.TextLength - 1] == '5') || (t.Text[t.TextLength - 1] == '6') || (t.Text[t.TextLength - 1] == '7') || (t.Text[t.TextLength - 1] == '8') || (t.Text[t.TextLength - 1] == '9'))
                {
                    str = t.Text.Substring(0, t.TextLength - 1) + '.' + t.Text[t.TextLength - 1];
                }
                else
                {
                    str = t.Text.Substring(0, t.TextLength - 1) + '.';
                }


                t.Text = str;
                t.SelectionStart = t.TextLength;
            }
        }
        private void textBox111_TextChanged(object sender, EventArgs e)
        {
            TextBox t=(TextBox)sender;
            dateCH(t);
            PaspCH(t, 10, 2);
            
        }

        private void textBox118_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            dateCH(t);
           // PaspCH(t, 10, 2);
        }

        private void textBox123_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            dateCH(t);
        }

        private void textBox108_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            dateCH(t);
            PaspCH(t, 10, 2);
        }

        private void textBox120_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            dateCH(t);
            PaspCH(t, 10, 2);
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 2, 2);
        }

        private void textBox107_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 7, 2);
        }

        private void textBox122_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 4, 2);
        }

        private void textBox121_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 6, 2);
        }
        //Konsul pasp-dateCH
        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            dateCH(t);
            PaspCH(t, 10, 2);
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 2, 1);
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            PaspCH((TextBox)sender, 7, 2);
        }

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            dateCH(t);
            PaspCH(t, 10, 2);
        }

        private void button29_Click(object sender, EventArgs e)
        {

            try
            {
                if ((textBox20.Text != "") && (textBox43.Text != ""))
                {
                    double sum = Convert.ToDouble(textBox43.Text) / Convert.ToDouble(textBox20.Text);
                    textBox44.Text = formatDouble(sum,3);
                }
                else
                {
                    MessageBox.Show("Введите курс и сумму в рублях");
                }
            }
            catch
            {
                MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в рублях ввиде - 1000,05");
                textBox15.Text = "";
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {

            try
            {
                if ((textBox44.Text != "") && (textBox20.Text != ""))
                {
                    double sum = Convert.ToDouble(textBox44.Text) * Convert.ToDouble(textBox20.Text);
                    textBox43.Text =  Convert.ToInt32(sum).ToString();
                }
                else
                {
                    MessageBox.Show("Введите курс и сумму в рублях");
                }
            }
            catch
            {
                MessageBox.Show("Введите курс ввиде - 29,93 и \r\n сумму в рублях ввиде - 1000,05");
                textBox15.Text = "";
            }
        }

        private void менеджеровToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            panel4.BringToFront();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            dataGridView22.Visible = true;
            //getData
            string query = "select dogovorInfo.id, dogovorInfo.dogovornum, dogovorInfo.currency, convert(varchar(20),dogovorInfo.dogovordate,105) as dogovordate,dogovorInfo.Tyroperator, clients.FIO, clients.ENpassportseriy, clients.ENpassportnum, clients.Birthdate, dogovorInfo.StartDate,dogovorInfo.EndDate, dogovorInfo.Country,  dogovorInfo.Hotel, dogovorInfo.RUPrice,dogovorInfo.Discount, dogovorInfo.ENPrice, dogovorInfo.Manager, dogovorInfo.Course from dogovorInfo,clients where dogovordate>='" + makeSQLdate(dateTimePicker11.Text, '.') + "' and dogovordate<='" + makeSQLdate(dateTimePicker10.Text, '.') + "' and clients.id=dogovorInfo.client" + " and dogovorInfo.DogovorType='Основной'"; 
            if (comboBox30.Text != "Все")
            {
                query += " and manager='" + comboBox30.Text + "'";
            }

                   
            try
            {
                getmanager_reportInfo(query);
            }
            catch
            {

            }
            //getData
        }

        private void button31_Click(object sender, EventArgs e)
        {

        }
        private void getmanager_reportInfo(string query)
        {
            //ArrayList idslist = new ArrayList();
            string tquery = ""; dataGridView22.RowCount = 1;
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader; SqlDataReader treader;
            SqlCommand sqlcom = null; double enDiscount=0;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
               
                int count = -1;
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        count++;
                        dataGridView22.Rows.Add();
                        dataGridView22.Rows[count].Cells[0].Value = count + 1;
                        dataGridView22.Rows[count].Cells[1].Value = reader["manager"].ToString();
                        dataGridView22.Rows[count].Cells[2].Value = reader["dogovornum"].ToString();
                        dataGridView22.Rows[count].Cells[3].Value = reader["dogovordate"].ToString();
                        dataGridView22.Rows[count].Cells[4].Value = reader["Tyroperator"].ToString();
                        dataGridView22.Rows[count].Cells[5].Value = reader["FIO"].ToString();
                        //dataGridView22.Rows[count].Cells[5].Value = reader["###"].ToString();
                        GetturistToDogovor((DataGridViewComboBoxCell)dataGridView22.Rows[count].Cells[6], reader["id"].ToString());
                        /*tquery = "select id,DInfoKey, fio from turistinfo where  DInfoKey='" + reader["id"].ToString() + "'";
                         sqlcom = new SqlCommand(tquery, connect);
                         treader = sqlcom.ExecuteReader();
                         count = 0;
                         if (treader.HasRows != false)
                         {
                             while (treader.Read())
                             {
                                 DataGridViewComboBoxCell tdata = (DataGridViewComboBoxCell)dataGridView22.Rows[count].Cells[5];
                                 tdata.Items.AddRange(treader["FIO"].ToString());
                             }
                         }*/
                        dataGridView22.Rows[count].Cells[7].Value = reader["ENpassportseriy"].ToString() + reader["ENpassportnum"].ToString();
                        dataGridView22.Rows[count].Cells[8].Value = reader["Birthdate"].ToString();
                        dataGridView22.Rows[count].Cells[9].Value = reader["StartDate"].ToString() +" - " + reader["EndDate"].ToString();
                        dataGridView22.Rows[count].Cells[10].Value = reader["Country"].ToString();
                        //12
                        dataGridView22.Rows[count].Cells[12].Value = reader["Hotel"].ToString();
                        if (reader["currency"].ToString() == "RUR")
                        {
                            dataGridView22.Rows[count].Cells[13].Value = reader["RUPrice"].ToString();
                            dataGridView22.Rows[count].Cells[16].Value = reader["Discount"].ToString();
                        }
                        else
                        {
                            dataGridView22.Rows[count].Cells[13].Value = reader["ENPrice"].ToString();
                            if ((reader["Discount"] != null) && (reader["Course"] != null) && (reader["Course"].ToString() != "") && (reader["Discount"].ToString() != ""))
                            {
                            enDiscount=(Convert.ToDouble(reader["Discount"])) * (Convert.ToDouble(reader["Course"]));
                            dataGridView22.Rows[count].Cells[16].Value = enDiscount.ToString();
                            }
                        }
                        dataGridView22.Rows[count].Cells[14].Value = reader["Course"].ToString();
                        
                        //dataGridView22.Rows[count].Cells[14].Value = reader["ruDiscount"].ToString();
                        //idslist.Add((object)reader["id"]);
                        //treader.Close();
                    }
                }
                reader.Close();
                //int razm = idslist.Count;
               /* int[] ids = new int[idslist.Count]; idslist.CopyTo(ids);
                string tquery = "select id,DInfoKey fio from turistinfo where ";
                for (int i=0; i < ids.Length;i++ )
                {
                    if (i == ids.Length - 1)
                    {
                        tquery += " DInfoKey='"+ids.GetValue(i).ToString()+"'";
                    }
                    else
                    {
                        tquery +="DInfoKey='"+ids.GetValue(i).ToString()+"' or " ;
                    }
                }
                sqlcom = new SqlCommand(tquery, connect);
                reader = sqlcom.ExecuteReader();
                
                count = 0;
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        if (dataGridView22.Rows[count].Cells[1].Value.ToString() == reader["dogovornum"].ToString())
                        {
                            DataGridViewComboBoxCell tdata = (DataGridViewComboBoxCell)dataGridView22.Rows[count].Cells[5];
                            tdata.Items.AddRange(reader["FIO"].ToString());
                        }
                    }
                }*/
                //reader.Close();
            }
            connect.Close();
        }
        private void GetturistFromDemand(DataGridViewComboBoxCell data, string id)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "select id, fio from Agent_demand_turist where  DemandKey='" + id + "'";
            connect.Open();
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                        /// DataGridViewComboBoxCell tdata = data;
                        data.Items.Add(reader["FIO"].ToString()); data.Style.NullValue = reader["FIO"].ToString();
                        //data.Selected = true;
                    }
                }
            }
           //DataGridViewComboBoxColumn c = dataGridView31.Rows[count].Cells[3];
            //DataGridViewComboBox 
            connect.Close();
        }
        private void GetturistToDogovor(DataGridViewComboBoxCell data,string id)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR(); 
            string query = "";
            SqlDataReader reader; 
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "select id,DInfoKey, fio from turistinfo where  DInfoKey='" + id + "'";
            connect.Open();
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    while (reader.Read())
                    {
                       /// DataGridViewComboBoxCell tdata = data;
                        data.Items.Add(reader["FIO"].ToString()); data.Style.NullValue = reader["FIO"].ToString();
                        //data.Selected = true;
                    }
                }
            }
            connect.Close();
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            this.textBox36.Text = this.dateTimePicker1.Text + "-" + this.dateTimePicker2.Text;
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            this.textBox18.Text = this.dateTimePicker3.Text + "-" + this.dateTimePicker4.Text;
        }

        private void button31_Click_1(object sender, EventArgs e)
        {
            try
            {
                make_manage_report_ex();
            }
            catch
            {

            }
        }
        private void make_manage_report_ex()
        {
            object[] ExcelParam = new object[1];

            Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
            object Excel = Activator.CreateInstance(obj_excel);

            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
            ExcelParam[0] = Basepath + @"Template\Manager_report.xls";
            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
            ExcelParam[0] = 1;
            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
            DataGridViewComboBoxCell data = new DataGridViewComboBoxCell();
            DataGridViewComboBoxCell.ObjectCollection obj; string turists = "";
            for (int i=0;i<dataGridView22.RowCount;i++)
            {
                turists = "";
                //for (int j = 0; j < dataGridView22.ColumnCount; j++)
                //{
                if (dataGridView22.Rows[i].Cells[0].Value != null) { SetCellData(dataGridView22.Rows[i].Cells[0].Value.ToString(), "A" + (2 + i), obj_worksheet); }
                    if (dataGridView22.Rows[i].Cells[1].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[1].Value.ToString(), "B" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[2].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[2].Value.ToString(), "C" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[3].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[3].Value.ToString(), "D" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[4].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[4].Value.ToString(), "E" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[5].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[5].Value.ToString(), "F" + (2 + i), obj_worksheet);}
                    data = (DataGridViewComboBoxCell)dataGridView22.Rows[i].Cells[6];
                    if (data.Items.Count != 0) 
                    {
                        
                        obj = data.Items;
                        for (int j = 0; j < obj.Count; j++)
                        {
                            turists += obj[j].ToString()+"; ";
                        }
                        SetCellData(turists, "G" + (2 + i), obj_worksheet);
                    }
                    if (dataGridView22.Rows[i].Cells[7].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[7].Value.ToString(), "H" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[8].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[8].Value.ToString(), "I" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[9].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[9].Value.ToString(), "J" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[10].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[10].Value.ToString(), "K" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[11].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[11].Value.ToString(), "L" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[12].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[12].Value.ToString(), "M" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[13].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[13].Value.ToString(), "N" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[14].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[14].Value.ToString(), "O" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[15].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[15].Value.ToString(), "P" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[16].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[16].Value.ToString(), "Q" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[17].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[17].Value.ToString(), "R" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[18].Value != null) {SetCellData(dataGridView22.Rows[i].Cells[18].Value.ToString(), "S" + (2 + i), obj_worksheet);}
                    if (dataGridView22.Rows[i].Cells[19].Value != null) { SetCellData(dataGridView22.Rows[i].Cells[19].Value.ToString(), "T" + (2 + i), obj_worksheet); }
                //}
            }
            ExcelParam[0] = "True";
            Excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, Excel, ExcelParam);
            //save
            string managername = "";
            Section sec = new Section();
            string path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini"));
            if (comboBox30.Text != "Все")
            {
                managername = comboBox30.Text;
                //managername = managername.Replace('.', '_');
            }
           /* if ((path != null) && (Directory.Exists(path)))
            {
                if (managername != "")
                {
                    path = Path.GetFullPath(sec.readkey("SavePath", "ClientPath", "app.ini") + "\\" + managername + "\\Отчеты");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                }*/
            if (path != null) 
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            if ((path != null)&&(Directory.Exists(path)))
            {
                if (managername != "")
                {
                    path += "\\" + managername + "\\Отчеты" + "\\" + DateTime.Now.Year.ToString();
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                }
                /*if (checkBox11.Checked == true)
                {
                    ExcelParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Заявка " + num + "(проект)", ".xls");

                    WordParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Заявка " + num + "(проект)", ".doc");
                }
                else
                {
                    ExcelParam[0] = CheckFileName(path + "\\"+ "(" + client + ")"+"Заявка " + num+"("+DateTime.Now.Day+" " +((Month)DateTime.Now.Month).ToString()+" "+DateTime.Now.Year+")", ".xls");

                    WordParam[0] = CheckFileName(path + "\\" + "(" + client + ")" + "Договор " + num + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".doc");
                }*/
                managername = managername.Replace('.', '_');
                ExcelParam[0] = CheckFileName(path + "\\Отчет по продажам " +managername + "(" + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + ")", ".xls");
                obj_workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, obj_workbook, ExcelParam);
                
            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\КлиентскиеДоговора"))
                {
                    Directory.CreateDirectory("c:\\КлиентскиеДоговора");
                    localpath = "c:\\КлиентскиеДоговора";
                }
                else
                {
                    localpath = "c:\\КлиентскиеДоговора";
                }

                //object Doc = arguments.Doc;
                //object WordApp = arguments.App;;
                ExcelParam[0] = localpath + "\\Отчет по продажам " + managername + "(" + DateTime.Now.ToLongDateString() + ")";
                obj_workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, obj_workbook, ExcelParam);

            }
            //
            //clean excel

            Marshal.ReleaseComObject(obj_workbooks);
            Marshal.ReleaseComObject(obj_workbook);
            Marshal.ReleaseComObject(obj_worksheet);
            Marshal.ReleaseComObject(obj_worksheets);
            //SetCellData(
        }

        private void dataGridView24_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridView data = (DataGridView)sender; Point p_cell;
            if (e.ColumnIndex == 1)
            {
                DataGridView data = (DataGridView)sender; Point p_cell;
                databox.Set(sender, e);
                Rectangle rect = new Rectangle();
                p_cell = getCellLocation(data, e.RowIndex, e.ColumnIndex);
                //DataGridView.HitTestInfo hit = data.HitTest( //data.Rows[e.RowIndex].Cells[e.ColumnIndex].
                //data.Container.
                rect.X = data.Location.X + groupBox17.Location.X + tabControl4.Location.X + p_cell.X;
                rect.Y = data.Location.Y + groupBox17.Location.Y + tabControl4.Location.Y + p_cell.Y;
                //data.PointToScreen(p);
                this.aviaCPanel.Location = rect.Location;
                this.aviaCPanel.BringToFront();
                this.aviaCPanel.Show();
            }
            if (e.ColumnIndex == 4)
            {
                DataGridView d1 = (DataGridView)sender;  Point p_cell;
               // databox.Set(sender, e);
                //Rectangle rect = new Rectangle();
                p_cell = getCellLocation(d1, e.RowIndex, e.ColumnIndex);
                databox.Set(sender, e);
                DataGridViewCell c1 = d1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                Point point = d1.Location;
                Rectangle rect = new Rectangle();
                rect.X = point.X + p_cell.X + c1.Size.Width / 2;
                rect.Y = point.Y +p_cell.Y+ c1.Size.Height / 2;
                this.calPanel.Location = rect.Location;
                //p1.Location.Y = rect.Y + 10;


                this.calPanel.Show();
            }
        }

        private void dataGridView24_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
            
        }
        private Point getCellLocation(DataGridView data, int row, int col)
        {
            Point p=new Point();
            p.X += data.ColumnHeadersHeight;
            p.Y += data.RowHeadersWidth;
            for (int i=0; i<row;i++)
            {
                p.Y += data.Rows[i].Height;
            }
            for (int i = 0; i < col; i++)
            {
                p.X += data.Columns[i].Width;
            }
            return p;
        }
        //avia_dog
        private void button40_Click(object sender, EventArgs e)
        {
            object obj_App;
            object obj_Doc;
            object obj_Bookmarks;
            //object obj_Bookmark;
            //object obj_Selection;
            //object obj_Range;
            object obj_Tables;
            object[] Param;
            string currency="";
            button40.Enabled = false;
            button35.Visible = true;
            Param = new object[1];
            Type obj_Class = Type.GetTypeFromProgID("Word.Application");
            object Word = Activator.CreateInstance(obj_Class);

            obj_App = Word.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, Word, null);
            obj_Doc = obj_App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj_App, null);
            /*if (comboBox25.Text == "Клиентский")
            {
                Param[0] = Basepath + @"Template\KonsulDog.doc";
            }
            else if (comboBox25.Text == "Агентский")
            {*/
                Param[0] = Basepath + @"Template\AviaDogovor.doc";
            //}
            object Doc = obj_Doc.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_Doc, Param);
            obj_Bookmarks = Doc.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, Doc, null);
            SetBookMarkText("DayNow", obj_Bookmarks, obj_App, this.comboBox34.SelectedItem.ToString());
            SetBookMarkText("MonthNow", obj_Bookmarks, obj_App, this.comboBox33.SelectedItem.ToString());
            SetBookMarkText("YearNow", obj_Bookmarks, obj_App, this.numericUpDown7.Value.ToString());
            SetBookMarkText("number", obj_Bookmarks, obj_App, this.textBox124.Text);
            SetBookMarkText("Company", obj_Bookmarks, obj_App, this.comboBox35.Text);
            SetBookMarkText("AgentDog", obj_Bookmarks, obj_App, this.textBox126.Text);
            SetBookMarkText("AgentDogDate", obj_Bookmarks, obj_App, this.textBox127.Text);
            SetBookMarkText("FIO", obj_Bookmarks, obj_App, this.comboBox32.Text);
            
            obj_Tables = Doc.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, Doc, null);
            if (dataGridView24.RowCount > 3)
            {
                TableSize(dataGridView24, obj_Tables, 1, 3);
            }
            TableProcess(dataGridView24, obj_Tables, 1);
            double Sum = 0;
            for (int i = 0; i < dataGridView24.RowCount; i++)
            {
                if (dataGridView24.Rows[i].Cells[2].Value!=null)
                {
                    Sum += Convert.ToInt32(dataGridView24.Rows[i].Cells[2].Value);
                }
            }
            double rubsum = Sum; double course = 0;
            if (textBox129.Text != "")
            {
                course=Convert.ToDouble(textBox129.Text);
            }
            if (checkBox59.Checked == true)
            {
                currency = "RUR";
            }
            if (checkBox60.Checked == true)
            {
                currency = "USD";
                SetBookMarkText("YESum", obj_Bookmarks, obj_App, "что эквивалентно "+ Sum+" y.e по курсу " + textBox129.Text+".");
                rubsum = Sum * course;
            }
            if (checkBox61.Checked == true)
            {
                currency = "EUR";
                SetBookMarkText("YESum", obj_Bookmarks, obj_App, "что эквивалентно " + Sum + " y.e по курсу " + textBox129.Text + ".");
                rubsum = Sum * course;
            }
            SetBookMarkText("Way", obj_Bookmarks, obj_App, this.textBox167.Text);
            SetBookMarkText("Tarif", obj_Bookmarks, obj_App, this.textBox168.Text);
            SetBookMarkText("RubSum", obj_Bookmarks, obj_App, rubsum.ToString());
            SetBookMarkText("manager", obj_Bookmarks, obj_App, this.comboBox31.Text);
            SetBookMarkText("FIO2", obj_Bookmarks, obj_App, this.comboBox32.Text);
            SetBookMarkText("Adress", obj_Bookmarks, obj_App, "Адрес:" + this.textBox125.Text);
            SetBookMarkText("Phone", obj_Bookmarks, obj_App, "Тел:" + this.textBox113.Text);
            if (this.comboBox35.SelectedItem != null)
            {
                SetOperatorReqvizit(Word, obj_Bookmarks, obj_App, Doc, this.comboBox35.SelectedItem.ToString());
            }
            //make zayvka
            object[] ExcelParam = new object[1];

            Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
            object Excel = Activator.CreateInstance(obj_excel);

            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
            ExcelParam[0] = Basepath + @"Template\zayavkaNaOlatyAvia.xls";
            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
            ExcelParam[0] = 1;
            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
            /* SetCellData(comboBox6.Text,"D3",obj_worksheet);
             SetCellData(textBox2.Text,"D4",obj_worksheet);
             string date = textBox4.Text + "-" + textBox5.Text;
             SetCellData(date,"D5",obj_worksheet);
             if (dataGridView1.Rows[0].Cells[1].Value != null) { SetCellData(dataGridView1.Rows[0].Cells[1].Value.ToString(), "D6", obj_worksheet); }*/
            if (comboBox35.Text == "Росинтур") 
            {
                SetCellData("ООО ТК \"РОСИНТУР\"", "A1", obj_worksheet);
            }
            else if (comboBox35.Text == "РосинтурЮг")
            {

                    SetCellData("ООО ТК \"РОСИНТУР-ЮГ\"", "A1", obj_worksheet);
             
            }
            
            SetCellData(comboBox32.Text, "B3", obj_worksheet);
            SetCellData("Авиабилет", "B4", obj_worksheet);
            if (dataGridView24.Rows[0].Cells[3].Value != null)
            {
                SetCellData(dataGridView24.Rows[0].Cells[3].Value.ToString(), "B5", obj_worksheet);
            }
            SetCellData(currency, "B18", obj_worksheet);
            SetCellData(rubsum.ToString(), "J18", obj_worksheet);
            SetCellData(textBox129.Text, "E18", obj_worksheet);
            SetCellData("Договор купли-продажи авиабилетов №" + textBox124.Text, "B7", obj_worksheet);
            //SetCellData(textBox18.Text, "D6", obj_worksheet);
            //SetCellData(textBox19.Text, "D7", obj_worksheet);
            //SetCellData("Основной договор № " + textBox49.Text, "D8", obj_worksheet);
            SetCellData(comboBox31.Text, "B22", obj_worksheet);
            for (int i = 0; i < dataGridView24.RowCount; i++)
            {
                if (dataGridView24.Rows[i].Cells[0].Value != null) {SetCellData(dataGridView24.Rows[i].Cells[0].Value.ToString(),"A"+(11+i),obj_worksheet); }
                //if (dataGridView24.Rows[i].Cells[1].Value != null) {SetCellData(dataGridView24.Rows[i].Cells[1].Value.ToString(),"B"+(11+i),obj_worksheet); }
                if (dataGridView24.Rows[i].Cells[2].Value != null) {SetCellData(dataGridView24.Rows[i].Cells[2].Value.ToString(),"B"+(11+i),obj_worksheet); }
            }
            SetCellData(Sum.ToString(), "B17", obj_worksheet);
            SetCellData(Sum.ToString(), "J17", obj_worksheet);
            Param[0] = "true";
            obj_App.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, obj_App, Param);
            ExcelParam[0] = "True";
            Excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, Excel, ExcelParam);
            DocumentsaveA(Doc, obj_workbook, this.textBox124.Text, this.comboBox31.Text, this.comboBox32.Text,"");
            object id="";
            try
            {
                Client ClientData = new Client(comboBox32.Text, "", "", "", "", "", "", "", "", textBox166.Text, "", textBox113.Text, "", textBox125.Text, "","");
                id = GetClientId(ClientData);
                /*if (id == null)
                {
                    id = ClientInsert(ClientData);
                }*/
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при получении клиента в основном договоре \n\r");
            }
            if (id == null)
            {
                id = "";
            }
            try
            {
                AviaDogovorInfo ai = new AviaDogovorInfo(this.textBox124.Text, DateTime.Today.ToShortDateString(), id.ToString(), this.comboBox31.Text, this.comboBox35.Text, this.textBox126.Text, this.textBox127.Text, this.comboBox32.Text, this.textBox125.Text, this.textBox113.Text, this.comboBox37.Text,this.textBox167.Text);
                AviaDogovorSave(ai, dataGridView24);
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при сохранении купли-продажи авиабилетов");
            }
            try
            {
                Section sec = new Section();
                int number = Convert.ToInt32(textBox124.Text);
                number++;
                sec.writekey("Avia", "number_" + this.comboBox35.Text, "app.ini",number.ToString());
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка ");
            }
            //clean W
            Marshal.ReleaseComObject(obj_Tables);
            // Marshal.ReleaseComObject(obj_Selection);
            //Marshal.ReleaseComObject(obj_Range);
            Marshal.ReleaseComObject(obj_Doc);
            Marshal.ReleaseComObject(obj_Bookmarks);
            //Marshal.ReleaseComObject(obj_Bookmark);
            Marshal.ReleaseComObject(obj_App);
            // Marshal.ReleaseComObject(Word);
            // GC.GetTotalMemory(true);
            //clean Ex

            Marshal.ReleaseComObject(obj_workbooks);
            Marshal.ReleaseComObject(obj_worksheet);
            Marshal.ReleaseComObject(obj_workbook);
            Marshal.ReleaseComObject(obj_worksheets);
            button40.Enabled = true;
            button35.Visible = false;
        }

        private void dataGridView24_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                int sum = 0;
                DataGridView data=(DataGridView)sender;
                for (int i = 0; i < data.RowCount; i++)
                {
                    if (data.Rows[i].Cells[2].Value != null)
                    {
                        sum += Convert.ToInt32(data.Rows[i].Cells[2].Value);
                        
                    }
                }
                textBox128.Text = sum.ToString();
            }
        }

        private void checkBox59_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox59.Checked == true)
            {
                checkBox60.Checked = false;
                checkBox61.Checked = false;
                //checkBox60.Enabled = false;
                //checkBox61.Enabled = false;
            }
            else
            {
                /*checkBox60.Enabled = true;
                checkBox61.Enabled = true;*/
                if ((checkBox60.Checked == false) && (checkBox61.Checked == false))
                {
                    checkBox59.Checked = true;
                }
            }
            if (checkBox59.Checked == true)
            {
                label231.Text = "RUR";
            }
        }

        private void checkBox60_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox60.Checked == true)
            {
                checkBox61.Checked = false;
                checkBox59.Checked = false;
                //checkBox59.Enabled = false;
                //checkBox61.Enabled = false;
            }
            else
            {
                //checkBox59.Enabled = true;
                //checkBox61.Enabled = true;
                if ((checkBox59.Checked == false) && (checkBox61.Checked == false))
                {
                    checkBox60.Checked = true;
                }
            }
            if (checkBox60.Checked == true)
            {
                label231.Text = "USD";
            }
        }

        private void checkBox61_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox61.Checked == true)
            {
                checkBox59.Checked = false;
                checkBox60.Checked = false;
                //checkBox60.Enabled = false;
                //checkBox59.Enabled = false;
            }
            else
            {
                //checkBox60.Enabled = false;
                //checkBox59.Enabled = false;
                if ((checkBox60.Checked == false) && (checkBox59.Checked == false))
                {
                    checkBox61.Checked = true;
                }
            }
            if (checkBox61.Checked == true)
            {
                label231.Text = "EUR";
            }
        }

        private void comboBox36_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox36.Text == "Данко Тревел")
            {
                textBox126.Text = "08-688290";
                textBox127.Text = "23.06.2010";
            }
            if (comboBox36.Text == "ПакТур")
            {
                textBox126.Text = "АДУ 2-0211209";
                textBox127.Text = "02.12.2009";
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            Section sec = new Section();
            string num = "";
            num = sec.readkey("Avia", "number_" + this.comboBox35.Text, "app.ini");
            /*if (num != "")
            {
                try
                {
                    number = Convert.ToInt32(num);
                    number++;
                }
                catch
                {
                    this.richTextBox1.AppendText("Ошибка преобразования номера договора авиа \n\r");
                }
            }*/
            textBox124.Text = num;
        }

        private void договорПоАвиабилетамToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox17.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
        }

        private void comboBox32_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void makeClientList(object sender)
        {
            Dictionary<string, string> str = new Dictionary<string, string>();
            ComboBox com = (ComboBox)sender;
            //this.comboBox6.Items.Clear();
            //  string strline = this.comboBox6.SelectedItem.ToString();
            str = Getclients(com.Text);
            if (str.Count != 0)
            {
                foreach (KeyValuePair<string, string> kvp in str)
                {
                    /*if (str[i] != null)
                    {
                        if (!com.Items.Contains(str[i]))
                        {*/
                    com.Items.Add(kvp.Value);
                    clientsSerarch.Add(kvp.Key, (com.Items.Count - 1).ToString());

                    /*     }
                     }*/
                }
                //this.comboBox6.Items.AddRange(str);co
            }
        }
        private void comboBox32_TextChanged(object sender, EventArgs e)
        {
            // makeClientList(sender);
            Dictionary<string,string> str=new Dictionary<string,string>();
            ComboBox com = (ComboBox)sender;
            //this.comboBox6.Items.Clear();
            //  string strline = this.comboBox6.SelectedItem.ToString();
            str = Getclients(com.Text);
            if (str.Count != 0)
            {
                foreach(KeyValuePair<string, string> kvp in str)
                {
                    /*if (str[i] != null)
                    {
                        if (!com.Items.Contains(str[i]))
                        {*/
                            //com.Items.Add(kvp.Value);
                            //clientsSerarch.Add(kvp.Key, (com.Items.Count - 1).ToString());

                   /*     }
                    }*/
                }
                //this.comboBox6.Items.AddRange(str);co
            }
       
        }

        private void label234_Click(object sender, EventArgs e)
        {

        }

        private void button42_Click(object sender, EventArgs e)
        {
            sendobject.Set(sender);
            panel5.Visible = true;
            panel5.BringToFront();
            
        }

        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                GetClientsData();
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при получении данных частных лиц\n\r");
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            panel5.Visible = false;
            sendobject.clean();
        }

        private void dataGridView23_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView data = (DataGridView)sender;
            if (sendobject.owner != null)
            {
                Control hh = (Control)sendobject.owner;

                if (data.Rows[e.RowIndex].Cells[0] != null)
                {
                    if (hh.Name == "button42")
                    {
                        comboBox6.Text = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                        textBox8.Text = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        comboBox9.Text = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                        textBox34.Text = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox23.Text = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {
                        comboBox32.Text = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                    }
                }
                if (data.Rows[e.RowIndex].Cells[1] != null)
                {
                    if (hh.Name == "button42")
                    {
                        maskedTextBox1.Text = data.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        maskedTextBox3.Text = data.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox50.Text = data.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {
                        textBox166.Text = data.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                }
                if (data.Rows[e.RowIndex].Cells[2] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox9.Text = data.Rows[e.RowIndex].Cells[2].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox33.Text = data.Rows[e.RowIndex].Cells[2].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox51.Text = data.Rows[e.RowIndex].Cells[2].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[3] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox110.Text = data.Rows[e.RowIndex].Cells[3].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox107.Text = data.Rows[e.RowIndex].Cells[3].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox52.Text = data.Rows[e.RowIndex].Cells[3].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[4] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox117.Text = data.Rows[e.RowIndex].Cells[4].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox122.Text = data.Rows[e.RowIndex].Cells[4].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {

                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[5] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox116.Text = data.Rows[e.RowIndex].Cells[5].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox121.Text = data.Rows[e.RowIndex].Cells[5].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {

                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[6] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox10.Text = data.Rows[e.RowIndex].Cells[6].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox32.Text = data.Rows[e.RowIndex].Cells[6].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox54.Text = data.Rows[e.RowIndex].Cells[6].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {
                        textBox125.Text = data.Rows[e.RowIndex].Cells[6].Value.ToString();
                    }
                }
                if (data.Rows[e.RowIndex].Cells[7] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox114.Text = data.Rows[e.RowIndex].Cells[7].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox119.Text = data.Rows[e.RowIndex].Cells[7].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {

                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[8] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox112.Text = data.Rows[e.RowIndex].Cells[8].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox109.Text = data.Rows[e.RowIndex].Cells[8].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox53.Text = data.Rows[e.RowIndex].Cells[8].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[9] != null)
                {
                    if (hh.Name == "button42")
                    {
                        maskedTextBox2.Text = data.Rows[e.RowIndex].Cells[9].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        maskedTextBox4.Text = data.Rows[e.RowIndex].Cells[9].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox106.Text = data.Rows[e.RowIndex].Cells[9].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {
                        textBox113.Text = data.Rows[e.RowIndex].Cells[9].Value.ToString();
                    }
                }
                if (data.Rows[e.RowIndex].Cells[10] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox12.Text = data.Rows[e.RowIndex].Cells[10].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox30.Text = data.Rows[e.RowIndex].Cells[10].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        //textBox53.Text = data.Rows[e.RowIndex].Cells[8].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[11] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox111.Text = data.Rows[e.RowIndex].Cells[11].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox108.Text = data.Rows[e.RowIndex].Cells[11].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {
                        textBox55.Text = data.Rows[e.RowIndex].Cells[11].Value.ToString();
                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
                if (data.Rows[e.RowIndex].Cells[12] != null)
                {
                    if (hh.Name == "button42")
                    {
                        textBox115.Text = data.Rows[e.RowIndex].Cells[12].Value.ToString();
                    }
                    else if (hh.Name == "button43")
                    {
                        textBox120.Text = data.Rows[e.RowIndex].Cells[12].Value.ToString();
                    }
                    else if (hh.Name == "button44")
                    {

                    }
                    else if (hh.Name == "button45")
                    {

                    }
                }
            }
                panel5.Visible = false;
                sendobject.clean();
        }

        private void button46_Click(object sender, EventArgs e)
        {
            Client ClientData = new Client(textBox165.Text, textBox157.Text, textBox156.Text, textBox155.Text, textBox154.Text, textBox161.Text, textBox160.Text, textBox159.Text, textBox158.Text, makeSQLdate(maskedTextBox7.Text, '.'), textBox162.Text, maskedTextBox8.Text, textBox152.Text, textBox164.Text, textBox151.Text, "");
            try
            {
                ClientInsert(ClientData);
                textBox130.Text = textBox165.Text;
                EventArgs ev = new EventArgs();
                button37_Click(this.button37, ev);
                tabControl5.SelectedTab = tabPage14;
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при добавлении частного лица в бд\n\r");
            }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            Client ClientData = new Client(textBox148.Text, textBox140.Text, textBox139.Text, textBox138.Text, textBox137.Text, textBox144.Text, textBox143.Text, textBox142.Text, textBox141.Text, makeSQLdate(maskedTextBox5.Text,'.'), textBox145.Text, maskedTextBox6.Text, textBox149.Text, textBox147.Text, textBox150.Text,"");
            try
            {
                DataGridViewSelectedCellCollection cc = dataGridView23.SelectedCells;
                DataGridViewRow row = dataGridView23.Rows[cc[0].RowIndex];
                if (row.Cells[15].Value != null)
                {
                    ClientUpdate(ClientData, row.Cells[15].Value.ToString());
                    textBox130.Text = textBox148.Text;
                    EventArgs ev = new EventArgs();
                    button37_Click(this.button37, ev);
                    tabControl5.SelectedTab = tabPage14;
                }
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при обновлении данных частного лица в бд\n\r");
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            tabControl5.SelectedTab = tabPage15; DataGridViewSelectedCellCollection cc = dataGridView23.SelectedCells;
            ///DataGridViewSelectedRowCollection rc=dataGridView23.SelectedRows;
            ///
            if (cc.Count != 0)
            {
                DataGridViewRow row = dataGridView23.Rows[cc[0].RowIndex];
                if (row.Cells[0].Value != null)
                {
                    textBox148.Text = row.Cells[0].Value.ToString();
                }
                if (row.Cells[1].Value != null)
                {
                    maskedTextBox5.Text = row.Cells[1].Value.ToString();
                }
                if (row.Cells[2].Value != null)
                {
                    textBox144.Text = row.Cells[2].Value.ToString();
                }
                if (row.Cells[3].Value != null)
                {
                    textBox143.Text = row.Cells[3].Value.ToString();
                }
                if (row.Cells[4].Value != null)
                {
                    textBox140.Text = row.Cells[4].Value.ToString();
                }
                if (row.Cells[5].Value != null)
                {
                    textBox139.Text = row.Cells[5].Value.ToString();
                }
                if (row.Cells[6].Value != null)
                {
                    textBox147.Text = row.Cells[6].Value.ToString();
                }
                if (row.Cells[7].Value != null)
                {
                    textBox137.Text = row.Cells[7].Value.ToString();
                }
                if (row.Cells[8].Value != null)
                {
                    textBox141.Text = row.Cells[8].Value.ToString();
                }

                if (row.Cells[9].Value != null)
                {
                    maskedTextBox6.Text = row.Cells[9].Value.ToString();
                }
                if (row.Cells[10].Value != null)
                {
                    textBox145.Text = row.Cells[10].Value.ToString();
                }
                if (row.Cells[11].Value != null)
                {
                    textBox142.Text = row.Cells[11].Value.ToString();
                }
                if (row.Cells[12].Value != null)
                {
                    textBox138.Text = row.Cells[12].Value.ToString();
                }
                if (row.Cells[13].Value != null)
                {
                    textBox149.Text = row.Cells[13].Value.ToString();
                }
                if (row.Cells[14].Value != null)
                {
                    textBox150.Text = row.Cells[14].Value.ToString();
                }
            }
        }

        private void tabPage15_Leave(object sender, EventArgs e)
        {

                textBox148.Text = "";


                maskedTextBox5.Text = "";

                textBox144.Text = "";
    

                textBox143.Text = "";

                textBox140.Text = "";

                textBox139.Text = "";
  
                textBox147.Text = "";
    
                textBox137.Text = "";
  
                textBox141.Text = "";
 
                textBox145.Text = "";
  
                maskedTextBox6.Text = "";

                textBox142.Text = "";
  
                textBox138.Text = "";

                textBox149.Text = "";

                textBox150.Text = "";
            
        }

        private void button35_Click(object sender, EventArgs e)
        {
            button40.Enabled = true;
        }

        private void button49_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Enabled == false)
            {
                richTextBox1.Enabled = true;
            }
            else
            {
                richTextBox1.Enabled = false;
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Point p = new Point();
            p.X = 0;
            p.Y = 0;
            this.AutoScrollPosition = p;
        }
        private ManagerInfo GetmanagerInfo(string name)
        {
            ManagerInfo manager = new ManagerInfo();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader = null;
            SqlCommand sqlcom = null;
            string query = "select name, number,email,icq, phone,skype from managers where name='" + name + "'";
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            if (connect.State == ConnectionState.Open)
            {
                sqlcom = new SqlCommand(query, connect);
                reader = sqlcom.ExecuteReader();
                if (reader.HasRows != false)
                {
                    while(reader.Read())
                    {
                      manager.name=reader["name"].ToString();
                      manager.number=reader["number"].ToString();
                      manager.email=reader["email"].ToString();
                      manager.icq=reader["icq"].ToString();
                      manager.skype=reader["skype"].ToString();
                      manager.skype=reader["phone"].ToString();
                    }
                   
                }
            }
            reader.Close();
            connect.Close();

            return manager;
        }
        //avia_dog
//Confirmation

        private void button55_Click(object sender, EventArgs e)
        {
            object obj_App;
            object obj_Doc;
            object obj_Bookmarks;
            //object obj_Bookmark;
            //object obj_Selection;
            //object obj_Range;
            object obj_Tables;
            object[] Param;
            Param = new object[1];
            //string managernum = getmanagerNum(comboBox52.Text);
            Type obj_Class = Type.GetTypeFromProgID("Word.Application");
            object Word = Activator.CreateInstance(obj_Class);
            Param[0] = Basepath + @"Template\Confirm.doc";
            obj_App = Word.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, Word, null);
            obj_Doc = obj_App.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj_App, null);
            ManagerInfo manager = GetmanagerInfo(comboBox41.Text);
            object Doc = obj_Doc.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_Doc, Param);
            obj_Bookmarks = Doc.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, Doc, null);
            SetBookMarkText("DateNow", obj_Bookmarks, obj_App, DateTime.Now.ToShortDateString());
            SetBookMarkText("FirmName", obj_Bookmarks, obj_App, this.textBox201.Text + "," + comboBox38.Text);
            SetBookMarkText("ClientName", obj_Bookmarks, obj_App, this.textBox177.Text);
            SetBookMarkText("ClientPhone", obj_Bookmarks, obj_App, this.textBox175.Text);
            SetBookMarkText("ClientEmail", obj_Bookmarks, obj_App, this.textBox172.Text);
            SetBookMarkText("Way", obj_Bookmarks, obj_App, this.textBox178.Text);
            SetBookMarkText("TravelDate", obj_Bookmarks, obj_App, this.dateTimePicker13.Value.ToShortDateString() + " - " + this.dateTimePicker12.Value.ToShortDateString());
            SetBookMarkText("ManNum", obj_Bookmarks, obj_App, this.textBox171.Text);
            SetBookMarkText("DogNum", obj_Bookmarks, obj_App, this.textBox179.Text);
            SetBookMarkText("zayavkanum", obj_Bookmarks, obj_App, this.textBox197.Text);
            obj_Tables = Doc.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, Doc, null);
            if (dataGridView25.RowCount > 3)
            {
                TableSize(dataGridView25, obj_Tables, 2, 3);
            }
            TableProcess(dataGridView25, obj_Tables, 2);
            if (dataGridView26.RowCount > 2)
            {
                TableSize(dataGridView26, obj_Tables, 3, 2);
            } 
            TableProcess(dataGridView26, obj_Tables, 3);
           /* if (checkBox69.Checked == true)
            {
                SetBookMarkText("Journey", obj_Bookmarks, obj_App, checkBox69.Text);
            }
            else if (checkBox68.Checked == true)
            {
                SetBookMarkText("Journey", obj_Bookmarks, obj_App, checkBox68.Text);
            }
            else if (checkBox67.Checked == true)
            {
                SetBookMarkText("Journey", obj_Bookmarks, obj_App, checkBox67.Text);
            }*/
            SetBookMarkText("Journey", obj_Bookmarks, obj_App, comboBox55.Text);
            SetBookMarkText("Transfer", obj_Bookmarks, obj_App, this.comboBox39.Text);
            SetBookMarkText("Hotel", obj_Bookmarks, obj_App, this.textBox185.Text);
            SetBookMarkText("Tour", obj_Bookmarks, obj_App, this.comboBox48.Text);
            SetBookMarkText("RoomType", obj_Bookmarks, obj_App, this.comboBox46.Text + "/" + this.comboBox45.Text);
            SetBookMarkText("FoodType", obj_Bookmarks, obj_App, this.comboBox40.Text);
            SetBookMarkText("Viza", obj_Bookmarks, obj_App, this.textBox186.Text);
            SetBookMarkText("Excoursion", obj_Bookmarks, obj_App, this.textBox189.Text);
            SetBookMarkText("AddService", obj_Bookmarks, obj_App, this.textBox177.Text);
            SetBookMarkText("MedicalStrach", obj_Bookmarks, obj_App, this.textBox188.Text);
            SetBookMarkText("CanselStrach", obj_Bookmarks, obj_App, this.textBox187.Text);
            SetBookMarkText("Partner", obj_Bookmarks, obj_App, this.comboBox47.Text);
            //mahangerdata
            string fullsum = "";string finalsum="";
            if (textBox180.Text == "1")
            {
                fullsum = textBox192.Text + " ";
            }
            else
            {
                fullsum += textBox192.Text + "x" + textBox180.Text;
            }
            if ((textBox173.Text!="")&&(textBox170.Text!=""))
            {
                fullsum += " + "+textBox173.Text + "x" + textBox170.Text;
            }
            if ((textBox176.Text != "") && (textBox174.Text != ""))
            {
                fullsum += " + " + textBox176.Text + "x" + textBox174.Text;
            }
            if ((textBox184.Text != "") && (textBox183.Text != ""))
            {
                fullsum += " + " + textBox184.Text + "x" + textBox183.Text;
            }
            if ((textBox191.Text != "") && (textBox190.Text != ""))
            {
                fullsum += " + " + textBox191.Text + "x" + textBox190.Text;
            }
            fullsum += " = " + dataGridView27.Rows[0].Cells[7].Value.ToString();
            SetBookMarkText("TourFullPrice", obj_Bookmarks, obj_App, fullsum);
            SetBookMarkText("AgentSum", obj_Bookmarks, obj_App, this.textBox182.Text + " % - " + this.textBox181.Text);
            if (dataGridView27.Rows[0].Cells[6].Value != null)
            {
                finalsum=dataGridView27.Rows[0].Cells[6].Value.ToString();
                SetBookMarkText("FinalPrice", obj_Bookmarks, obj_App, dataGridView27.Rows[0].Cells[8].Value.ToString());
            }
           

            SetBookMarkText("Manager_name", obj_Bookmarks, obj_App, manager.name);
            SetBookMarkText("Manager_ICQ", obj_Bookmarks, obj_App, "ICQ "+manager.icq);
            SetBookMarkText("Manager_email", obj_Bookmarks, obj_App, manager.email);
            SetBookMarkText("PredPayHalf", obj_Bookmarks, obj_App, this.dateTimePicker15.Value.ToShortDateString());
            SetBookMarkText("PayFull", obj_Bookmarks, obj_App, this.dateTimePicker14.Value.ToShortDateString());
            SetBookMarkText("DocumentDate", obj_Bookmarks, obj_App, this.dateTimePicker16.Value.ToShortDateString());
            SetBookMarkText("penalty100", obj_Bookmarks, obj_App, this.textBox193.Text);
            SetBookMarkText("penalty50", obj_Bookmarks, obj_App, this.textBox194.Text);
            SetBookMarkText("FirmaCours", obj_Bookmarks, obj_App, this.textBox195.Text);
            //make_invoice


            //sqlsave
            string agentid = "";
            if ((agent_key != null) && (agent_key != ""))
            {
                agentid = GetagentKey(textBox201.Text, textBox177.Text, comboBox38.Text, textBox175.Text, textBox172.Text);
            }
            else
            {
                agentid = agent_key;
            }
            DataGridViewCell cell = dataGridView31.SelectedCells[0];
            DataGridViewRow row = dataGridView31.Rows[cell.RowIndex];
            string cId="";
            ConfirInfo cinfo = new ConfirInfo(row.Cells[11].Value.ToString(), textBox178.Text, numericUpDown12.Value.ToString(),numericUpDown13.Value.ToString(), dateTimePicker13.Value.ToShortDateString(), dateTimePicker12.Value.ToShortDateString(), textBox185.Text, comboBox48.Text, comboBox41.Text, fullsum, textBox181.Text, textBox182.Text, finalsum, comboBox47.Text, agent_key,comboBox49.Text,comboBox52.Text);
            try
            {
                cId = ConfirmSave(cinfo);
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при получении клиента в основном договоре \n\r");
            }
            //make bill
            object[] ExcelParam = new object[1];
            string touroperator = comboBox52.Text;
            Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
            object Excel = Activator.CreateInstance(obj_excel);

            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
            if (touroperator == "Росинтур")
            {
                ExcelParam[0] = Basepath + @"Template\Invoice_rosintour.xls";
            }
            else if (touroperator == "РосинтурЮг")
            {
                ExcelParam[0] = Basepath + @"Template\Invoice_rosintourUg.xls";
            }
            else if (touroperator == "Магазин Путешествий")
            {
                ExcelParam[0] = Basepath + @"Template\Invoice_travelMag.xls";
            }
            //ExcelParam[0] = Basepath + @"Template\Manager_report.xls";
            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
            ExcelParam[0] = 1;
            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
            SetCellData(label345.Text + " от " + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year + " г.", "B11", obj_worksheet);
            SetCellData(textBox212.Text, "C14", obj_worksheet);
            SetCellData(textBox211.Text, "C15", obj_worksheet);
            for (int i = 0; i < dataGridView29.RowCount; i++)
            {

                if (dataGridView29.Rows[i].Cells[0].Value != null)
                {
                    SetCellData(dataGridView29.Rows[i].Cells[0].Value.ToString(), "B" + (17 + i).ToString(), obj_worksheet);
                }
                if (dataGridView29.Rows[i].Cells[1].Value != null)
                {
                    SetCellData(dataGridView29.Rows[i].Cells[1].Value.ToString(), "c" + (17 + i).ToString(), obj_worksheet);
                }
                if (dataGridView29.Rows[i].Cells[2].Value != null)
                {
                    SetCellData(dataGridView29.Rows[i].Cells[2].Value.ToString(), "D" + (17 + i).ToString(), obj_worksheet);
                }
                if (dataGridView29.Rows[i].Cells[3].Value != null)
                {
                    SetCellData(dataGridView29.Rows[i].Cells[3].Value.ToString(), "F" + (17 + i).ToString(), obj_worksheet);
                }
                if (dataGridView29.Rows[i].Cells[4].Value != null)
                {
                    SetCellData(dataGridView29.Rows[i].Cells[4].Value.ToString(), "G" + (17 + i).ToString(), obj_worksheet);
                }
                if (dataGridView29.Rows[i].Cells[5].Value != null)
                {
                    SetCellData(dataGridView29.Rows[i].Cells[5].Value.ToString(), "H" + (17 + i).ToString(), obj_worksheet);
                }
            }
            if (dataGridView30.Rows[0].Cells[0].Value != null)
            {
                SetCellData(dataGridView30.Rows[0].Cells[0].Value.ToString(), "H19", obj_worksheet);
            }
            if (dataGridView30.Rows[1].Cells[0].Value != null)
            {
                SetCellData(dataGridView30.Rows[1].Cells[0].Value.ToString(), "H20", obj_worksheet);
            }
            string currency = "";
            if (checkBox65.Checked == true)
            {
                currency = "RUR";
            }
            else if (checkBox64.Checked == true)
            {
                currency = "USD";
            }
            else if (checkBox63.Checked == true)
            {
                currency = "EUR";
            }
            string curname = "";
            if (currency == "RUR")
            {
                curname = getcurrencyname(textBox214.Text,currency);
                SetCellData("Всего наименований " + textBox213.Text + " на сумму " + textBox214.Text + " " + curname, "B23", obj_worksheet);
                SetCellData("Счет действителен в течении " + textBox216.Text + " " + getbankdayword(Convert.ToInt32(textBox216.Text)), "B25", obj_worksheet);
            }
            else if ((currency == "EUR") || (currency == "USD"))
            {
                curname = "y.e";
                SetCellData("Курс 1 у.е = 1 " + currency + " = " + textBox215.Text, "B22", obj_worksheet);
                SetCellData("Всего наименований " + textBox213.Text + " на сумму " + textBox214.Text + " " + curname, "B23", obj_worksheet);
                SetCellData("Оплата по курсу туроператора на день оплаты.", "B24", obj_worksheet);
                //SetCellData("Оплата по курсу туроператора на день оплаты. "+"1 y.e = 1 "+currency+" = "+textBox215.Text, "B24", obj_worksheet);
                SetCellData("Счет действителен в течении " + textBox216.Text + " " + getbankdayword(Convert.ToInt32(textBox216.Text)), "B25", obj_worksheet);
            }
            //view
            Param[0] = "true";
            obj_App.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, obj_App, Param);
            ExcelParam[0] = "True";
            Excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, Excel, ExcelParam);
            setdefmanager(comboBox41.Text);
            try
            {
                DocumentAgentSave(Doc, obj_workbook, textBox179.Text, manager.name, manager.number, textBox201.Text + "+" + (Convert.ToInt32(textBox171.Text) - 1).ToString(), comboBox52.Text);
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка при сохранении потверждения и счета \n\r");
            }
            try
            {
                Section sec = new Section();
                int number = Convert.ToInt32(textBox179.Text);
                number++;
                sec.writekey("AgentDocCount", "number_" + this.comboBox52.Text, "app.ini", number.ToString());
            }
            catch
            {
                this.richTextBox1.AppendText("Ошибка увеличения номера договора");
            }

            //clean W
            Marshal.ReleaseComObject(obj_Tables);
            // Marshal.ReleaseComObject(obj_Selection);
            //Marshal.ReleaseComObject(obj_Range);
            Marshal.ReleaseComObject(obj_Doc);
            Marshal.ReleaseComObject(obj_Bookmarks);
            //Marshal.ReleaseComObject(obj_Bookmark);
            Marshal.ReleaseComObject(obj_App);
            // Marshal.ReleaseComObject(Word);
            // GC.GetTotalMemory(true);
            //clean Ex

            Marshal.ReleaseComObject(obj_workbooks);
            Marshal.ReleaseComObject(obj_worksheet);
            Marshal.ReleaseComObject(obj_workbook);
            Marshal.ReleaseComObject(obj_worksheets);
        }

        private void DocumentAgentSave(object ODoc, object OWorkBook, string num, string managername, string managernum, string client,string touroperator)
        {
            Section sec = new Section();

            object[] WordParam = new object[1];
            object[] ExcelParam = new object[1];
            string path = Path.GetFullPath(sec.readkey("SavePath", "AgentPath", "app.ini"));
            string confirmpath = "";string invoicepath = "";
            if (!Directory.Exists(path))
            {
                    Directory.CreateDirectory(path);
            }
            if ((path != null) && (Directory.Exists(path)))
            {
                confirmpath=path+ "\\Подтвержения";
                invoicepath=path+ "\\Счета";
                    if (touroperator != "")
                    {
                        confirmpath+="\\"+touroperator;
                        invoicepath+="\\"+touroperator;
                    }


                    if (managername != "")
                    {
                        confirmpath += "\\" + managername + " " + managernum + "\\" + DateTime.Now.Year.ToString();
                        invoicepath += "\\" + managername + " " + managernum + "\\" + DateTime.Now.Year.ToString();
                    }
                if (!Directory.Exists(confirmpath))
                {
                    Directory.CreateDirectory(confirmpath);
                }
                if (!Directory.Exists(invoicepath))
                {
                    Directory.CreateDirectory(invoicepath);
                }
                ExcelParam[0] = CheckFileName(invoicepath + "\\Счет " + num + " " + client, ".xls");
                WordParam[0] = CheckFileName(confirmpath + "\\Подтверждение " + num + " " + client, ".doc");
                ODoc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ODoc, WordParam);
                OWorkBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, OWorkBook, ExcelParam);
            }
            else
            {
                string localpath;
                if (!Directory.Exists("c:\\Счета и Подтверждения"))
                {
                    Directory.CreateDirectory("c:\\Счета и Подтверждения");
                    localpath = "c:\\Счета и Подтверждения";
                }
                else
                {
                    localpath = "c:\\Счета и Подтверждения";
                }

                //object Doc = arguments.Doc;
                //object WordApp = arguments.App;
                WordParam[0] = localpath + "\\Подтверждение " + num + " " + client;
                ODoc.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ODoc, WordParam);
                ExcelParam[0] = localpath + "\\Счет " + num + " " +client;
                OWorkBook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, OWorkBook, ExcelParam);

            }
        }
       /*private void make_invoice_a()
        {
            object[] ExcelParam = new object[1];
            string touroperator = comboBox52.Text;
            Type obj_excel = Type.GetTypeFromProgID("Excel.Application");
            object Excel = Activator.CreateInstance(obj_excel);

            object obj_workbooks = Excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, Excel, null);
            if (touroperator=="Росинтур")
            {
                ExcelParam[0] = Basepath + @"Template\Invoice_rosintour.xls";
            }
            else if (touroperator=="РосинтурЮг")
            {
                ExcelParam[0] = Basepath + @"Template\Invoice_rosintourUg.xls";
            }
            else if (touroperator=="Магазин Путешествий")
            {
                ExcelParam[0] = Basepath + @"Template\Invoice_travelMag.xls";
            }
            //ExcelParam[0] = Basepath + @"Template\Manager_report.xls";
            object obj_workbook = obj_workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, obj_workbooks, ExcelParam);
            object obj_worksheets = obj_workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, obj_workbook, null);
            ExcelParam[0] = 1;
            object obj_worksheet = obj_worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, obj_worksheets, ExcelParam);
            SetCellData(label345.Text + " от " + DateTime.Now.Day + " " + ((Month)DateTime.Now.Month).ToString() + " " + DateTime.Now.Year+" г.", "B11", obj_worksheet);
            SetCellData(textBox212.Text, "C14", obj_worksheet);
            SetCellData(textBox211.Text, "C15", obj_worksheet);
            for (int i = 0; i < dataGridView27.RowCount; i++)
            {
               
                    if (dataGridView29.Rows[i].Cells[0].Value != null)
                    {
                        SetCellData(dataGridView29.Rows[i].Cells[0].Value.ToString(), "B"+(17+i).ToString(), obj_worksheet);
                    }
                    if (dataGridView29.Rows[i].Cells[1].Value != null)
                    {
                        SetCellData(dataGridView29.Rows[i].Cells[1].Value.ToString(), "c" + (17 + i).ToString(), obj_worksheet);
                    }
                    if (dataGridView29.Rows[i].Cells[2].Value != null)
                    {
                        SetCellData(dataGridView29.Rows[i].Cells[2].Value.ToString(), "D" + (17 + i).ToString(), obj_worksheet);
                    }
                    if (dataGridView29.Rows[i].Cells[3].Value != null)
                    {
                        SetCellData(dataGridView29.Rows[i].Cells[3].Value.ToString(), "F" + (17 + i).ToString(), obj_worksheet);
                    }
                    if (dataGridView29.Rows[i].Cells[4].Value != null)
                    {
                        SetCellData(dataGridView29.Rows[i].Cells[4].Value.ToString(), "G" + (17 + i).ToString(), obj_worksheet);
                    }
                    if (dataGridView29.Rows[i].Cells[5].Value != null)
                    {
                        SetCellData(dataGridView29.Rows[i].Cells[5].Value.ToString(), "H" + (17 + i).ToString(), obj_worksheet);
                    }
            }
            if (dataGridView30.Rows[0].Cells[1].Value != null)
            {
                SetCellData(dataGridView30.Rows[0].Cells[1].Value.ToString(), "H19", obj_worksheet);
            }
            if (dataGridView30.Rows[1].Cells[1].Value != null)
            {
                SetCellData(dataGridView30.Rows[1].Cells[1].Value.ToString(), "H20", obj_worksheet);
            }
            string currency = "";
            if (checkBox65.Checked == true)
            {
                currency = "RUR";
            }
            else if (checkBox64.Checked == true)
            {
                currency = "USD";
            }
            else if (checkBox63.Checked == true)
            {
                currency = "EUR";
            }
            string curname = "";
            if (currency == "RUR")
            {
                curname = getcurrencyname(textBox214.Text);
                SetCellData("Всего наименований " + textBox213.Text + " на сумму " + textBox214.Text + " " + curname, "B23", obj_worksheet);
                SetCellData("Счет действителен в течении " +textBox216.Text+" "+ getbankdayword(Convert.ToInt32(textBox216.Text)), "B25", obj_worksheet);
            }
            else if ((currency == "EUR")||(currency == "USD"))
            {
             curname = "y.e";
                SetCellData("Всего наименований " + textBox213.Text + " на сумму " + textBox214.Text + " " + curname, "B23", obj_worksheet);
                SetCellData("Оплата по курсу туроператора на день оплаты.", "B24", obj_worksheet);
                //SetCellData("Оплата по курсу туроператора на день оплаты. "+"1 y.e = 1 "+currency+" = "+textBox215.Text, "B24", obj_worksheet);
                SetCellData("Счет действителен в течении " +textBox216.Text+" "+ getbankdayword(Convert.ToInt32(textBox216.Text)), "B25", obj_worksheet);
            }
            
        }*/
        /*private int getfullprice(string tour, string tournum,string agentsum)
        {
            int price = 0;


            
            return price;
        }*/
        private string getbankdayword(int str)
        {
            string result = "";
            switch (str)
            {
                case 1: result = "(одного) банковского дня"; break;
                case 2: result = "(двух) банковских дней"; break;
                case 3: result = "(трех) банковских дней"; break;
                case 4: result = "(четырех) банковских дней"; break;
                case 5: result = "(пяти) банковских дней"; break;
                case 6: result = "(шести) банковских дней"; break;
                case 7: result = "(семи) банковских дней"; break;
                case 8: result = "(восьми) банковских дней"; break;
                case 9: result = "(девяти) банковских дней"; break;
                case 10: result = "(десяти) банковских дней"; break;
                case 11: result = "(одиннадцати) банковских дней"; break;
                case 12: result = "(двенадцати) банковских дней"; break;
                case 13: result = "(тринадцати) банковских дней"; break;
                case 14: result = "(четырнадцати) банковских дней"; break;
                case 15: result = "(пятнадцати) банковских дней"; break;
                case 16: result = "(шеснадцати) банковских дней"; break;
                case 17: result = "(семнадцати) банковских дней"; break;
                case 18: result = "(восемнадцати) банковских дней"; break;
                case 19: result = "(девятнадцати) банковских дней"; break;
                case 20: result = "(двадцати) банковских дней"; break;
            }
            return result;
        }
        private string getcurrencyname(string str,string currencytype)
        {
            string result = "";
            int currency = 0;
            if (currencytype == "RUR")
            {
                currency = 1;
            }
            if (currencytype == "EUR")
            {
                currency = 2;
            }
            if (currencytype == "USD")
            {
                currency = 3;
            }
            //Regex exp = new Regex(@"[0-9]|(?=([0-9]))[0-9]|(?=([0-9][0-9]))[0-9]|1[0-9]|(?=([0-9]))1[0-9]");
            //Match m = exp.Match(str);
            switch (currency)
            {
                case 1:
                    if ((str.Length >= 2) && (str[str.Length - 2].ToString() == "1"))
                    {
                            result = "рублей";
                    }
                    else
                    {
                        if (str[str.Length - 1].ToString() == "1") { result = "рубль"; }
                        if ((str[str.Length - 1].ToString() == "2") || (str[str.Length - 1].ToString() == "3") || (str[str.Length - 1].ToString() == "4")) { result = "рубля"; }
                        if ((str[str.Length - 1].ToString() != "1") && (str[str.Length - 1].ToString() != "2") && (str[str.Length - 1].ToString() != "3") && (str[str.Length - 1].ToString() != "4")) { result = "рублей"; }
                    }
                    break;
                case 2:
                    if ((str.Length >= 2)&&(str[str.Length - 2].ToString() == "1"))
                    {
                           result = "евро";
                    }
                    else
                    {
                        if (str[str.Length - 1].ToString() == "1") { result = "евро"; }
                        if ((str[str.Length - 1].ToString() == "2") || (str[str.Length - 1].ToString() == "3") || (str[str.Length - 1].ToString() == "4")) { result = "евро"; }
                        if ((str[str.Length - 1].ToString() != "1") && (str[str.Length - 1].ToString() != "2") && (str[str.Length - 1].ToString() != "3") && (str[str.Length - 1].ToString() != "4")) { result = "евро"; }
                    }
                    break;
                case 3: if ((str.Length >= 2) && (str[str.Length - 2].ToString() == "1"))
                    {
                            result = "долларов";
                    }
                    else
                    {
                        if (str[str.Length - 1].ToString() == "1") { result = "доллар"; }
                        if ((str[str.Length - 1].ToString() == "2") || (str[str.Length - 1].ToString() == "3") || (str[str.Length - 1].ToString() == "4")) { result = "доллара"; }
                        if ((str[str.Length - 1].ToString() != "1") && (str[str.Length - 1].ToString() != "2") && (str[str.Length - 1].ToString() != "3") && (str[str.Length - 1].ToString() != "4")) { result = "долларов"; }
                    }
                    break;
            }
            return result;
        }
        private string GetagentKey(string name, string agentname, string city, string phone, string email)
        {
            string id = "";
            string query = "";
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "select id,name from agency where name='"+name+"'";
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            if (reader["id"] != null)
                            {
                                id = reader["id"].ToString();
                            }
                        }
                    }
                    else
                    {
                       id = AddAgency(name, agentname, city, phone, email);
                    }
                    reader.Close();
                }
            }
            connect.Close();
            return id;
        }
        private string AddAgency(string name, string agentname,string city, string phone, string email)
        {
            string id = "";
            string query = "";
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "insert into agency values('"+name+"','"+agentname+"','"+city+"','"+phone+"','"+email+"')";
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                    query = "select MAX(id) as id from agency";
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            if (reader["id"] != null)
                            {
                                id = reader["id"].ToString();
                            }
                        }
                    }
                }

            }
            return id;
        }
        private void textBox192_TextChanged(object sender, EventArgs e)
        {
            int sum = 0; double agentsum = 0; int viza = 0; int medstrach = 0; double canselstrach = 0; int addsum = 0; double sumturpaket = 0;
            if ((this.textBox192.Text != "") && (this.textBox180.Text != ""))
            {
                sum = (Convert.ToInt32(textBox192.Text) * Convert.ToInt32(textBox180.Text));
            }
            if (this.textBox182.Text != "")
            {
                try
                {
                agentsum = ((Convert.ToDouble(sum) /100)* Convert.ToDouble(textBox182.Text));
                this.textBox181.Text = agentsum.ToString();
                   }
                catch
                {
                    MessageBox.Show("Вводите цифры");
                }
            }
            if ((this.textBox173.Text != "") && (this.textBox170.Text != ""))
            {
                viza = (Convert.ToInt32(textBox173.Text) * Convert.ToInt32(textBox170.Text));
                textBox186.Text = textBox173.Text + " x " + textBox170.Text;
            }
            if ((this.textBox176.Text != "") && (this.textBox174.Text != ""))
            {
                medstrach = (Convert.ToInt32(textBox176.Text) * Convert.ToInt32(textBox174.Text));
                textBox188.Text = textBox176.Text + " x " + textBox174.Text;
            }
            if ((this.textBox191.Text != "") && (this.textBox190.Text != ""))
            {
                addsum = (Convert.ToInt32(textBox191.Text) * Convert.ToInt32(textBox190.Text));
                textBox196.Text = textBox191.Text + " x " + textBox190.Text;
            }
            if ((this.textBox184.Text != "")&&(this.textBox183.Text != ""))
            {
                try
                {
                    //sumturpaket = Convert.ToDouble(sum + addsum + viza);
                    //canselstrach = ((sumturpaket/ 100) * Convert.ToDouble(textBox184.Text));
                    canselstrach = (Convert.ToInt32(textBox184.Text) * Convert.ToInt32(textBox183.Text));
                    textBox187.Text = textBox184.Text + " x " + textBox183.Text;
                    //this.textBox183.Text = canselstrach.ToString();
                }
                catch
                {
                    MessageBox.Show("Вводите цифры");
                }
            }
            dataGridView27.Rows[0].Cells[0].Value = sum;
            dataGridView27.Rows[0].Cells[1].Value=this.textBox182.Text;
            dataGridView27.Rows[0].Cells[2].Value=agentsum;
            dataGridView27.Rows[0].Cells[3].Value=medstrach;
            dataGridView27.Rows[0].Cells[4].Value=canselstrach;
            dataGridView27.Rows[0].Cells[5].Value = viza;
            dataGridView27.Rows[0].Cells[6].Value = addsum;
        }

        private void textBox182_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView27_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            double finalSum = 0; double fullSum = 0;
            for (int i = 0; i < dataGridView27.ColumnCount-2; i++)
            {
                if ((dataGridView27.Rows[0].Cells[i].Value != null) && (dataGridView27.Rows[0].Cells[i].Value.ToString() != "") && (i != 2) && (i != 1))
                {
                    finalSum += Convert.ToDouble(dataGridView27.Rows[0].Cells[i].Value);
                    fullSum += Convert.ToDouble(dataGridView27.Rows[0].Cells[i].Value);
                }
                else  if (i == 2)
                {
                    finalSum -= Convert.ToDouble(dataGridView27.Rows[0].Cells[i].Value);
                }
            }
            dataGridView27.Rows[0].Cells[7].Value = Convert.ToInt32(fullSum);
            dataGridView27.Rows[0].Cells[8].Value = Convert.ToInt32(finalSum);
        }

        //SQL
        private string ConfirmSave(ConfirInfo cinfo)
        {
            string c_id = "";
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "insert into Agent_confirm values('" + cinfo.Path + "','" + cinfo.Adult + "','" + cinfo.Child + "','" + cinfo.Startdate + "','" + cinfo.EndDate + "','" + cinfo.Hotel + "','" + cinfo.Tour + "','" + cinfo.Manager + "','" + cinfo.FullSum + "','" + cinfo.AgentSum + "','" + cinfo.AgentProcent + "','" + cinfo.FinalSum + "','" + cinfo.Partner + "','" + cinfo.Agentkey + "','" + cinfo.Touroperator+"')";
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                    query = "select MAX(id) from Agent_confirm";
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            c_id = reader["id"].ToString();
                        }
                    }
                    reader.Close();
                }

            }
            connect.Close();
            return c_id;
        }

        private void button60_Click(object sender, EventArgs e)
        {
            panel6.Visible = false;
        }

        private void button51_Click(object sender, EventArgs e)
        {
            panel6.BringToFront();
            panel6.Visible = true;
        }

        private void button59_Click(object sender, EventArgs e)
        {
            if (dataGridView28.SelectedCells[0] != null)
            {
                DataGridViewCell sc = dataGridView28.SelectedCells[0];
                DataGridViewRow row = dataGridView28.Rows[sc.RowIndex];
                if (row.Cells[0].Value != null)
                {
                    textBox207.Text = row.Cells[0].Value.ToString();
                }
                if (row.Cells[1].Value != null)
                {
                    textBox210.Text = row.Cells[1].Value.ToString();
                }
                if (row.Cells[2].Value != null)
                {
                    textBox208.Text = row.Cells[2].Value.ToString();
                }
                if (row.Cells[3].Value != null)
                {
                    textBox209.Text = row.Cells[3].Value.ToString();
                }
                if (row.Cells[4].Value != null)
                {
                    comboBox51.Text = row.Cells[4].Value.ToString();
                }
                tabControl7.SelectedTab = tabPage21;
                //tabControl7.TabPages[0].
                //tabPage19.CanSelect = false;
                //tabPage20.CanSelect = false;
                /*if (row.Cells[5].Value != null)
                {
                    agent_key=row.Cells[5].Value.ToString();
                }*/
                
            }
        }

        private void button61_Click(object sender, EventArgs e)
        {
            try
            {
                getagencyresult(textBox202.Text, textBox200.Text,comboBox50.Text,textBox198.Text,textBox199.Text);
            }
            catch
            {

            }


        }
        private void getagencyresult(string name, string agentname, string city, string phone, string email)
        {
            int first = 0;
            dataGridView28.RowCount = 1;
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "select * from agency";
            if ((name != "") || (agentname != "") || (city != "") || (phone != "") || (email != ""))
            {
                query += " where ";
                if (name != "")
                {
                    if (first == 0)
                    {
                        query += " name='" + name + "'";
                        first = 1;
                    }
                }
                if (agentname != "")
                {
                    if (first == 0)
                    {
                        query += " agentname='" + agentname + "'";
                        first = 1;
                    }
                    else
                    {
                        query += " and agentname='" + agentname + "'";
                    }
                }
                if (city != "")
                {
                    if (first == 0)
                    {
                        query += " city='" + city + "'";
                        first = 1;
                    }
                    else
                    {
                        query += " city='" + city + "'";
                    }
                }
                if (phone != "")
                {
                    if (first == 0)
                    {
                        query += " phone='" + phone + "'";
                        first = 1;
                    }
                    else
                    {
                        query += " phone='" + phone + "'";
                    }
                }
                if (email != "")
                {
                    if (first == 0)
                    {
                        query += " email='" + email + "'";
                        first = 1;
                    }
                    else
                    {
                        query += " email='" + email + "'";
                    }
                }
            }
            int count=0;
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                             dataGridView28.Rows.Add();
                            if (reader["Name"] != null)
                            {
                                dataGridView28.Rows[count].Cells[0].Value=reader["Name"].ToString();
                            }
                            if (reader["City"] != null)
                            {
                                dataGridView28.Rows[count].Cells[1].Value = reader["City"].ToString();
                            }
                            if (reader["Agentname"] != null)
                            {
                                dataGridView28.Rows[count].Cells[2].Value = reader["Agentname"].ToString();
                            }
                            if (reader["Phone"] != null)
                            {
                                dataGridView28.Rows[count].Cells[3].Value = reader["Phone"].ToString();
                            }
                            if (reader["Email"] != null)
                            {
                                dataGridView28.Rows[count].Cells[4].Value = reader["Email"].ToString();
                            }
                            if (reader["id"] != null)
                            {
                                dataGridView28.Rows[count].Cells[5].Value = reader["id"].ToString();
                            }
                            count++;
                        }
                    }
                    reader.Close();
                }
            }
            connect.Close();
        }

        private void button57_Click(object sender, EventArgs e)
        {
            string id="";
            try
            {
                id=AddAgency(textBox203.Text, textBox206.Text,comboBox42.Text, textBox204.Text, textBox205.Text);
                textBox202.Text = textBox203.Text;
                textBox200.Text = textBox206.Text;
                comboBox50.Text = comboBox42.Text;
                //textBox204.Text = textBox205.Text;
                getagencyresult(textBox202.Text, textBox200.Text, comboBox50.Text, textBox198.Text, textBox199.Text);
                tabControl7.SelectedTab = tabPage19;
                
            }
            catch
            {

            }

        }

        private void button58_Click(object sender, EventArgs e)
        {
            DataGridViewCell sc=dataGridView28.SelectedCells[0];
            DataGridViewRow row = dataGridView28.Rows[sc.RowIndex];
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlCommand sqlcom = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "update agency set name='" + textBox207.Text + "', Agentname='" + textBox210.Text + "', City='" + comboBox51.Text + "', Phone='" + textBox208.Text + "', Email='" + textBox209.Text + "' where id='"+row.Cells[5].Value.ToString()+"'";
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    sqlcom.ExecuteNonQuery();
                }
            }
            connect.Close();
            textBox202.Text = textBox207.Text;
            textBox200.Text = textBox210.Text;
            comboBox50.Text = comboBox51.Text;
            //textBox204.Text = textBox205.Text;
            getagencyresult(textBox202.Text, textBox200.Text, comboBox50.Text, textBox198.Text, textBox199.Text);
            tabControl7.SelectedTab = tabPage19;
        }

        private void tabPage21_Click(object sender, EventArgs e)
        {
        }

        private void tabControl7_Selected(object sender, TabControlEventArgs e)
        {
            if (tabControl7.SelectedTab == tabPage21)
            {
                if (dataGridView28.SelectedCells[0] != null)
                {
                    DataGridViewCell sc = dataGridView28.SelectedCells[0];
                    DataGridViewRow row = dataGridView28.Rows[sc.RowIndex];
                    if (row.Cells[0].Value != null)
                    {
                        textBox207.Text = row.Cells[0].Value.ToString();
                    }
                    if (row.Cells[1].Value != null)
                    {
                        textBox210.Text = row.Cells[1].Value.ToString();
                    }
                    if (row.Cells[2].Value != null)
                    {
                        textBox208.Text = row.Cells[2].Value.ToString();
                    }
                    if (row.Cells[3].Value != null)
                    {
                        textBox209.Text = row.Cells[3].Value.ToString();
                    }
                    if (row.Cells[4].Value != null)
                    {
                        comboBox51.Text = row.Cells[4].Value.ToString();
                    }
                }
            }
        }

        private void button56_Click(object sender, EventArgs e)
        {
            DataGridViewCell sc = dataGridView28.SelectedCells[0];
            DataGridViewRow row = dataGridView28.Rows[sc.RowIndex];
            panel6.Visible = false;
            if ((row.Cells[5].Value!=null)&&(row.Cells[5].Value.ToString()!=""))
            {   
                agent_key=row.Cells[5].Value.ToString();
                if (row.Cells[0].Value != null)
                {
                    textBox201.Text = row.Cells[0].Value.ToString();
                }
                if (row.Cells[1].Value != null)
                {
                    comboBox38.Text = row.Cells[1].Value.ToString();
                }
                if (row.Cells[2].Value != null)
                {
                    textBox177.Text = row.Cells[2].Value.ToString();
                }
                if (row.Cells[3].Value != null)
                {
                    textBox175.Text = row.Cells[3].Value.ToString();
                }
                if (row.Cells[4].Value != null)
                {
                    textBox172.Text = row.Cells[4].Value.ToString();
                }
            }
        }
        private string getmanagerNum(string manager)
        {
            string num = "";
            if (manager=="Малий Е.В")
            {
                num="105";
            }
            if (manager == "Бровко Л.Ю")
            {
                num="121";
            }
            if (manager == "Дьякова Е.Е")
            {
                num="106";
            }
            if (manager == "Чумакова О.В")
            {
                num="107";
            }

            return num;
        }
        private void tabControl6_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_agent_invoice();
        }
        private void fill_agent_invoice()
        {
            //0
            label345.Text = "Счет № " + textBox179.Text;
            textBox212.Text = textBox201.Text +" "+ comboBox38.Text;
            textBox211.Text = textBox201.Text + " " + comboBox38.Text;
            dataGridView29.Rows[0].Cells[0].Value = "1";
            dataGridView29.Rows[0].Cells[1].Value = "Оплата путевки " + comboBox49.Text + ", " + textBox178.Text + " " + dateTimePicker13.Value.ToShortDateString() + "-" + dateTimePicker12.Value.ToShortDateString() ;
            if (dataGridView25.Rows[0].Cells[1].Value != null)
            {
                dataGridView29.Rows[0].Cells[1].Value += dataGridView25.Rows[0].Cells[1].Value.ToString() + "+" + (Convert.ToInt32(textBox171.Text) - 1).ToString();
            }
            dataGridView29.Rows[0].Cells[2].Value = "шт";
            dataGridView29.Rows[0].Cells[3].Value = textBox171.Text;
            textBox213.Text = "2";
            //dataGridView29.Rows[0].Cells[1].Value = "1";
            if (dataGridView27.Rows[0].Cells[7].Value != null)
            {
                dataGridView29.Rows[0].Cells[5].Value = (Convert.ToInt32(dataGridView27.Rows[0].Cells[8].Value.ToString()) + (Convert.ToInt32(textBox171.Text) * Convert.ToInt32(numericUpDown14.Value))).ToString();
            }
            //1
            dataGridView29.Rows[1].Cells[0].Value = "2";
            dataGridView29.Rows[1].Cells[1].Value = "Агентское вознаграждение ";// +textBox182.Text;
            /*if (dataGridView27.Rows[0].Cells[2].Value != null)
            {
                dataGridView29.Rows[1].Cells[5].Value = dataGridView27.Rows[0].Cells[2].Value.ToString();
            }*/
            dataGridView29.Rows[1].Cells[5].Value = (Convert.ToInt32(textBox171.Text) * Convert.ToInt32(numericUpDown14.Value)).ToString();
            if ((dataGridView27.Rows[0].Cells[2].Value != null) && (dataGridView27.Rows[0].Cells[7].Value != null))
            {
                dataGridView30.Rows[1].Cells[0].Value = Convert.ToInt32(dataGridView27.Rows[0].Cells[7].Value) - Convert.ToInt32(dataGridView27.Rows[0].Cells[2].Value);
                textBox214.Text = GetWordSum(dataGridView30.Rows[1].Cells[0].Value.ToString());
            }
             
        }
        private void setdefmanager(string name)
        {
            Section sec = new Section();
            sec.writekey("Defaultmanager", "name", "app.ini", "1");
        }

        private void textBox171_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {
            textBox171.Text = (numericUpDown12.Value + numericUpDown13.Value).ToString();
        }

        private void button52_Click(object sender, EventArgs e)
        {
            Section sec = new Section();
            if (comboBox52.Text != "")
            {
                textBox179.Text = sec.readkey("AgentDocCount", "number_" + comboBox52.Text, "app.ini");
            }
            else
            {
                MessageBox.Show("Выберите туроператора");
            }
        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string curryear = DateTime.Now.Year.ToString();
            DataGridView data = (DataGridView)sender;
            string currdate = ""; string currdatenew = "";
            if (e.ColumnIndex == 0)
            {
                if (data.Rows[e.RowIndex].Cells[0].Value != null)
                {
                    currdate = data.Rows[e.RowIndex].Cells[0].Value.ToString();
                    if (currdate.Length != 10)
                    {
                        currdatenew += currdate.Substring(0, 2) + ".";
                        currdatenew += currdate.Substring(3, 2) + ".20";
                        currdatenew += currdate.Substring(6, 2);
                        data.Rows[e.RowIndex].Cells[0].Value = currdatenew;
                    }

                }
            }
        }

        private void button62_Click(object sender, EventArgs e)
        {
            panel7.Visible = false;
        }

        private void списокЗаявокАгенствToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel7.Visible = true;
            panel7.BringToFront();
            List<string> managers = getmangerList("agent");
            foreach(string manager in managers)
            {
                comboBox53.Items.Add(manager);
            }
            List<string> countrys = getcountryList();
            foreach (string country in countrys)
            {
                comboBox54.Items.Add(country);
            }
            DateTime d = DateTime.Now.AddDays(-30);
            dateTimePicker17.Value = d;
            getagentdemandlist("", "", "", "");
        }
        private List<string> getmangerList(string type)
        {
            List<string> managers=new List<string>();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlCommand sqlcom = null; SqlDataReader reader;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            if (type == "agent")
            {
                query = "select name from managers where isagentmanager='1'";
            }
            else if (type == "client")
            {
                query = "select name from managers where isagentmanager<>'1'";
            }
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    while (reader.Read())
                    {
                        if (reader["name"] != null)
                            {
                                managers.Add(reader["name"].ToString());
                            }
                    }
                }

            }
            connect.Close();
            return managers; 
        }
       /* private List<string> getcountryList()
        {
            List<string> countrys = new List<string>();
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlCommand sqlcom = null; SqlDataReader reader;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "select Runame from country";
            
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    while (reader.Read())
                    {
                        if (reader["Runame"] != null)
                        {
                            countrys.Add(reader["Runame"].ToString());
                        }
                    }
                }

            }
            connect.Close();
            return countrys;
        }*/
        public string getTourInfo(string type, string tour)
        {
            string tourname="";
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "select name from tours where type_key='" + type + "' and id='" + tour+"'";
            SqlCommand sqlcom = null; SqlDataReader reader;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    while (reader.Read())
                    {
                        if (reader["name"] != null)
                        {
                            tourname=reader["name"].ToString();
                        }
                    }
                }

            }
            connect.Close();
            return tourname;
        }
        private void getagentdemandlist(string st_date,string end_date, string manager, string country)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = "";
            SqlCommand sqlcom = null;SqlDataReader  reader;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            int first=0;
            query = "select Agent_demand.id,Agent_demand.FirmName, Agent_demand.Client, Agent_demand.Phone,Agent_demand.Email,Agent_demand.Adult,Agent_demand.Child,Agent_demand.Path,Agent_demand.Days,Agent_demand.nights,convert(varchar,Agent_demand.startdate,105) as startdate ,convert(varchar,Agent_demand.enddate,105) as enddate,Agent_demand.Manager,convert(varchar,Agent_demand.DateCreate,105) as DateCreate, Country.Runame as country,managers.name as manager_name, managers.id as manager, Agent_demand.number from Agent_demand, managers, country where Agent_demand.country=country.id and Agent_demand.manager=managers.id";
            if ((st_date != "") ||(end_date != "") || (manager != "") || (country != ""))
            {
                //query+=" where ";
                if (st_date != "")
                {
                    /*if (first == 0)
                    {
                        query += "datecreate>='" + makeSQLdate(st_date,'.') + "'";
                        first = 1;
                    }
                    else
                    {*/
                query += " and Agent_demand.datecreate>='" + makeSQLdate(st_date, '.') + "'";
                    //}
                }
                if (end_date != "")
                {
                    /*if (first == 0)
                    {
                        query += "datecreate<='" + makeSQLdate(end_date,'.') + " 23:59:59'";
                        first = 1;
                    }
                    else
                    {*/
                    query += " and Agent_demand.datecreate<='" + makeSQLdate(end_date, '.') + " 23:59:59'";
                    // }
                }
                if (manager != "")
                {
                   /* if (first == 0)
                    {
                        query += "Manager in (select id from managers where name='" + manager + "')";
                        first = 1;
                    }
                    else
                    {*/
                query += " and Agent_demand.Manager in (select id from managers where name='" + manager + "')";
                    //}
                }
                if (country != "")
                {
                   /* if (first == 0)
                    {
                        query += "Country in (select id from country where Runame='" + country + "')";
                        first = 1;
                    }
                    else
                    {*/
                query += " and Agent_demand.Country in (select id from country where Runame='" + country + "')";
                    //}
                }
            }
            query += " order by Agent_demand.number desc";
            //FirmName, Client, Phone,Email,Adult,Child,Path,Days,nights,convert(varchar,startdate,105) ,convert(varchar,enddate,105),Manager,convert(varchar,DateCreate,105),Country";
            //query = "update agency set name='" + textBox207.Text + "', Agentname='" + textBox210.Text + "', City='" + comboBox51.Text + "', Phone='" + textBox208.Text + "', Email='" + textBox209.Text + "' where id='" + row.Cells[5].Value.ToString() + "'";
            int count = 0;
            connect.Open();
            {
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    dataGridView31.RowCount = 1;
                    if (reader.HasRows != false)
                    {
                        
                        while(reader.Read())
                        {
                            dataGridView31.Rows.Add();
                            if (reader["Number"] != null)
                            {
                                dataGridView31.Rows[count].Cells[0].Value = reader["Number"].ToString();
                            }
                            if (reader["DateCreate"] != null)
                            {
                                dataGridView31.Rows[count].Cells[1].Value = reader["DateCreate"].ToString();
                            }
                            if (reader["FirmName"] != null)
                            {
                                dataGridView31.Rows[count].Cells[2].Value = reader["FirmName"].ToString();
                            }
                            if (reader["id"] != null)
                            {
                                GetturistFromDemand((DataGridViewComboBoxCell)dataGridView31.Rows[count].Cells[3], reader["id"].ToString());
                               // dataGridView31.Rows[count].Cells[3]
                            }
                            if (reader["startdate"] != null)
                            {
                                dataGridView31.Rows[count].Cells[4].Value = reader["startdate"].ToString();
                            }
                            if (reader["enddate"] != null)
                            {
                                dataGridView31.Rows[count].Cells[5].Value = reader["enddate"].ToString();
                            }
                            if (reader["Path"] != null)
                            {
                                dataGridView31.Rows[count].Cells[6].Value = reader["Path"].ToString();
                            }
                            if (reader["Days"] != null)
                            {
                                dataGridView31.Rows[0].Cells[7].Value = reader["Days"].ToString();
                            }
                            if (reader["nights"] != null)
                            {
                                dataGridView31.Rows[0].Cells[8].Value = reader["nights"].ToString();
                            }
                            if (reader["Country"] != null)
                            {
                                dataGridView31.Rows[count].Cells[9].Value = reader["Country"].ToString();
                            }
                            if (reader["Manager_name"] != null)
                            {
                                dataGridView31.Rows[count].Cells[10].Value = reader["Manager_name"].ToString();
                            }
                            if (reader["id"] != null)
                            {
                                dataGridView31.Rows[count].Cells[11].Value = reader["id"].ToString();
                            }
                            /*if (reader["FirmName"]!=null)
                            {
                                dataGridView31.Rows[count].Cells[0].Value = reader["FirmName"].ToString();
                            }
                            if (reader["Client"] != null)
                            {
                                dataGridView31.Rows[count].Cells[1].Value = reader["Client"].ToString();
                            }
                            if (reader["Phone"] != null)
                            {
                                dataGridView31.Rows[count].Cells[2].Value = reader["Phone"].ToString();
                            }
                            if (reader["Email"] != null)
                            {
                                dataGridView31.Rows[count].Cells[3].Value = reader["Email"].ToString();
                            }
                            if (reader["Adult"] != null)
                            {
                                dataGridView31.Rows[count].Cells[4].Value = reader["Adult"].ToString();
                            }
                            if (reader["Child"] != null)
                            {
                                dataGridView31.Rows[count].Cells[5].Value = reader["Child"].ToString();
                            }
                            if (reader["Path"] != null)
                            {
                                dataGridView31.Rows[count].Cells[6].Value = reader["Path"].ToString();
                            }
                    
                            if (reader["startdate"] != null)
                            {
                                dataGridView31.Rows[0].Cells[0].Value = reader["startdate"].ToString();
                            }
                            if (reader["enddate"] != null)
                            {
                                dataGridView31.Rows[0].Cells[0].Value = reader["enddate"].ToString();
                            }
                            if (reader["Manager"] != null)
                            {
                                dataGridView31.Rows[0].Cells[0].Value = reader["Manager"].ToString();
                            }
                            if (reader["DateCreate"] != null)
                            {
                                dataGridView31.Rows[0].Cells[0].Value = reader["DateCreate"].ToString();
                            }
                            if (reader["Country"] != null)
                            {
                                dataGridView31.Rows[0].Cells[0].Value = reader["Country"].ToString();
                            }*/
                            count++;
                        }
                    }
                }
            }
            connect.Close();
        }

        private void button63_Click(object sender, EventArgs e)
        {
            groupBox24.Visible = true;
            panel7.Visible = false;
            //DataGridViewSelectedCellCollection cc=dataGridView31.SelectedCells[;
            if (dataGridView31.SelectedCells[0] != null)
            {
                
                DataGridViewCell cell = dataGridView31.SelectedCells[0];
                DataGridViewRow row = dataGridView31.Rows[cell.RowIndex];
                SqlConnectionStringBuilder connectStr = GetConnectSTR();
                string query = "";
                SqlCommand sqlcom = null; SqlDataReader reader;
                SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                query = "select Agent_demand.id,Agent_demand.FirmName, Agent_demand.Client, Agent_demand.Number, Agent_demand.Phone,Agent_demand.Email,Agent_demand.Adult,Agent_demand.Child,Agent_demand.Path,Agent_demand.Days,Agent_demand.nights,convert(varchar,Agent_demand.startdate,105) as startdate ,convert(varchar,Agent_demand.enddate,105) as enddate,Managers.name as manager,Managers.number as managernum,convert(varchar,Agent_demand.DateCreate,105) as DateCreate,Country.Runame as country,agency.city as ag_city,Agent_demand.number,Agent_demand.Transport,Agent_demand.Transfer,Agent_demand.Transport,Agent_demand.Tour,Agent_demand.TourType,Agent_demand.Hotel,Agent_demand.NomerType,Agent_demand.Categorytype,Agent_demand.FoodType,Agent_demand.Excursion, Agent_demand.Adv from Agent_demand, agency,country,managers where Agent_demand.country=country.id and Agent_demand.manager=managers.id and Agent_demand.agency_key=agency.id and Agent_demand.id='" + row.Cells[11].Value.ToString() + "'";
                connect.Open();
                {
                    if (connect.State == ConnectionState.Open)
                    {
                        sqlcom = new SqlCommand(query, connect);
                        reader = sqlcom.ExecuteReader();
                        if (reader.HasRows != false)
                        {
                            //dataGridView31.RowCount=1;
                            while (reader.Read())
                            {
                                if (reader["FirmName"] != null)
                                {
                                    textBox201.Text = reader["FirmName"].ToString();
                                }
                                if (reader["Client"] != null)
                                {
                                    textBox177.Text = reader["Client"].ToString();
                                }
                                if (reader["Phone"] != null)
                                {
                                    textBox175.Text = reader["Phone"].ToString();
                                }
                                if (reader["Email"] != null)
                                {
                                    textBox172.Text = reader["Email"].ToString();
                                }
                                if (reader["ag_city"] != null)
                                {
                                    comboBox38.Text = reader["ag_city"].ToString();
                                }
                                if (reader["Adult"] != null)
                                {
                                    numericUpDown12.Value = Convert.ToInt32(reader["Adult"]);
                                }
                                if (reader["Path"] != null)
                                {
                                    textBox178.Text = reader["Path"].ToString();
                                }
                                if (reader["Child"] != null)
                                {
                                    numericUpDown13.Value = Convert.ToInt32(reader["Child"]);
                                }
                                if (reader["Country"] != null)
                                {
                                    comboBox49.Text = reader["Country"].ToString();
                                }
                                if (reader["startdate"] != null)
                                {
                                    dateTimePicker13.Value = DateTime.Parse(reader["startdate"].ToString());
                                }
                                if (reader["enddate"] != null)
                                {
                                    dateTimePicker12.Value = DateTime.Parse(reader["enddate"].ToString());
                                }
                                if (reader["manager"] != null)
                                {
                                    comboBox41.SelectedItem = reader["manager"].ToString();
                                }
                                if (reader["Number"] != null)
                                {
                                    textBox197.Text = reader["Number"].ToString();
                                    textBox179.Text = reader["managernum"].ToString() + "-" + reader["Number"].ToString();
                                }

                                if (reader["Transport"] != null)
                                {
                                    comboBox55.SelectedItem = reader["Transport"].ToString();
                                }
                                if (reader["Transfer"] != null)
                                {
                                    comboBox39.SelectedItem = reader["Transfer"].ToString();
                                }
                                if (reader["Tour"] != null)
                                {
                                    comboBox48.Text = getTourInfo(reader["TourType"].ToString(),reader["Tour"].ToString());
                                }
                                if (reader["Hotel"] != null)
                                {
                                    textBox185.Text = reader["Hotel"].ToString();
                                }
                                if (reader["NomerType"] != null)
                                {
                                    comboBox46.Text = reader["NomerType"].ToString();
                                }
                                if (reader["Categorytype"] != null)
                                {
                                    comboBox45.Text = reader["Categorytype"].ToString();
                                }
                                if (reader["FoodType"] != null)
                                {
                                    comboBox40.Text = reader["FoodType"].ToString();
                                }
                                if (reader["Excursion"] != null)
                                {
                                    textBox189.Text = reader["Excursion"].ToString();
                                }
                                if (reader["Adv"] != null)
                                {
                                    textBox196.Text = reader["Adv"].ToString();
                                }
                            }
                
                        }
                        reader.Close();
                        query = "select FIO,PassportNum,Birthdate,PassportEnd from Agent_demand_turist where DemandKey='"+row.Cells[11].Value.ToString() + "'";
                        int count=0;
                        if (connect.State == ConnectionState.Open)
                        {
                            sqlcom = new SqlCommand(query, connect);
                            reader = sqlcom.ExecuteReader();
                            if (reader.HasRows != false)
                            {
                                dataGridView25.RowCount = 1;
                                while (reader.Read())
                                {
                                    dataGridView25.Rows.Add();
                                    dataGridView25.Rows[count].Cells[0].Value = count + 1;
                                    if (reader["FIO"] != null)
                                    {
                                        dataGridView25.Rows[count].Cells[1].Value= reader["FIO"].ToString();
                                    }
                                    if (reader["PassportNum"] != null)
                                    {
                                        dataGridView25.Rows[count].Cells[2].Value = reader["PassportNum"].ToString();
                                    }
                                    if (reader["Birthdate"] != null)
                                    {
                                        dataGridView25.Rows[count].Cells[3].Value = reader["Birthdate"].ToString();
                                    }
                                    if (reader["PassportEnd"] != null)
                                    {
                                        dataGridView25.Rows[count].Cells[4].Value = reader["PassportEnd"].ToString();
                                    }
                                    count++;
                                }
                            }
                        }
                    }
                    //connect.Close();
                }
                connect.Close();
            }
        }

        private void button64_Click(object sender, EventArgs e)
        {
            //dataGridView31.RowCount = 1;
            getagentdemandlist(dateTimePicker17.Value.ToShortDateString(), dateTimePicker18.Value.ToShortDateString(), comboBox53.Text, comboBox54.Text);
        }

        private void button65_Click(object sender, EventArgs e)
        {
            getagentdemandlist("", "", "", "");
        }

        private void формаПодтвержденияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox24.Visible = true;
        }

        private void папкаСПотдвеждениямиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Section sec = new Section();
            string path = Path.GetFullPath(sec.readkey("SavePath", "AgentPath", "app.ini"));
            if ((path != null) && (Directory.Exists(path)))
            {
                Process.Start(path);
            }
        }

        private void textBox11_Click(object sender, EventArgs e)
        {
            TextBox tobj= (TextBox)sender;
            if (tobj.Text == "")
            {
                tobj.Text = "+7";
            }
        }

        private void button66_Click(object sender, EventArgs e)
        {
            //DataGridViewCell cell = dataGridView32.SelectedCells[0];
            //DataGridViewRow row = dataGridView32.Rows[cell.RowIndex];
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = ""; int count=0;
            SqlCommand sqlcom = null; SqlDataReader reader;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            query = "select d.id, d.country, d.dogovornum, convert(varchar, d.dogovordate,104) as date, d.DogovorType, d.manager,d.email_yes,d.sms_yes,  cl.phone, cl.fio, cl.email,convert(varchar,cl.birthdate,104) as birthdate, cl.state_phone from clients as cl, dogovorinfo as d where cl.id=d.client and d.dogovordate>='" + makeSQLdate(dateTimePicker20.Text, '.') + "' and d.dogovordate<='" + makeSQLdate(dateTimePicker19.Text, '.') + "'";
            connect.Open();
            {
                
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        dataGridView32.RowCount=1;
                        while (reader.Read())
                        {
                            dataGridView32.Rows.Add();
                            //dataGridView32.Rows[count].Cells[0].Value = count + 1;
                            if (reader["date"] != null)
                            {
                                dataGridView32.Rows[count].Cells[0].Value = reader["date"].ToString();
                            }
                            if (reader["manager"] != null)
                            {
                                dataGridView32.Rows[count].Cells[1].Value = reader["manager"].ToString();
                            }
                            if (reader["fio"] != null)
                            {
                                dataGridView32.Rows[count].Cells[2].Value = reader["fio"].ToString();
                            }
                            if (reader["birthdate"] != null)
                            {
                                dataGridView32.Rows[count].Cells[3].Value = reader["birthdate"].ToString();
                            }
                            if (reader["Dogovornum"] != null)
                            {
                                dataGridView32.Rows[count].Cells[4].Value = reader["Dogovornum"].ToString();
                            }
                            if (reader["phone"] != null)
                            {
                                dataGridView32.Rows[count].Cells[5].Value = reader["phone"].ToString();
                            }
                            if (reader["state_phone"] != null)
                            {
                                dataGridView32.Rows[count].Cells[6].Value = reader["state_phone"].ToString();
                            }
                            if (reader["email"] != null)
                            {
                                dataGridView32.Rows[count].Cells[7].Value = reader["email"].ToString();
                            }
                            if (reader["email_yes"] != null)
                            {
                                dataGridView32.Rows[count].Cells[8].Value = reader["email_yes"].ToString();
                            }
                            if (reader["sms_yes"] != null)
                            {
                                dataGridView32.Rows[count].Cells[9].Value = reader["sms_yes"].ToString();
                            }
                            if (reader["country"] != null)
                            {
                                dataGridView32.Rows[count].Cells[10].Value = reader["country"].ToString();
                            }
                            if (reader["DogovorType"] != null)
                            {
                                dataGridView32.Rows[count].Cells[11].Value = reader["DogovorType"].ToString();
                            }
                            count++;
                        }
                    }
                }
            }
            connect.Close();
        }

        private void списокДоговоровToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel8.Visible = true;
        }

        private void button67_Click(object sender, EventArgs e)
        {
            panel8.Visible = false;
        }

        private void button68_Click(object sender, EventArgs e)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = ""; int count = 0;
            SqlCommand sqlcom = null; SqlDataReader reader;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            //query = "select id,fio, birthdate, phone, email, state_phone from clients where month(birthdate)=month('"+ makeSQLdate(dateTimePicker22.Text, '.') + "') and day(birthday)=day('"+ makeSQLdate(dateTimePicker22.Text, '.') +"')";
            query = "use rosintour exec GetbirthDates '"+textBox219.Text+"'";
            connect.Open();
            {

                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        dataGridView33.RowCount = 1;
                        while (reader.Read())
                        {
                            dataGridView33.Rows.Add();
                            //dataGridView32.Rows[count].Cells[0].Value = count + 1;
                            if (reader["birthdate"] != null)
                            {
                                dataGridView33.Rows[count].Cells[0].Value = reader["birthdate"].ToString();
                            }
                            if (reader["fio"] != null)
                            {
                                dataGridView33.Rows[count].Cells[1].Value = reader["fio"].ToString();
                            }
                            if (reader["phone"] != null)
                            {
                                dataGridView33.Rows[count].Cells[2].Value = reader["phone"].ToString();
                            }
                            if (reader["state_phone"] != null)
                            {
                                dataGridView33.Rows[count].Cells[3].Value = reader["state_phone"].ToString();
                            }
                            if (reader["email"] != null)
                            {
                                dataGridView33.Rows[count].Cells[4].Value = reader["email"].ToString();
                            }
                            count++;
                        }
                    }
                }
                //reader.Close();
            }
            connect.Close();
        }

        private void button69_Click(object sender, EventArgs e)
        {
            SqlConnectionStringBuilder connectStr = GetConnectSTR();
            string query = ""; int count = 0; string query1 = ""; ArrayList c = new ArrayList(); ArrayList c1= new ArrayList();
            SqlCommand sqlcom = null; SqlDataReader reader; SqlCommand sqlcom1 = null;
            SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
            SqlConnection connect1 = new SqlConnection(connectStr.ConnectionString);
            query = "select birthdate,id from Clients";
            connect.Open();
            {

                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);
                    reader = sqlcom.ExecuteReader();
                    while (reader.Read())
                    {
                        if (reader["id"] != null)
                        {
                            //c1.Add(reader["id"].ToString()); c1.Add(reader["birthdate"].ToString());
                            //c.Add(c1);
                            connect1.Open();
                            {
                                query1 = "update clients set birthdate1='" + makeSQLdate(reader["birthdate"].ToString(), '.') + "' where id='" + reader["id"].ToString() + "'";
                                sqlcom1 = new SqlCommand(query1, connect1);
                                sqlcom1.ExecuteNonQuery();
                            }
                            connect1.Close();
                            // c1.Clear();
                        }
                    }
                }
            }
            connect.Close(); MessageBox.Show("end");
        }

        private void button69_Click_1(object sender, EventArgs e)
        {
            panel9.Visible = false;
        }

        private void дниРожденияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel9.Visible = true;
        }

        private void button70_Click(object sender, EventArgs e)
        {
            string[] birtharr;
            string str = "ФИО,E-mail,День рожд.,Компания,Должность,Телефон (дом.),Телефон (раб.),Адрес (дом.),Город (дом.),Штат (дом.),Индекс (дом.),Страна (дом.),Адрес (раб.),Город (раб.),Штат (раб.),Индекс (раб.),Страна (раб.),Заметки,Префикс,Суффикс\r\n";
            for (int i = 0; i < dataGridView32.RowCount-1; i++)
            {
                //str = "";
                if ((dataGridView32.Rows[i].Cells[7].Value.ToString() != "")&&(dataGridView32.Rows[i].Cells[8].Value.ToString()=="Да"))
                {
                    str += dataGridView32.Rows[i].Cells[2].Value.ToString() + ",";
                    str += dataGridView32.Rows[i].Cells[7].Value.ToString() + ",";
                    birtharr=dataGridView32.Rows[i].Cells[3].Value.ToString().Split('.');
                    str += birtharr[2] + birtharr[1] + birtharr[0] + ",,,";
                    str += dataGridView32.Rows[i].Cells[5].Value.ToString() + ",,,,,,,,,,,,\0D\0A,,\r\n";
                }
            }
           // FileStream fs = new FileStream("c:\\Variables.csv", FileMode.Append, FileAccess.Write, FileShare.Write);
           // byte[] info = new UTF8Encoding(true).GetBytes(str);
           // fs.Write(info, 0, info.Length);
            //fs.Close();
            string path = "";
            if (!Directory.Exists(Basepath + "Email_контакты_выгрузка"))
            {
                Directory.CreateDirectory(Basepath + "Email_контакты_выгрузка");
                path = Basepath + "Email_контакты_выгрузка";
            }
            else
            {
                path = Basepath + "Email_контакты_выгрузка";
            }
            StreamWriter sw = new StreamWriter(path+"\\"+dateTimePicker20.Value.ToShortDateString()+" - "+dateTimePicker19.Value.ToShortDateString()+ "email_contacs.csv", true, Encoding.Default);
            //string NextLine = "This is the appended line.";
            sw.Write(str);
            sw.Close();
        }
        private void Save_Dogovor(DogovorInfo dog)
        {
            string query = null;
            if (dog.id != null)
            {
                query = "update DogovorInfo set";
            }
            else
            {
                query = "insert into DogovorInfo values()";
            }
            if (query!=null)
            {
                try
                {
                    SqlConnectionStringBuilder connectStr = GetConnectSTR();
                    SqlCommand sqlcom = null; SqlDataReader reader = null;
                    SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                    connect.Open();
                    {
                        if (connect.State == ConnectionState.Open)
                        {
                            sqlcom = new SqlCommand(query, connect);
                            sqlcom.ExecuteNonQuery();
                        }
                    }
                    connect.Close();
                }
                catch
                {
                    erorrFSave("error.txt", query);
                }
            }
        }
        private void button71_Click(object sender, EventArgs e)
        {

            Save_Dogovor(dogovor);
        }


        //Confirmation
        
        //Konsul pasp-dateCH
    }
    public class SenderObj
    {
        public object owner = null;
        public void Set(object o1)
        {
            this.owner = o1;
        }
        public void clean()
        {
            this.owner = null;
        }
    }

    public class Dataview
    {
        public object owner = null;
        public DataGridViewCellEventArgs args = null;
        public void Set(object o1, DataGridViewCellEventArgs args)
        {
            this.owner = o1;
            this.args = args;
        }
        public void clean()
        {
            this.owner = null;
            this.args = null;
        }

    }
    public class Client
    {
        public Client()
        {
        FIO = null;
        RUpaspSeriy=null;
        RUpaspnum=null;
        RUpaspDate=null;
        RUpaspOwn=null;
        ENpaspSeriy=null;
        ENpaspnum=null;
        ENpaspDate=null;
        ENpaspOwn=null;
        Birthdate=null;
        Email=null;
        Phone=null;
        Skype=null;
        Adress = null;
        ICQ = null;
        }
        public Client(string fio, string rupaspSeriy, string rupaspnum, string rupaspDate, string rupaspOwn, string enpaspSeriy, string enpaspnum, string enpaspDate, string enpaspOwn, string birthdate, string email, string phone, string skype, string adress, string icq, string state_phone)
        {
            FIO = fio;
            RUpaspSeriy = rupaspSeriy;
            RUpaspnum = rupaspnum;
            RUpaspDate = rupaspDate;
            RUpaspOwn = rupaspOwn;
            ENpaspSeriy = enpaspSeriy;
            ENpaspnum = enpaspnum;
            ENpaspDate = enpaspDate;
            ENpaspOwn = enpaspOwn;
            Birthdate = birthdate;
            Email = email;
            Phone = phone;
            Skype = skype;
            Adress = adress;
            ICQ = icq;
            State_phone = state_phone;

        }
        public string FIO;
        public string RUpaspSeriy;
        public string RUpaspnum;
        public string RUpaspDate;
        public string RUpaspOwn;
        public string ENpaspSeriy;
        public string ENpaspnum;
        public string ENpaspDate;
        public string ENpaspOwn;
        public string Birthdate;
        public string Email;
        public string Phone;
        public string ICQ;
        public string Skype;
        public string Adress;
        public string State_phone;

    }
    public class ConfirInfo
    {
        public string Demand_key;
        public string Path;
        public string Adult;
        public string Child;
        public string Startdate;
        public string EndDate;
        public string Hotel;
        public string Tour;
        public string Manager;
        public string FullSum;
        public string AgentSum;
        public string AgentProcent;
        public string FinalSum;
        public string Partner;
        public string Agentkey;
        public string Country;
        public string Touroperator;

        public ConfirInfo()
        {
            Demand_key = null;
            Path = null;
            Adult = null;
            Child = null;
            Startdate = null;
            EndDate = null;
            Hotel = null;
            Tour = null;
            Manager = null;
            FullSum = null;
            AgentSum = null;
            AgentProcent = null;
            FinalSum = null;
            Partner = null;
            Agentkey = null;
            Country = null;
            Touroperator = null;
        }
        public ConfirInfo(string demand_key, string path,string adult,string child,string startdate,string enddate,string hotel,string tour,string manager,string fullsum,string agentsum,string agentProcent, string finalsum,string partner,string agentkey,string country,string touroperator)
        {
            Demand_key = demand_key;
            Path = path;
            Adult = adult;
            Child = child;
            Startdate = startdate;
            EndDate = enddate;
            Hotel = hotel;
            Tour = tour;
            Manager = manager;
            FullSum = fullsum;
            AgentSum = agentsum;
            AgentProcent = agentProcent;
            FinalSum = finalsum;
            Partner = partner;
            Agentkey = agentkey;
            Country = country;
            Touroperator = touroperator;
        }
    }
    public class DogovorInfo
    {
        public DogovorInfo()
        {
        id = null;
        DogovorDate = null;
        //public string Client;
        TyrName=null;
        TravelPath = null; 
        StartDate = null; 
        EndDate = null; 
        Hotel = null; 
        PayType = null; 
        Currency = null; 
        Course = null; 
        RUPrice = null; 
        ENPrice = null; 
        DogovorType = null; 
        Manager = null; 
        Tyroperator = null;
        clientID = null;
        Discount = null;
        Country = null;

        
        }
        public DogovorInfo(string num, string date, string tyrname, string travelpath, string startdate, string enddate, string hotel, string paytype, string currency, string course, string ruprice, string enprice, string type, string manager, string tyroperator, string id, string country, string discount, string sms_yes, string email_yes, string cardnum)
        {
            Dogovornum = num;
            DogovorDate = date;
            clientID=id;
            TyrName = tyrname;
            TravelPath = travelpath; 
            StartDate = startdate; 
            EndDate = enddate; 
            Hotel = hotel; 
            PayType = paytype; 
            Currency = currency; 
            Course = course; 
            RUPrice = ruprice; 
            ENPrice = enprice; 
            DogovorType = type; 
            Manager = manager; 
            Tyroperator = tyroperator;
            Country = country;
            Discount = discount;
            Sms_yes = sms_yes;
            Email_yes = email_yes;
            CardNum = cardnum;
            
        }
        public string id;
        public string Dogovornum;
        public string DogovorDate;
        public string clientID;
        public string TyrName;
        public string TravelPath;
        public string StartDate;
        public string EndDate;
        public string Hotel;
        public string PayType;
        public string Currency;
        public string Course;
        public string RUPrice;
        public string ENPrice;
        public string DogovorType;
        public string Manager;
        public string Tyroperator;
        public string Country;
        public string Discount;
        public string Sms_yes;
        public string Email_yes;
        public string CardNum;
        
    }
    public class AviaDogovorInfo
    {
        public AviaDogovorInfo()
        {
         Dogovornum=null;
         DogovorDate = null;
         clientID = null;
         Manager = null;
         Company = null;
         AgentDogNum = null;
         AgentDogDate = null;
         FIO = null;
         Adress = null;
         Phone = null;
         Country = null;
         TravelPath = null;
        }
        public AviaDogovorInfo(string dogovornum_a,string dogovorDate_a,string clientID_a,string manager_a,string company_a,string agentDogNum_a,string agentDogDate_a,string FIO_a,string adress_a,string phone_a,string country, string travelpath)
        {
            Dogovornum = dogovornum_a;
            DogovorDate = dogovorDate_a;
            clientID = clientID_a;
            Manager = manager_a;
            Company = company_a;
            AgentDogNum = agentDogNum_a;
            AgentDogDate = agentDogDate_a;
            FIO = FIO_a;
            Adress = adress_a;
            Phone = phone_a;
            Country = country;
            TravelPath = travelpath;
        }
        public string Dogovornum;
        public string DogovorDate;
        public string clientID;
        public string Manager;
        public string Company;
        public string AgentDogNum;
        public string AgentDogDate;
        public string FIO;
        public string Adress;
        public string Phone;
        public string Country;
        public string TravelPath;
    }
    public class FlightInfo
    {
        public FlightInfo()
        {
        //Country=null;
        Date = null;
        Time = null;
        FlightNum = null;
        StartCity = null;
        EndCity = null;
        Tariff = null;
        Hotel = null;
        FIO = null;
        clientID = null;
        Mannum = null;
        DogovorNum = null;
       //dInfoKey = null;
        }
        public FlightInfo(string dogovornum, string date, string time, string flightnum, string startcity, string endcity,string mannum, string tariff, string hotel, string fio, string id)
        {
           // Country = country;
            Date = date;
            Time = time;
            FlightNum = flightnum;
            StartCity = startcity;
            EndCity = endcity;
            Tariff = tariff;
            Hotel = hotel;
            FIO = fio;
            clientID = id;
            Mannum = mannum;
            DogovorNum = dogovornum;

        }

        //public string Country;
        public string Date;
        public string Time;
        public string FlightNum;
        public string StartCity;
        public string EndCity;
        public string Tariff;
        public string Hotel;
        public string FIO;
        public string clientID;
        public string Mannum;
        public string DogovorNum;
        //public string dInfoKey;
    }
    public class ManagerInfo
    {
        public  ManagerInfo()
        {
         this.name=null;
         this.number=null;
         this.email=null;
         this.icq=null;
         this.skype=null;
         this.phone=null;
        }
        public string name;
        public string number;
        public string email;
        public string icq;
        public string skype;
        public string phone;
    }
    public class Touroperator
    {
        public Touroperator()
        {
            this.to_id = null;
            this.to_shortname = null;
            this.to_name = null;
            this.to_adress = null;
            this.to_rn = null;
            this.to_inn = null;
            this.to_ogrn = null;
            this.to_tel = null;
            this.to_fax = null;
            this.ins_name = null;
            this.ins_adress = null;
            this.ins_d_num = null;
            this.ins_fin_cap = null;
            this.ins_d_date = null;
            this.ins_d_sdate = null;
            this.ins_d_edate = null;
        }
        public void getinfo(SqlConnectionStringBuilder constr, string name)
        {
            SqlConnectionStringBuilder connectStr = constr;
            SqlDataReader reader;
            SqlCommand sqlcom = null;
            string query = "select * from touroperators where to_shortname='"+name+"'";
            try
            {
                SqlConnection connect = new SqlConnection(connectStr.ConnectionString);
                connect.Open();
                if (connect.State == ConnectionState.Open)
                {
                    sqlcom = new SqlCommand(query, connect);

                    reader = sqlcom.ExecuteReader();
                    if (reader.HasRows != false)
                    {
                        while (reader.Read())
                        {
                            this.to_id = reader["id"].ToString();
                            this.to_shortname = reader["to_shortname"].ToString();
                            this.to_name = reader["to_name"].ToString();
                            this.to_adress = reader["to_adress"].ToString();
                            this.to_rn = reader["to_rn"].ToString();
                            this.to_inn = reader["to_inn"].ToString();
                            this.to_ogrn = reader["to_ogrn"].ToString();
                            this.to_tel = reader["to_tel"].ToString();
                            this.to_fax = reader["to_fax"].ToString();
                            this.ins_name = reader["ins_name"].ToString();
                            this.ins_adress = reader["ins_adress"].ToString();
                            this.ins_d_num = reader["ins_d_num"].ToString();
                            this.ins_fin_cap = reader["ins_fin_cap"].ToString();
                            this.ins_d_date = reader["ins_d_date"].ToString();
                            this.ins_d_sdate = reader["ins_d_sdate"].ToString();
                            this.ins_d_edate = reader["ins_d_edate"].ToString();
                        }
                    }
                }
                connect.Close();
            }
            catch
            {

            }
        }
        public string to_id;
        public string to_shortname;
        public string to_name;
        public string to_adress;
        public string to_rn;
        public string to_inn;
        public string to_ogrn;
        public string to_tel;
        public string to_fax;
        public string ins_name;
        public string ins_adress;
        public string ins_d_num;
        public string ins_fin_cap;
        public string ins_d_date;
        public string ins_d_sdate;
        public string ins_d_edate;
    }
    public class Arg
    {
        public Arg()
        {
            this.Doc = null;
            this.Word = null;
            this.App = null;
            this.Excel = null;
            this.WorkBook = null;
        }
        public void setparam(object o1, object o2, object o3)
        {
            this.Doc = o1;
            this.Word = o2;
            this.App = o3;
        }
        public void setparamE(object o1, object o2)
        {
            this.Excel = o1;
            this.WorkBook = o2;
        }
        public void clean()
        {
            this.Doc = null;
            this.Word = null;
            this.App = null;
            this.Excel = null;
            this.WorkBook = null;
        }
        public object Doc;
        public object Word;
        public object App;
        public object Excel;
        public object WorkBook;

    }
}