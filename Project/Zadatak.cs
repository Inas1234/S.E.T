using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Project
{
    public partial class Zadatak : Form
    {
        myDatabase con = new myDatabase();

        public string ImeZadatka{ set; get; }
        public string Serial { get; set; }
        public DateTime timeP { get; set; }
        public string EST;
        public string opis { set; get; }
        string ass = "";
        string balls = "";
        string gaydin = "";
        DateTime time;
        string TimeP;
        public Zadatak()
        {
            InitializeComponent();
            con.Connect();
        }

        private void Zadatak_Load(object sender, EventArgs e)
        {
            balls = ImeZadatka;
            ass = EST;
            gaydin = opis;
            label1.Text = balls;
            label2.Text = ass;
            label3.Text = gaydin;
        }

        private void Zadatak_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                con.cn.Open();
                MySqlCommand cmd2 = new MySqlCommand();
                cmd2.CommandText = "SELECT vrijeme_pocetka FROM tasks WHERE task_name = '" + balls + "'";
                cmd2.Connection = con.cn;
                MySqlDataReader sdr = cmd2.ExecuteReader();
                while (sdr.Read())
                {
                    TimeP = sdr["vrijeme_pocetka"].ToString();
                    DateTime TimeL = Convert.ToDateTime(TimeP);
                    time = DateTime.Now;
                    TimeSpan dateU = time.Subtract(TimeL);
                    DateTime DateU = DateTime.Today + dateU;
                    string sex = new DateTime(dateU.Ticks).ToString("HH:mm:ss");

                    DateTime TimePO = Convert.ToDateTime(ass);
                    DateTime TIMEPOU = TimePO.Subtract(dateU);
                    long timpou = 864000000000 - TIMEPOU.TimeOfDay.Ticks;
                    TimeSpan deezNuts = new TimeSpan(timpou);
                    string sex2 = new DateTime(deezNuts.Ticks).ToString("HH:mm:ss");

                    if (DateU > TimePO)
                    {
                        sdr.Close();
                        MySqlCommand cmd = new MySqlCommand("UPDATE tasks SET vrijeme_kraja ='" + time.ToString("HH:mm:ss") + "', ukupno_vrijeme_rada = '" + sex + "', uradjeno = 'YES', podbacaj = '" + sex2 + "' WHERE task_name ='" + balls + "'", con.cn);
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        sdr.Close();
                        MySqlCommand cmd = new MySqlCommand("UPDATE tasks SET vrijeme_kraja ='" + time.ToString("HH:mm:ss") + "', ukupno_vrijeme_rada = '" + sex + "', uradjeno = 'YES', prebacaj = '" + TIMEPOU.ToString("HH:mm:ss") + "' WHERE task_name ='" + balls + "'", con.cn);
                        cmd.ExecuteNonQuery();
                    }
                    Main main = new Main();
                    this.Hide();
                    main.ShowDialog();
                }
                
                con.cn.Close();
            }
            catch(Exception ex)
            {

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Main main = new Main();
            this.Hide();
            main.ShowDialog();
        }

    
    }
}
