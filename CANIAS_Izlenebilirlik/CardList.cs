using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO.Ports;

namespace CANIAS_Izlenebilirlik
{
    public partial class CardList : Form
    {
        DataTable dataTable;
        Form1 form1;
        Boolean boolean = true;
        private static bool portSituation;
        public CardList()
        {
            InitializeComponent();
        }

        private void CardList_Load(object sender, EventArgs e)
        {
            getList();
            Control.CheckForIllegalCrossThreadCalls = false;
            form1 = new Form1();
            foreach (var item in SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(item);
            }
            txtPublicFolder.Text = Setting.Default.publicFolder;
            checkBox1.Checked = Setting.Default.publicFolderUsingSit;
            comboBox1.Text = Setting.Default.portNameSet;
            portSituation = Setting.Default.portNameSituation;
            if (portSituation)
            {
                btnOpenSerialPortAlarm.BackColor = Color.Olive;
                btnOpenSerialPortAlarm.Text = "ON";
            }
            else
            {
                btnOpenSerialPortAlarm.BackColor = Color.DarkRed;
                btnOpenSerialPortAlarm.Text = "OFF";
            }

        }
        private void getList()
        {
            string server = @"192.168.10.22";
            string database = "ALP802";
            string user = "otomasyon";
            string pass = "123KUM*";
            String connection = @"Data Source=" + server + ";Initial Catalog=" + database + ";User ID=" + user + ";Password=" + pass;
            String command = "select distinct(m.STEXT),b.MATERIAL from IASMATBASIC b " +
            "inner join IASMATX m on b.MATERIAL = m.MATERIAL AND m.TEXTTYPE = 'M' " +
            "where m.PLANT = '20' order by b.MATERIAL";
            SqlConnection sqlConnection = new SqlConnection(connection);
            sqlConnection.Open();
            dataTable = new DataTable();
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, sqlConnection);
            sqlDataAdapter.Fill(dataTable);
            dataGridViewCL.DataSource = dataTable;
            sqlConnection.Close();
        }

        private void txtCode_TextChanged(object sender, EventArgs e)
        {
            DataView dataView = dataTable.DefaultView;
            dataView.RowFilter = string.Format("MATERIAL like '%{0}%'", txtCode.Text);
            dataGridViewCL.DataSource = dataView.ToTable();
        }
        private void txtMaterialName_TextChanged(object sender, EventArgs e)
        {
            DataView dataView = dataTable.DefaultView;
            dataView.RowFilter = string.Format("STEXT like '%{0}%'", txtMaterialName.Text);
            dataGridViewCL.DataSource = dataView.ToTable();
        }

        private void CardList_FormClosed(object sender, FormClosedEventArgs e)
        {
            form1.isCardListOpened = false;
        }
     
        private void btnOpenSerialPortAlarm_Click(object sender, EventArgs e)
        {
           
            if (portSituation)
            {
                btnOpenSerialPortAlarm.BackColor = Color.DarkRed;
                btnOpenSerialPortAlarm.Text = "OFF";
                Setting.Default.portNameSituation = false;
                Setting.Default.portNameSet = comboBox1.SelectedItem.ToString();
                portSituation = false;
  
            }
            else
            {
                if (comboBox1.Text != "")
                {

                    Setting.Default.portNameSituation = true;
                    portSituation = true;
                    btnOpenSerialPortAlarm.BackColor = Color.Olive;
                    Setting.Default.portNameSet = comboBox1.SelectedItem.ToString();
                    MessageBox.Show(Setting.Default.portNameSet);
                    btnOpenSerialPortAlarm.Text = "ON";
                    MessageBox.Show("AKTİF !");
                }
                else
                {
                    MessageBox.Show("Port Seçiniz !");
                }
            }

            Setting.Default.Save();
        }
    
        string data = "null1";
        private void btnWarningLamb_Click(object sender, EventArgs e)
        {
            if (boolean)
            {
                visibleTrue(boolean);
                boolean = false;
            }
            else
            {
                visibleTrue(boolean);
                boolean = true;
            }
          
        }
        private void visibleTrue(Boolean boolean)
        {
            label3.Visible = boolean;
            comboBox1.Visible = boolean;
            btnOpenSerialPortAlarm.Visible = boolean;
            btnSave.Visible = boolean;
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            Setting.Default.publicFolder = txtPublicFolder.Text;
            Setting.Default.portNameSet = comboBox1.SelectedItem.ToString();
            Setting.Default.Save();
        }

        private void CardList_FormClosing(object sender, FormClosingEventArgs e)
        {
            Setting.Default.Save();
        }

        private void dataGridViewCL_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridViewCL.CurrentRow.Selected = true;
            Form1.materialCodeCopyData = dataGridViewCL.Rows[e.RowIndex].Cells[1].Value.ToString();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                Setting.Default.publicFolderUsingSit = true;
            }
            else
            {
                Setting.Default.publicFolderUsingSit = false;
            }
            Setting.Default.Save();        }
    }
}
