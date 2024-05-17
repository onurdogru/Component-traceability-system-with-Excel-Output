using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Sheet = Microsoft.Office.Interop.Excel.Workbook;
using WorkSheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace CANIAS_Izlenebilirlik
{
    public partial class Form1 : Form
    {
        SqlCommand sqlCommand;
        SqlConnection sqlConnection;
        SqlDataReader sqlDataReader;
        int rowNumberExcel=0,indexRow = 7, constNumber = 17, indexView = 0, indexExcelCell = 0, indexComponentRow, timeBarcode = 0, timerMess = 0, tempEpsCodeIndex, firstEps;
        float dataNumber, excelPart = 0;
        string referenceFileURL, excelFail, finishedFile, t20ExcelName, date, excelFile, barcodeData, rCode, epsCode, quantity, lotNo, templotNo, tempEpsCode, changeMessage, productName, productNo, productDate;
        ExcelApp excel;
        Sheet sheet;
        WorkSheet worksheet;
        Excel.Range userRange, userPrCode;
        bool flagMatch =false,flagRecover = false, flagProduct = true, isOpened = false, changeFlag, searchAlgorithm, searchAlgorithmLot, searchAlgorithmEps, addComponent2Flag = false;
        public bool isCardListOpened = false;
        List<string> top_code_list;
        List<string> material_code_list;
        List<string> description_list;
        List<string> rcode_list;
        List<DataGridView> dataGridView_list, dataGridView_ExtraList;
        public static SerialPort serialPortO;
        CardList cardList;
        DataGridView dataGridView5, dataGridView6, dataGridView7, dataGridView8, dataGridView9, dataGridView10, dataGridView11, dataGridView12,
        dataGridView13, dataGridView14, dataGridView15, dataGridView16;
        public static string materialCodeCopyData;
        public static int dataGridViewRowCount,constRowNumber=0;
        int rowX = 0;
        int constNumberPage = 7;
        int constNumberIndex = 16;
        int coefficientNumber = 0;
        int headIndex = 0;
        static int indexView2 = 0;
        static int indexView3 = 0;
        string letData = "";
        private Object thisLock = new Object();
        byte[] byteArray = new byte[10];
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           // Control.CheckForIllegalCrossThreadCalls = false;
            openSerialPort();

            dataAccessProccess();
            setDataGridViewCss(dataGridViewMaterial, dataGridView1, dataGridView2, dataGridView3, dataGridView4);
            disableButton(3, false, btnFirstRecord, btnChange, btnDone);
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.btnRecover, "Geri yüklenecek excel dosyasını seçin");

        }
        private void openSerialPort()
        {
            try
            {
                if (Setting.Default.portNameSituation)
                {
                    serialPortAlarm.PortName = Setting.Default.portNameSet;
                    serialPortAlarm.Open();
                    Setting.Default.portNameSituation = true;
                    Setting.Default.Save();
                }
                else
                {
                    MessageBox.Show("İkaz Cihazının Kapalı !");
                }
           
            }
            catch (Exception)
            {

                MessageBox.Show("İkaz Cihazının Portu Bulunumadı !");
            }

        }
        private void setDefaultToParameter()
        {
            constNumberPage = 7;
            constNumberIndex = 16;
            coefficientNumber = 0;
            headIndex = 0;
        }
        private void extraDataGridViewListCall(params dynamic[] parameter)
        {
            dataGridView_ExtraList = new List<DataGridView>();
            int m = 0;
            for (int i = 0; i < parameter[0]; i++)
            {
                parameter[i + 1] = new DataGridView();
                dataGridView_ExtraList.Add(parameter[i + 1]);
            }
        }

        private void setDataGridViewCss(params DataGridView[] dataGridView)
        {
            pictureBox1.Padding = new Padding(32);
            dataGridView[0].EnableHeadersVisualStyles = false;
            dataGridView[0].ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            for (int i = 0; i < 3; i++)
            {
                dataGridView[0].Columns[i].HeaderCell.Style.Font = new Font("Tahoma", 10, FontStyle.Bold);
                dataGridView[0].Columns[i].HeaderCell.Style.ForeColor = Color.White;
                dataGridView[0].Columns[i].HeaderCell.Style.BackColor = Color.Green;
                dataGridView[0].Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 10, FontStyle.Bold);
                dataGridView[0].Columns[3].HeaderCell.Style.ForeColor = Color.White;
                dataGridView[0].Columns[3].HeaderCell.Style.BackColor = Color.Green;
                dataGridView[i + 1].EnableHeadersVisualStyles = false;
                dataGridView[i + 1].ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView[4].EnableHeadersVisualStyles = false;
                dataGridView[4].ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            for (int i = 0; i < 3; i++)
            {
                dataGridView[i + 1].Size = new Size(tableLayoutPanel7.Width, 237);
                dataGridView[1].Columns[i].HeaderCell.Style.Font = new Font("Tahoma", 9, FontStyle.Bold);
                dataGridView[1].Columns[i].HeaderCell.Style.ForeColor = Color.White;
                dataGridView[1].Columns[i].HeaderCell.Style.BackColor = Color.DarkSlateBlue;
                dataGridView[2].Columns[i].HeaderCell.Style.Font = new Font("Tahoma", 9, FontStyle.Bold);
                dataGridView[2].Columns[i].HeaderCell.Style.ForeColor = Color.White;
                dataGridView[2].Columns[i].HeaderCell.Style.BackColor = Color.DarkSlateBlue;
                dataGridView[3].Columns[i].HeaderCell.Style.Font = new Font("Tahoma", 9, FontStyle.Bold);
                dataGridView[3].Columns[i].HeaderCell.Style.ForeColor = Color.White;
                dataGridView[3].Columns[i].HeaderCell.Style.BackColor = Color.DarkSlateBlue;
                dataGridView[4].Columns[i].HeaderCell.Style.Font = new Font("Tahoma", 9, FontStyle.Bold);
                dataGridView[4].Columns[i].HeaderCell.Style.ForeColor = Color.White;
                dataGridView[4].Columns[i].HeaderCell.Style.BackColor = Color.DarkSlateBlue;
            }


            extraDataGridViewListCall(12, dataGridView5, dataGridView6, dataGridView7, dataGridView8, dataGridView9, dataGridView10, dataGridView11, dataGridView12, dataGridView13, dataGridView14, dataGridView15, dataGridView16);

            int top = 50;
            int left = 50;
            int m = 0;
            for (int i = 0; i < 12; i++)
            {

                dataGridView_ExtraList[i].Columns.Add("colName" + i.ToString(), "No");
                dataGridView_ExtraList[i].Columns.Add("colName" + i + 1.ToString(), "LotNo");
               
                dataGridView_ExtraList[i].Columns.Add("colName" + i + 2.ToString(), "Adet");
                dataGridView_ExtraList[i].EnableHeadersVisualStyles = false;
                dataGridView_ExtraList[i].ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView_ExtraList[i].ColumnHeadersDefaultCellStyle.BackColor = Color.DarkSlateBlue;
                dataGridView_ExtraList[i].ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_ExtraList[i].ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9, FontStyle.Bold);
                dataGridView_ExtraList[i].RowHeadersVisible = false;
                dataGridView_ExtraList[i].ColumnHeadersHeight = 18;
                dataGridView_ExtraList[i].AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView_ExtraList[i].AllowUserToAddRows = false;
                dataGridView_ExtraList[i].Left = left;
                dataGridView_ExtraList[i].Top = top;
                dataGridView_ExtraList[i].Size = new Size(dataGridView1.Width, tableLayoutPanel7.Height - 2);
                dataGridView_ExtraList[i].AllowUserToResizeRows = false;
                this.tableLayoutPanel7.Controls.Add(dataGridView_ExtraList[i]);
                left += dataGridView_ExtraList[i].Height + 60;
                dataGridView_ExtraList[i].ScrollBars = ScrollBars.None;
                dataGridView_ExtraList[i].Enabled = false;
            }

            for (int i = 0; i < 12; i++)
            {
                dataGridView_ExtraList[i].Columns[0].Width = 40;
                dataGridView_ExtraList[i].Columns[1].Width = 100;
            }

        }
        private bool dataAccessProccess()
        {

            string server = @"192.168.10.22";
            string database = "ALP802";
            string user = "otomasyon";
            string pass = "123KUM*";
            String connection = @"Data Source=" + server + ";Initial Catalog=" + database + ";User ID=" + user + ";Password=" + pass;
            sqlConnection = new SqlConnection(connection);
            try
            {
                sqlConnection.Open();
                return true;
            }
            catch (Exception)
            {
                lblProductMessage.Text = "Veritabanı bağlantısı kurulamadı !";
                return false;
            }

        }
        // parameter[0] : number of length for
        // parameter[1] : elements of dataGridViewList 

        private void dataGridViewListCall(params dynamic[] parameter)
        {
            dataGridView_list = new List<DataGridView>();
            for (int i = 0; i < parameter[0]; i++)
            {
                dataGridView_list.Add(parameter[i + 1]);
            }
            foreach (var dataGridView in dataGridView_ExtraList)
            {
                dataGridView_list.Add(dataGridView);
            }
        }

        private void sqlToListData(String productCode)
        {
            dataAccessProccess();
            try
            {

                sqlCommand = new SqlCommand("Select distinct (b.COMPONENT), b.BOMNUMBER,m.STEXT from IASBOMITEM b inner join IASMATX m on b.COMPONENT=m.MATERIAL inner join IASPRDORDER p on p.MATERIAL=b.MATERIAL where  p.PRDORDER= '" + productCode + "' and p.PLANT='20'", sqlConnection);
                sqlDataReader = sqlCommand.ExecuteReader();
                top_code_list = new List<string>();
                material_code_list = new List<string>();
                description_list = new List<string>();
                rcode_list = new List<string>();
                while (sqlDataReader.Read())
                {
                    top_code_list.Add(sqlDataReader.GetString(1));
                    material_code_list.Add(sqlDataReader.GetString(0));
                    description_list.Add(sqlDataReader.GetString(2));
                }
                sqlDataReader.Close();
                dataGridViewRowCount = top_code_list.Count;
                dataNumber = float.Parse(top_code_list.Count.ToString());
                if (dataNumber != 0)
                {
                    flagProduct = true;
                    lblPrdOrder.Text = "İş Emri : " + txtProductCode.Text;
                    lblYMaterial.Text = "Y. Mamul : " + top_code_list[0];
                    sqlCommand = new SqlCommand("select STEXT from IASMATX where MATERIAL='" + top_code_list[0] + "' AND TEXTTYPE='M'", sqlConnection);
                    sqlDataReader = sqlCommand.ExecuteReader();
                    if (sqlDataReader.Read())
                    {
                        productName = sqlDataReader["STEXT"].ToString();
                    }
                    sqlDataReader.Close();
                }
                else
                {
                    flagProduct = false;
                }

                excelPart = dataNumber / 17;
                if (excelPart - (int)excelPart == 0)
                {
                    excelPart = (int)excelPart;
                }
                else
                {
                    excelPart = (int)excelPart + 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veritabanı bağlantı sorunu");
            }
        }

        private void openExcelFile(String filePath)
        {
            excel = new ExcelApp();
            sheet = excel.Workbooks.Open(filePath);
            worksheet = excel.ActiveSheet as WorkSheet;
            userRange = worksheet.UsedRange;
            worksheet.Columns.AutoFit();//Columns auto fitting
            indexRow = 7;
            constNumber = 17;

        }
        private void titleExcel()
        {
           
            productNo =top_code_list[0];
            productDate = DateTime.Now.ToString("dd.MM.yyyy");
        }

        //parameter[0] : reference excel path
        private bool dataWriteToExcel(params dynamic[] parameter)
        {
            try
            {
                openExcelFile(parameter[0]);

                for (int i = 0; i < top_code_list.Count; i++)
                {
                    if (constNumber > i)
                    {
                        worksheet.Cells[indexRow, 1] = top_code_list[i];
                        worksheet.Cells[indexRow, 2] = material_code_list[i];
                        worksheet.Cells[indexRow, 3] = description_list[i];
                        worksheet.Columns.AutoFit();
                        indexRow++;
                    }
                    if (constNumber == i + 1)
                    {
                        constNumber = constNumber + 17;
                        indexRow = indexRow + 8;
                    }
                    worksheet.Cells[3, 1] = "Ürün Adı : " + productName;
                    worksheet.Cells[3, 12] = productNo;
                    worksheet.Cells[4, 7] = productDate;
                    worksheet.Cells[5, 15] =txtProductCode.Text;
                    worksheet.Columns.AutoFit();
                }
                constRowNumber = material_code_list.Count;

                // New Page

                /*  worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheet.Worksheets.Add();
                  sheet = excel.Workbooks.Add(Type.Missing);
                  worksheet.Name="New Sheet"*/

                // Copy Sheet
                /* WorkSheet sheet1 = sheet.Sheets[1];
                 worksheet.Copy(Type.Missing, sheet1);

                 var copySheetIndex = sheet1.Index + 1;

                 WorkSheet copySheet = sheet.Sheets.get_Item(copySheetIndex);
                 copySheet.Name = "Copy Sheet";*/


                /*  Excel.Range firstCell = (worksheet.Cells[6, 1] as Excel.Range);
                  string firstCellValue = userRange.Value.ToString();
                  if (firstCellValue == "System.Object[,]")
                  {
                  }/*

                  /*int countRecords = userRange.Rows.Count;
                  int add = countRecords + 1;
                  worksheet.Cells[add, 1] = "Total Rows2231" + countRecords;*/


                int n = 0;

                for (int i = 0; i < top_code_list.Count; i++)
                {
                    n = dataGridViewMaterial.Rows.Add();
                    dataGridViewMaterial.Rows[n].Cells[0].Value = (i + 1);
                    dataGridViewMaterial.Rows[n].Cells[1].Value = top_code_list[i];
                    dataGridViewMaterial.Rows[n].Cells[2].Value = material_code_list[i];
                    dataGridViewMaterial.Rows[n].Cells[3].Value = description_list[i];
                }
                worksheet.Columns.AutoFit();

                int m = 0;
                for (int i = 0; i < dataGridViewMaterial.RowCount; i++)
                {
                    m = dataGridView1.Rows.Add();
                    dataGridView1.Rows[m].Cells[0].Value = (i + 1);
                    dataGridView1.Rows[m].Cells[1].Value = "";
                    dataGridView1.Rows[m].Cells[2].Value = "";

                    m = dataGridView2.Rows.Add();
                    dataGridView2.Rows[m].Cells[0].Value = (i + 1);
                    dataGridView2.Rows[m].Cells[1].Value = "";
                    dataGridView2.Rows[m].Cells[2].Value = "";

                    m = dataGridView3.Rows.Add();
                    dataGridView3.Rows[m].Cells[0].Value = (i + 1);
                    dataGridView3.Rows[m].Cells[1].Value = "";
                    dataGridView3.Rows[m].Cells[2].Value = "";

                    m = dataGridView4.Rows.Add();
                    dataGridView4.Rows[m].Cells[0].Value = (i + 1);
                    dataGridView4.Rows[m].Cells[1].Value = "";
                    dataGridView4.Rows[m].Cells[2].Value = "";

                    for (int j = 0; j < dataGridView_ExtraList.Count; j++)
                    {
                        m = dataGridView_ExtraList[j].Rows.Add();
                        dataGridView_ExtraList[j].Rows[m].Cells[0].Value = (i + 1);
                        dataGridView_ExtraList[j].Rows[m].Cells[1].Value = "";
                        dataGridView_ExtraList[j].Rows[m].Cells[2].Value = "";
                    }



                }
                sheet.Save();

                excelFail = "";
                return true;
            }
            catch (Exception)
            {
                excelFail = "Excel Dosyası Bulunamadı !";
                return false;
            }
        }



        int x = 0;
        //parameter[0]=eps code
        //parameter[1]=datagridview
        //parameter[2]=excel lotNum coloumn index
        //parameter[3]=excel quantity column index

        //threadCall(threadAddComponent, addComponent, epsCode, dataGridView_list[indexView + 1], indexExcelCell, indexExcelCell + 3);
        private void addComponent(params dynamic[] parameter)
        {
            try
            {

         
            bool flagMatchReturn = false;

            if (checkBox1.Checked == true)
            {
                // MessageBox.Show(bomNumberNow.ToString() + " / " + epsCode.ToString());

                int y = 0;


                top_code_list.Add(top_code_list[0]);
                material_code_list.Add(epsCode);
                description_list.Add("-");
                y = dataGridViewMaterial.Rows.Add();
                dataGridViewMaterial.Rows[y].Cells[0].Value = dataGridViewMaterial.RowCount;
                dataGridViewMaterial.Rows[y].Cells[2].Value = epsCode;
                dataGridViewMaterial.Rows[y].Cells[1].Value = top_code_list[top_code_list.Count - 1];
                dataGridViewMaterial.Rows[y].Cells[3].Value = "-";
                rowX = findNextEpsCodeIndexRow((dataGridViewMaterial.RowCount) - 1);
                worksheet.Cells[rowX, 2] = epsCode;
                worksheet.Cells[rowX, 1] = top_code_list[top_code_list.Count - 1];
                worksheet.Cells[rowX, 3] ="-";
                flagMatchReturn = true;
                flagMatch = false;
                int x = 0;
                for (int j = 0; j < dataGridView_list.Count; j++)
                {
                    x = dataGridView_list[j].Rows.Add();
                    dataGridView_list[j].Rows[x].Cells[0].Value = dataGridViewMaterial.RowCount;
                    dataGridView_list[j].Rows[x].Cells[1].Value = "";
                    dataGridView_list[j].Rows[x].Cells[2].Value = "";

                }
                dataGridView_list[0].Rows[x].Cells[1].Value = lotNo + "/" + rCode;
                dataGridView_list[0].Rows[x].Cells[2].Value = quantity;
                worksheet.Cells[rowX, 4] = lotNo + "/" + rCode;
                worksheet.Cells[rowX, 7] = quantity;
                rcode_list.Add(0.ToString() + "Y" + (dataGridViewMaterial.RowCount - 1).ToString() + rCode);

            }

            indexRow = 7;
            constNumber = 17;
            bool flag = false;
            if (!flagMatch)
            {
                for (int i = 0; i < dataGridViewMaterial.Rows.Count; i++)
                {
                    if (constNumber > i)
                    {
                        userRange = (worksheet.Cells[indexRow, 2] as Excel.Range);
                        if (!flag && parameter[0] == userRange.Value)
                        {
                            bool flagN = false, flagN2 = false;

                            for (int j = 0; j < rcode_list.Count; j++)
                            {
                                if (calculateComp(rcode_list[j]) == rCode)
                                {
                                    flagN = true;
                                }
                            }
                            if (parameter[1].Rows[i].Cells[1].Value.ToString() == "")
                            {
                                flagN2 = true;
                            }
                            if (!flagN && flagN2 && (addComponent2Flag || constRowNumber >= indexView3 + 1))
                            {
                                x++;
                                worksheet.Cells[indexRow, parameter[2]] = lotNo + "/" + rCode;
                                DataGridViewRow row = parameter[1].Rows[i];
                                userRange = (worksheet.Cells[indexRow, parameter[2]] as Excel.Range);
                                parameter[1].Rows[i].Cells[1].Value = userRange.Value.ToString();
                                worksheet.Cells[indexRow, parameter[3]] = quantity;
                                userRange = (worksheet.Cells[indexRow, parameter[3]] as Excel.Range);
                                parameter[1].Rows[i].Cells[2].Value = userRange.Value.ToString();
                                rcode_list.Add(parameter[4] + "Y" + i.ToString() + rCode);
                                flag = true;
                                lblFirstRecord.Text = "Listeye Eklendi !";
                                lblFirstRecord.ForeColor = Color.Green;
                                worksheet.Columns.AutoFit();
                                lblDoneMessage.ForeColor = Color.Green;
                                lblDoneMessage.Text = "Malzeme Eklendi";
                                lblDoneMess(true, "Malzeme Eklendi");
                                if (serialPortAlarm.IsOpen)
                                {
                                    byteArray[0] = 83;
                                    serialPortAlarm.Write(byteArray,0,1);
                                    serialBufferClear();
                                }
                            }
                            else
                            {
                                flag = true;
                                int n = 0;
                                rowX = findNextEpsCodeIndexRow(dataGridViewMaterial.RowCount);
                                bool flagM = false;

                                for (int j = 0; j < rcode_list.Count; j++)
                                {
                                    if (calculateComp(rcode_list[j]) == rCode)
                                    {
                                        flagM = true;
                                    }
                                }

                                if (!flagM)
                                {
                                    if (addComponent2Flag)
                                    {
                                        n = dataGridViewMaterial.Rows.Add();
                                        dataGridViewMaterial.Rows[n].Cells[0].Value = dataGridViewMaterial.RowCount;
                                        userRange = (worksheet.Cells[indexRow, 2] as Excel.Range);
                                        dataGridViewMaterial.Rows[n].Cells[2].Value = userRange.Value.ToString();
                                        userRange = (worksheet.Cells[indexRow, 1] as Excel.Range);
                                        dataGridViewMaterial.Rows[n].Cells[1].Value = userRange.Value.ToString();
                                        userRange = (worksheet.Cells[indexRow, 3] as Excel.Range);
                                        dataGridViewMaterial.Rows[n].Cells[3].Value = userRange.Value.ToString();
                                        worksheet.Cells[rowX, 2] = epsCode;
                                        findDataComponentToExcel(epsCode, rowX);
                                        worksheet.Columns.AutoFit();
                                    }
                                    if (addComponent2Flag)
                                    {
                                        n = dataGridView_list[0].Rows.Add();
                                        dataGridView_list[0].Rows[n].Cells[0].Value = dataGridViewMaterial.RowCount;
                                        dataGridView_list[0].Rows[n].Cells[1].Value = lotNo + "/" + rCode;
                                        worksheet.Cells[rowX, 4] = lotNo + "/" + rCode;
                                        dataGridView_list[0].Rows[n].Cells[2].Value = quantity;
                                        worksheet.Cells[rowX, 7] = quantity;
                                        rcode_list.Add(0.ToString() + "Y" + n.ToString() + rCode);
                                        for (int j = 1; j < dataGridView_list.Count; j++)
                                        {
                                            n = dataGridView_list[j].Rows.Add();
                                            dataGridView_list[j].Rows[n].Cells[0].Value = dataGridViewMaterial.RowCount;
                                            dataGridView_list[j].Rows[n].Cells[1].Value = "";
                                            dataGridView_list[j].Rows[n].Cells[2].Value = "";
                                            }
                                        worksheet.Columns.AutoFit();
                                    }
                                    else
                                    {

                                        rowX = findNextEpsCodeIndexRow(indexView3);
                                        worksheet.Cells[rowX, parameter[2]] = lotNo + "/" + rCode;
                                        worksheet.Cells[rowX, parameter[3]] = quantity.ToString();

                                        for (int t = 0; t < rowX + 5; t++)
                                        {
                                            worksheet.Columns.AutoFit();
                                        }
                                        for (int z = 0; z < 15; z++)
                                        {
                                            worksheet.Columns[(4*z)+8].ColumnWidth = 102;

                                        }
                                      
                                        dataGridView_list[indexView2 + 1].Rows[indexView3].Cells[0].Value = (indexView3 + 1).ToString();
                                        dataGridView_list[indexView2 + 1].Rows[indexView3].Cells[1].Value = lotNo + "/" + rCode;
                                        dataGridView_list[indexView2 + 1].Rows[indexView3].Cells[2].Value = quantity;
                                        rcode_list.Add((indexView2 + 1).ToString() + "Y" + indexView3.ToString() + rCode);
                                        worksheet.Columns.AutoFit();
                                        }
                                    lblFirstRecord.Text = "Listeye Eklendi !";
                                    lblFirstRecord.ForeColor = Color.Green;
                                    lblDoneMess(true, "Malzeme Eklendi");
                                    if (serialPortAlarm.IsOpen)
                                    {
                                            byteArray[0] = 83;
                                            serialPortAlarm.Write(byteArray, 0, 1);
                                            // serialPortAlarm.Write("S");
                                            serialBufferClear();
                                        }
                                }
                                else
                                {
                                    lblDoneMessage.Visible = true;
                                    lblDoneMessage.ForeColor = Color.Red;
                                    lblDoneMessage.Text = "Malzeme Eklenmedi";

                                    lblFirstRecord.Text = "Aynı Malzeme !";
                                    lblFirstRecord.ForeColor = Color.Red;
                                    lblDoneMess(false, "Malzeme Eklenmedi");
                                    if (serialPortAlarm.IsOpen)
                                    {
                                            byteArray[0] = 69;
                                            serialPortAlarm.Write(byteArray, 0, 1);
                                         //   serialPortAlarm.Write("E");
                                        serialBufferClear();
                                    }
                                }
                                setDefaultToParameter();
                            }

                        }
                        indexRow++;
                    }
                    if (constNumber == i + 1)
                    {
                        constNumber = constNumber + 17;
                        indexRow = indexRow + 8;
                    }
                }
            }


            if (flag == false && !flagMatchReturn)
            {
                    if (serialPortAlarm.IsOpen)
                    {
                        byteArray[0] = 69;
                        serialPortAlarm.Write(byteArray, 0, 1);
                       // serialPortAlarm.Write("E");
                        serialBufferClear();
                    }
                lblFirstRecord.ForeColor = Color.Red;
                lblFirstRecord.Text = "Ürün Bulunamadı !";
                lblDoneMess(false, "Malzeme Eklenmedi !");
            }
            if (flagMatchReturn)
            {
                lblFirstRecord.Text = "Listeye Eklendi !";
                lblFirstRecord.ForeColor = Color.Green;
                lblDoneMess(true, "Malzeme Eklendi");
                    if (serialPortAlarm.IsOpen)
                    {
                        byteArray[0] = 83;
                        serialPortAlarm.Write(byteArray, 0, 1);
                     //   serialPortAlarm.Write("S");
                        serialBufferClear();
                    }
                txtFirstRecordBarcode.Focus();
                checkBox1.Checked = false;
            }

            sheet.Save();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel veri yazma hatası" + ex);
            }

        }
        private bool findComponent()
        {
            bool flag = false;
            for (int i = 0; i < rcode_list.Count; i++)
            {
                if (!emptyControlData(rCodeI(rcode_list[i]), rCodeJ(rcode_list[i])) && calculateComp(rcode_list[i]) != rCode && rCodeJ(rcode_list[i]) == dataGridView_list.Count && flag)
                {
                    lblChangeBarcodeO.Text = "Değişiklik yapılamaz !";
                    lblChangeBarcodeO.ForeColor = Color.Red;
                    txtChangeBarcodeO.Text = "";
                    flag = false;
                    break;
                }
                else if (emptyControlData(rCodeI(rcode_list[i]), rCodeJ(rcode_list[i])) && calculateComp(rcode_list[i]) == rCode && !flag)
                {

                    int rI = rCodeI(rcode_list[i]);
                    int rJ = rCodeJ(rcode_list[i]);
                    setColorFindCells(dataGridView_list[rI], rJ, rI + 1, Color.Yellow, Color.Black);
                    indexView2 = rI;
                    indexView3 = rJ;
                    lockToObject(txtChangeBarcodeO, btnFirstRecord);
                    lockToObject(txtChangeBarcodeO, btnRecover);
                    openToObject(txtChangeBarcodeN, btnChangeConfirmN);
                    indexView = rI;
                    indexComponentRow = rJ;
                    flag = true;
                    break;

                }

            }

            return flag;
        }
        private string calculateComp(string rCodeData)
        {

            string firstCharacter = "";
            char[] characters = rCodeData.ToCharArray();
            for (int i = 0; i < rCodeData.Length; i++)
            {

                if (characters[i].ToString() == "R" && Char.IsLetter(characters[i]))
                {

                    firstCharacter = rCodeData.Substring(i, rCodeData.Length - i);
                }

            }
            return firstCharacter;
        }




        private int rCodeI(string rCodeData)
        {
            try
            {
                int dataIndex = 0;
                char[] characters = rCodeData.ToCharArray();
                for (int i = 0; i < rCodeData.Length; i++)
                {

                    if (characters[i].ToString() == "Y" && Char.IsLetter(characters[i]))
                    {

                        dataIndex = i;
                    }

                }
                dataIndex = Int16.Parse(rCodeData.Substring(0, dataIndex));
                return dataIndex;
            }
            catch (Exception)
            {

                throw;
            }

        }

        private int rCodeJ(string rCodeData)
        {
            try
            {
                int dataRow = 0, index = 0;
                char[] characters = rCodeData.ToCharArray();
                for (int i = 0; i < rCodeData.Length; i++)
                {
                    if (characters[i].ToString() == "Y" && Char.IsLetter(characters[i]))
                    {
                        index = i + 1;
                    }

                    if (characters[i].ToString() == "R" && Char.IsLetter(characters[i]))
                    {

                        dataRow = i - index;
                        break;
                    }


                }
                dataRow = Int16.Parse(rCodeData.Substring(index, dataRow));
                return dataRow;
            }
            catch (Exception)
            {

                throw;
            }

        }
        private bool emptyControlData(int i, int j)
        {
            bool flag = true;
            if (dataGridView_list[i + 1].Rows[j].Cells[1].Value.ToString() == "")
            {
                flag = true;
            }
            else
            {
                flag = false;
            }
            return flag;
        }
        private void findDataComponentToExcel(string epscode, int rowX)
        {
            for (int i = 0; i < material_code_list.Count; i++)
            {
                if (material_code_list[i] == epscode)
                {
                    worksheet.Cells[rowX, 1] = top_code_list[i];
                    worksheet.Cells[rowX, 3] = description_list[i];
                }
            }
        }
        private void setColorFindCells(DataGridView dataGridView, int row, int nextDataGridViewListİndex, Color color, Color colorText)
        {
            try
            {
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    dataGridView.Rows[row].Cells[i].Style.BackColor = color;
                    dataGridView_list[nextDataGridViewListİndex].Rows[row].Cells[i].Style.BackColor = color;
                    dataGridView.Rows[row].Cells[i].Style.ForeColor = colorText;
                    dataGridView_list[nextDataGridViewListİndex].Rows[row].Cells[i].Style.ForeColor = colorText;
                }
            }
            catch (Exception)
            {

            }
          
        }

        private bool epsControl(DataGridView dataGridView, int indexRow, int indexCell)
        {
            if (dataGridView.Rows[indexRow].Cells[indexCell].Value.ToString() == epsCode)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void openToObject(TextBox textBox, Button button)
        {
            textBox.Text = "";
            textBox.Enabled = true;
            button.Enabled = true;
            textBox.Focus();
        }
        private void lockToObject(TextBox textBox, Button button)
        {
            textBox.Text = "";
            textBox.Enabled = false;
            button.Enabled = false;
            textBox.Focus();
        }

        private bool componentControl(DataGridView dataGridView, int rowNumber)
        {

            DataGridViewRow row = dataGridView.Rows[rowNumber];

            if (dataGridView.Rows[rowNumber].Cells[1].Value.ToString() != "")
            {

                return true;
            }
            else
            {
                return false;
            }
        }
        private bool componentControl(DataGridView dataGridView, int rowNumber, string lotNumber, string epsCodeL)
        {

            DataGridViewRow row = dataGridView.Rows[rowNumber];

            // MessageBox.Show("My data : " + dataGridView.Rows[rowNumber].Cells[1].Value.ToString() +" lot "+ lotNumber + "my data 2 : "+ dataGridViewMaterial.Rows[rowNumber].Cells[2].Value.ToString() +" "+epsCodeL);
            if (/*dataGridView.Rows[rowNumber].Cells[1].Value.ToString() != lotNumber &&*/ dataGridViewMaterial.Rows[rowNumber].Cells[2].Value.ToString() == epsCodeL)
            {
                // MessageBox.Show("true");
                return true;

            }
            else
            {
                // MessageBox.Show("false");
                return false;
            }
        }


        int timerNum = 0;

        private void txtChangeBarcodeO_TextChanged(object sender, EventArgs e)
        {
            if (searchAlgorithm)
            {
                searchAlgorithmLot = true;
            }
            else
            {
                searchAlgorithmEps = true;
            }

            timer4.Enabled = true;
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            if (timerNum == 2 && ((1 <= txtChangeBarcodeO.Text.Length && txtChangeBarcodeO.Text.Length <= 90) || (1 <= txtChangeBarcodeN.Text.Length && txtChangeBarcodeN.Text.Length <= 90)))
            {
                if (searchAlgorithm)
                {
                    if (searchAlgorithmLot)
                    {
                        ConfirmO();
                    }
                    else
                    {
                        ConfirmN();
                    }
                }
                else
                {
                    if (searchAlgorithmEps)
                    {
                        compareEpsO();

                    }
                    else
                    {
                        compareEpsN();
                    }
                }
                timerNum = 0;
                timer4.Enabled = false;
            }
            else if (timerNum == 3)
            {
                timerNum = 0;
                timer4.Enabled = false;
            }

            timerNum++;

        }

        private void btnMatch_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Muadil malzeme eklemek istediğinize emin misiniz ?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                checkBox1.Checked = true;
                flagMatch = true;
                txtFirstRecordBarcode.Focus();
            }

        }
        private void equivalentComponent() 
        {
          
        }

        private void txtChangeBarcodeN_TextChanged(object sender, EventArgs e)
        {
            if (searchAlgorithm)
            {
                searchAlgorithmLot = false;
            }
            else
            {
                searchAlgorithmEps = false;
            }

            timer4.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                flagMatch = true;
            }
            else
            {
                flagMatch = false;
            }
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
             
        }
        int l = 0;
        private void button2_Click(object sender, EventArgs e)
        {
            //string tt = "R847614,MEPS648311,PSC11-210404295,U26.04.2021,Q4000.0,X25.04.2022";
            ///MessageBox.Show(tt.Substring(0, 1));
            /*  string tt2 = "R1066644,MEPSB0600019,P22-1TMQK306791,U29.07.2020,Q3000.0,X27.07.2022";
              string tt3 = "R905271,MEPS400532,P21-1TWEpA24143G1A,U25.06.2021,Q3000.0,X23.06.2023";
              barcodeSplit(tt);
              MessageBox.Show(epsCode+" "+lotNo+" "+quantity);*/
            /*   if (serialPortAlarm.IsOpen)
               {
                   byteArray[0] = 70;
                   serialPortAlarm.Write(byteArray, 0, 1);
                 //  serialPortAlarm.Write("S");
                   serialBufferClear();
               }
               else
               {
                   MessageBox.Show("close");
               }*/

            MessageBox.Show(findNextEpsCodeIndexRow(l).ToString());
            l++;
         
           
        }

        private void txtProductCode_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataGridViewMaterial.FirstDisplayedScrollingRowIndex = dataGridViewMaterial.FirstDisplayedScrollingRowIndex + 1;
                for (int i = 0; i < dataGridView_list.Count; i++)
                {
                    this.dataGridView_list[i].FirstDisplayedScrollingRowIndex = dataGridView_list[i].FirstDisplayedScrollingRowIndex + 1;
                }
                button1.Focus();
            }
            catch (Exception)
            {


            }

        }

        private void dataGridViewMaterial_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            try
            {
                this.dataGridViewMaterial.FirstDisplayedScrollingRowIndex = dataGridViewMaterial.FirstDisplayedScrollingRowIndex - 1;
                for (int i = 0; i < dataGridView_list.Count; i++)
                {
                    this.dataGridView_list[i].FirstDisplayedScrollingRowIndex = dataGridView_list[i].FirstDisplayedScrollingRowIndex - 1;
                }
                button1.Focus();
            }
            catch (Exception)
            {


            }

        }

        private void btnAlt_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void getExcelFilesURL(String fileName, String copyFileName)
        {
            referenceFileURL = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\CANIAS_Izlenebilirlik\ExcelFile\" + fileName + (excelPart + 2).ToString() + ".xlsx";
            date = DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss");
            copyFileName = copyFileName + "_" + date + "_" + txtProductCode.Text;
            string destinationURL = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\CANIAS_Izlenebilirlik\Devam_Eden_Listeler\" + copyFileName +".xlsx";
            File.Copy(referenceFileURL, destinationURL);
            excelFile = destinationURL;
            finishedFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\CANIAS_Izlenebilirlik\Biten_Listeler\" + copyFileName + ".xlsx";
            t20ExcelName = copyFileName + (excelPart + 2).ToString() + ".xlsx";
        }

        private void moveToFinishedFile(String onGoingFile, String finishedFile)
        {
            File.Move(onGoingFile, finishedFile);
        }


        private void txtProductCode_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            txtProductCode.Text = materialCodeCopyData;


        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            progressBarProductCode.Increment(10);
            if (progressBarProductCode.Value >= progressBarProductCode.Maximum)
            {
                if (flagProduct)
                {
                    lblProductMessage.ForeColor = Color.Green;
                    lblProductMessage.Text = "Ürün Yüklendi !";
                }
                else
                {
                    lblProductMessage.Text = "Ürün Bulunamadı !";
                    lblProductMessage.ForeColor = Color.Red;
                    disableButton(2, true, btnProductConfirm, btnRecover);
                    btnProductConfirm.BackColor = Color.FromArgb(0, 0, 64);
                    btnProductConfirm.ForeColor = Color.White;
                    progressBarProductCode.Value = 0;
                    isOpened = false;
                    timer3.Start();
                    changeFlag = false;
                    changeMessage = "Ürün Onaylanmadı";

                }
                timer1.Stop();
            }
            if (excelFail == "Excel Dosyası Bulunamadı !")
            {
                lblProductMessage.Text = excelFail;
                lblProductMessage.ForeColor = Color.Red;

            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (isOpened)
            {
                if (MessageBox.Show("Çıkış yapmak istediğinize emin misiniz ?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    sheet.Save();
                    sheet.Close(true, Type.Missing, Type.Missing);
                    excel.Quit();
                    isOpened = false;
                    moveToFinishedFile(excelFile, finishedFile);
                    Application.Exit();
                }
                else
                {
                    sheet.Save();
                    isOpened = true;
                }
             
            }
            else
            {
                Application.Exit();
            }
           
        }

        private void btnExit_MouseHover(object sender, EventArgs e)
        {
            btnExit.ForeColor = Color.Red;
        }

        private int findNextEpsCodeIndexRow(int index)
        {
            indexComponentRow = 0;
            setDefaultToParameter();
            do
            {
                indexComponentRow = indexComponentRow + 17;
                coefficientNumber++;
            } while (index >= indexComponentRow);
            constNumberPage = constNumberPage + (25 * (coefficientNumber - 1));
            headIndex = headIndex + (17 * (coefficientNumber - 1));
            constNumberPage = constNumberPage + (index - headIndex);
            return constNumberPage;
            
        }
        private void txtFirstRecordBarcode_TextChanged(object sender, EventArgs e)
        {
            timer2.Start();
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            barcodeSplitCall();
        }

        private void barcodeSplitCall()
        {
            timeBarcode++;
            bool flag = false;
            if (barcodeSplit(txtFirstRecordBarcode.Text) && timeBarcode == 1)
            {
                addComponent2Flag = true;
                addComponent( epsCode, dataGridView1, 4, 7, 0);
                txtFirstRecordBarcode.Text = "";
                txtFirstRecordBarcode.Focus();
                flag = true;
            }
            else if (timeBarcode == 1)
            {
                lblFirstRecord.Text = "Eksik veya fazla barcode okundu !";
                lblFirstRecord.ForeColor = Color.Red;
                txtFirstRecordBarcode.Text = "";
                lblDoneMess(false, "Malzeme Eklenmedi !");
                if (serialPortAlarm.IsOpen)
                {
                    byteArray[0] = 69;
                    serialPortAlarm.Write(byteArray, 0, 1);
                    //serialPortAlarm.Write("E");
                    serialBufferClear();
                }

            }
            if (timeBarcode == 3)
            {
                lblDoneMessRemove();
                timer2.Stop();
                timeBarcode = 0;
                lblFirstRecord.Text = "Lütfen Barkodu Okutunuz !";
                lblFirstRecord.ForeColor = Color.Black; ;
            }
        }

        private bool epsCodeMatchFirstDGV()
        {
            int index = 0;
            bool flag = false;
            for (int i = 0; i < dataGridViewMaterial.RowCount; i++)
            {
                if (dataGridViewMaterial.Rows[i].Cells[2].Value.ToString() == epsCode && flag == false)
                {
                    //MessageBox.Show("l : " + lotNo + " eps " + epsCode + " qu " + quantity);
                    index = dataGridViewMaterial.Rows.Add();
                    dataGridViewMaterial.Rows[index].Cells[0].Value = dataGridViewMaterial.RowCount;
                    dataGridViewMaterial.Rows[index].Cells[1].Value = dataGridViewMaterial.Rows[i].Cells[1].Value;
                    dataGridViewMaterial.Rows[index].Cells[2].Value = epsCode;
                    dataGridViewMaterial.Rows[index].Cells[3].Value = dataGridViewMaterial.Rows[i].Cells[3].Value;
                    flag = true;
                }
            }
            return flag;
        }

        private void materialCompare()
        {
            sqlCommand = new SqlCommand("Select distinct (b.COMPONENT), b.BOMNUMBER,m.STEXT from IASBOMITEM b inner join IASMATX m on b.COMPONENT=m.MATERIAL inner join IASPRDORDER p on p.MATERIAL=b.MATERIAL where  p.PRDORDER= '" + txtProductCode.Text + "' and p.PLANT='20'", sqlConnection);
            SqlDataReader sqlDataReaderMatCom = sqlCommand.ExecuteReader();
            material_code_list = new List<string>();
            int data = 0;
            dataGridView_list = new List<DataGridView>();
            dataGridView_list.Add(dataGridView1);
            dataGridView_list.Add(dataGridView2);
            lblPrdOrder.Text = "İş Emri : " + txtProductCode.Text;
            while (sqlDataReaderMatCom.Read())
            {
                material_code_list.Add(sqlDataReaderMatCom.GetString(0));
            }
            sqlDataReaderMatCom.Close();
            int n = 0;
            for (int i = 0; i < material_code_list.Count; i++)
            {
                n = dataGridView_list[0].Rows.Add();
                dataGridView_list[0].Rows[n].Cells[0].Value = (i + 1);
                dataGridView_list[0].Rows[n].Cells[1].Value = material_code_list[i];
                data++;
            }
            if (data > 0)
            {
                int m = 0;
                for (int j = 0; j < material_code_list.Count; j++)
                {
                    m = dataGridView_list[1].Rows.Add();
                    dataGridView_list[1].Rows[m].Cells[0].Value = (j + 1);
                    dataGridView_list[1].Rows[m].Cells[1].Value = "";
                    dataGridView_list[1].Rows[m].Cells[2].Value = "";
                }
                searchAlgorithm = false;
                disableButton(2, true, btnDone, btnChange);
                disableButton(3, false, btnProductConfirm, btnRecover, btnFirstRecordConfirm);
            }
            else
            {
                timer3.Start();
                changeFlag = false;
                changeMessage = "Ürüm Kodu Bulunamadı !";
            }

        }

        private void compareEpsO()
        {
            if (barcodeSplit(txtChangeBarcodeO.Text))
            {
                btnChangeConfirmO.Text = epsCode;
                findDataGridViewEps();
                txtChangeBarcodeO.Text = "";
               
            }
            else
            {
                lblChangeBarcodeO.Text = "Eksik Barcode Kodu !";
                lblChangeBarcodeO.ForeColor = Color.Red;
                timer3.Start();
                changeFlag = false;
                changeMessage = "Eksik Barcode Kodu !";
            }
        }
        private void compareEpsN()
        {
            if (barcodeSplit(txtChangeBarcodeN.Text))
            {
                txtChangeBarcodeN.Text = "";
                btnChangeConfirmN.Text = epsCode;
                findDataGridEmptyIndex();
                lblChangeBarcodeN.Text = "Yeni Bileşenin Barkodunu Okutunuz !";
                lblChangeBarcodeN.ForeColor = Color.White;
                
            }
            else
            {
                timer3.Start();
                changeFlag = false;
                changeMessage = "Eksik Barcode Kodu !";
            }
           
        }

        private int findDataRowIndex()
        {
            int data = -1;
            for (int i = 0; i < material_code_list.Count; i++)
            {
                if (material_code_list[i] == epsCode)
                {
                    data = i;
                    tempEpsCodeIndex = i;
                    epsCode = "";
                }
            }
            return data;
        }
        private void findDataGridViewEps()
        {
            bool flagLoop = true;
            int indexEpsRow = findDataRowIndex();
            if (indexEpsRow >= 0)
            {
                tempEpsCodeIndex = indexEpsRow;
                for (int i = 1; i < dataGridView_list.Count; i++)
                {
                    if (dataGridView_list[1].Rows[indexEpsRow].Cells[1].Value.ToString() == "" && flagLoop)
                    {

                        setColorFindCells(dataGridView_list[i - 1], indexEpsRow, i, Color.Yellow, Color.Black);
                        firstEps = indexEpsRow;
                        lblChangeBarcodeO.Text = "Biten Bileşenin Barkodunu Okutunuz !";
                        lblChangeBarcodeO.ForeColor = Color.White;
                        txtChangeBarcodeN.Enabled = true;
                        btnChangeConfirmO.BackColor = Color.ForestGreen;
                        btnChangeConfirmN.BackColor = Color.ForestGreen;
                        timer3.Start();
                        changeFlag = true;
                        changeMessage = "Malzeme Onaylandı";
                        flagLoop = false;
                        txtChangeBarcodeN.Focus();
                    }
                }
            }
            else
            {
                txtChangeBarcodeO.Focus();
                lblChangeBarcodeO.Text = "EPS Kodu Bulunamadı !";
                lblChangeBarcodeO.ForeColor = Color.Red;
                btnChangeConfirmO.BackColor = Color.Red;
                timer3.Start();
                changeFlag = false;
                changeMessage = "EPS Kodu Bulunamadı !";
            }
        }

        private void findDataGridEmptyIndex()
        {

            if (dataGridView_list[0].Rows[firstEps].Cells[1].Value.ToString() == epsCode)
            {
                setColorFindCells(dataGridView_list[0], tempEpsCodeIndex, 1, Color.ForestGreen, Color.White);
                timer3.Start();
                changeFlag = true;
                changeMessage = "Malzeme Onaylandı";
                btnChangeConfirmO.BackColor = Color.Green;
                btnChangeConfirmN.BackColor = Color.Green;
                txtChangeBarcodeO.Focus();

            }
            else
            {
                setColorFindCells(dataGridView_list[0], firstEps, 1, Color.Red, Color.White);
                btnChangeConfirmO.BackColor = Color.Red;
                btnChangeConfirmN.BackColor = Color.Red;
                timer3.Start();
                changeFlag = false;
                changeMessage = "Eşleşme Hatası !";
                txtChangeBarcodeN.Focus();
            }
               

            

        }
        private void getCardList()
        {
            if (!isCardListOpened)
            {
                isCardListOpened = true;
                cardList = new CardList();
                cardList.ShowDialog();
                isCardListOpened = false;
            }
        }

        private void btnRecover_Click(object sender, EventArgs e)
        {
            recoverList();
        }

        private void recoverList()
        {
            try
            {
                disableButton(1, false, btnProductConfirm);
                btnProductConfirm.BackColor = Color.FromArgb(190, 151, 144);
                lblDoneMessRemove();
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\CANIAS_Izlenebilirlik\Biten_Listeler\";
                fileDialog.DefaultExt = ".xlsx";
                fileDialog.Filter = "Excel Dosyası |*.xlsx"; // Optional file extensions
                if (!isOpened)
                {
                    if (fileDialog.ShowDialog() != DialogResult.OK)
                    {
                        disableButton(1, true, btnProductConfirm);
                        btnProductConfirm.BackColor = Color.FromArgb(0, 0, 64);
                    }

                    excelFile = fileDialog.FileName;
                    finishedFile = fileDialog.InitialDirectory + fileDialog.SafeFileName;
                    if (excelFile != "")
                    {
                        searchAlgorithm = true;
                        flagRecover = true;
                        openExcelFile(excelFile);
                        dataGridViewListCall(4, dataGridView1, dataGridView2, dataGridView3, dataGridView4);
                        setDefaultToParameter();
                        rcode_list = new List<string>();
                        int x = 0, n = 0, m = 0;
                        bool flag = true;
                        int i = 0;
                        string[] dataRcodeArray;
                        do
                        {
                            if (constNumber > i)
                            {
                                userRange = (worksheet.Cells[indexRow, 1] as Excel.Range);

                                if (userRange.Value == null)
                                {
                                    flag = false;
                                }
                                else
                                {
                                    userPrCode = (worksheet.Cells[3, 12] as Excel.Range);
                                    lblYMaterial.Text = "Y. M. : " + userPrCode.Value.ToString();
                                    userPrCode = (worksheet.Cells[5, 15] as Excel.Range);
                                    lblPrdOrder.Text = "İş Emri : " + userPrCode.Value.ToString();
                                    n = dataGridViewMaterial.Rows.Add();
                                    dataGridViewMaterial.Rows[n].Cells[0].Value = (x + 1);
                                    userRange = (worksheet.Cells[indexRow, 1] as Excel.Range);
                                    dataGridViewMaterial.Rows[n].Cells[1].Value = userRange.Value;
                                    userRange = (worksheet.Cells[indexRow, 2] as Excel.Range);
                                    dataGridViewMaterial.Rows[n].Cells[2].Value = userRange.Value;
                                    userRange = (worksheet.Cells[indexRow, 3] as Excel.Range);
                                    dataGridViewMaterial.Rows[n].Cells[3].Value = userRange.Value;

                                    for (int j = 0; j < 16; j++)
                                    {
                                        m = dataGridView_list[j].Rows.Add();
                                        dataGridView_list[j].Rows[n].Cells[0].Value = (x + 1);
                                        userRange = (worksheet.Cells[indexRow, 4 + (j * 4)] as Excel.Range);
                                        if (userRange.Value != null)
                                        {
                                            dataGridView_list[j].Rows[n].Cells[1].Value = userRange.Value;
                                            userRange = (worksheet.Cells[indexRow, 7 + (j * 4)] as Excel.Range);
                                            dataGridView_list[j].Rows[n].Cells[2].Value = userRange.Value;
                                            dataGridView_list[j].Rows[n].Cells[0].Style.BackColor = Color.Green;
                                            dataGridView_list[j].Rows[n].Cells[0].Style.ForeColor = Color.White;
                                            dataGridView_list[j].Rows[n].Cells[1].Style.BackColor = Color.Green;
                                            dataGridView_list[j].Rows[n].Cells[1].Style.ForeColor = Color.White;
                                            dataGridView_list[j].Rows[n].Cells[2].Style.BackColor = Color.Green;
                                            dataGridView_list[j].Rows[n].Cells[2].Style.ForeColor = Color.White;
                                            userRange = (worksheet.Cells[indexRow, 4 + (j * 4)] as Excel.Range);
                                            string dataRcode = userRange.Value.ToString();
                                            dataRcodeArray = dataRcode.Split('/');
                                            rcode_list.Add(j.ToString() + "Y" + n.ToString() + dataRcodeArray[1]);

                                        }
                                        else
                                        {
                                            dataGridView_list[j].Rows[n].Cells[1].Value = "";
                                            dataGridView_list[j].Rows[n].Cells[2].Value = "";
                                        }
                                    }
                                    worksheet.Columns.AutoFit();
                                    indexRow++;
                                    x++;
                                }

                            }
                            if (constNumber == i + 1)
                            {
                                constNumber = constNumber + 17;
                                indexRow = indexRow + 8;
                            }
                            i++;
                        } while (flag);
                        top_code_list = new List<string>();
                        material_code_list = new List<string>();
                        description_list = new List<string>();
                        for (int j = 0; j < dataGridViewMaterial.Rows.Count; j++)
                        {
                            top_code_list.Add(dataGridViewMaterial.Rows[j].Cells[1].Value.ToString());
                            material_code_list.Add(dataGridViewMaterial.Rows[j].Cells[2].Value.ToString());
                            description_list.Add(dataGridViewMaterial.Rows[j].Cells[3].Value.ToString());
                        }

                        disableButton(3, true, btnFirstRecord, btnChange, btnDone);
                        sheet.Save();
                        isOpened = true;
                    }
                }
                else
                {
                    MessageBox.Show("Lütfen iş bitir butonuna basınız");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel yazma hatası" +ex);
            }
          

        }

        //Process.Start(@"c:\Windows\Sysnative\cmd.exe", "/c osk.exe" + "& exit")

        private void btnCardList_Click(object sender, EventArgs e)
        {
            lblDoneMessRemove();
            getCardList();

        }

        private int findFirstCharacter(params dynamic[] parameter)
        {
            int firstCharacter = 0;
            char[] characters = parameter[0].ToCharArray();
            for (int i = 0; i < parameter[0].Length; i++)
            {
                if (characters[0].ToString() == "M") 
                {
                    firstCharacter = 1;
                    break;
                }
                if (characters[i].ToString() == parameter[1].ToString() && Char.IsLetter(characters[i]))
                {
                    firstCharacter = i;
                    break;
                }
                if (Char.IsDigit(characters[i]))
                {
                    firstCharacter = i;
                    break;
                }
            }
            return firstCharacter;
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            lblDoneChangeMess(changeFlag, changeMessage);
        }

        private void lblDoneChangeMess(bool flag, string message)
        {
            timerMess++;
            if (timerMess == 1 && flag)
            {
                lblDoneMess(true, message);
                if (serialPortAlarm.IsOpen)
                {
                    byteArray[0] = 83;
                    serialPortAlarm.Write(byteArray, 0, 1);
                  // serialPortAlarm.Write("S");
                    serialBufferClear();
                }
            }
            else if (timerMess == 1 && !flag)
            {
                lblDoneMess(false, message);
                if (serialPortAlarm.IsOpen)
                {
                    byteArray[0] = 69;
                    serialPortAlarm.Write(byteArray, 0, 1);
                 // serialPortAlarm.Write("E");
                    serialBufferClear();
                }
            }
            if (timerMess == 2)
            {
                lblDoneMessRemove();
                timerMess = 0;
                timer3.Stop();
            }
        }

        private void btnExit_MouseLeave(object sender, EventArgs e)
        {
            btnExit.ForeColor = Color.Black;
        }
        private void settingsFirstRecordVisible()
        {
            if (txtFirstRecordBarcode.Visible == true)
            {
                btnFirstRecord.Text = "İLK KAYIT";
                btnFirstRecord.BackColor = Color.Gray;
                txtFirstRecordBarcode.Visible = false;
                lblFirstRecord.Visible = false;
                btnFirstRecordConfirm.Visible = false;
                btnChange.Enabled = true;
                btnDone.Enabled = true;
                btnMatch.Visible = false;
            }
            else
            {
                btnFirstRecord.Text = "BİTİR !";
                btnFirstRecord.BackColor = Color.Green;
                txtFirstRecordBarcode.Visible = true;
                lblFirstRecord.Visible = true;
                btnFirstRecordConfirm.Visible = false;
                lblFirstRecord.Text = "Lütfen Barkodu Okutunuz !";
                lblFirstRecord.ForeColor = Color.Black;
                btnChange.Enabled = false;
                btnDone.Enabled = false;
                btnMatch.Visible = true;
            }
            txtFirstRecordBarcode.Focus();
        }
        private void btnFirstRecord_Click(object sender, EventArgs e)
        {
            settingsFirstRecordVisible();
            lblDoneMessRemove();
        }
        private void settingChangeBarcodeVisible()
        {
            if (txtChangeBarcodeO.Visible == true)
            {
                btnChange.Text = "DEĞİŞTİR";
                btnChangeConfirmO.Text = "-";
                btnChangeConfirmN.Text = "-";
                txtChangeBarcodeO.Text = "";
                txtChangeBarcodeN.Text = "";
                btnChange.BackColor = Color.FromArgb(128, 64, 64);
                txtChangeBarcodeO.Visible = false;
                lblChangeBarcodeO.Visible = false;
                btnChangeConfirmO.Visible = false;
                txtChangeBarcodeN.Visible = false;
                lblChangeBarcodeN.Visible = false;
                btnChangeConfirmN.Visible = false;
                btnDone.Enabled = true;
                openToObject(txtFirstRecordBarcode, btnFirstRecord);
                openToObject(txtChangeBarcodeN, btnChangeConfirmN);
            }
            else
            {
                btnChange.Text = "BİTİR !";
                btnChangeConfirmO.Text = "-";
                btnChangeConfirmN.Text = "-";
                txtChangeBarcodeO.Text = "";
                txtChangeBarcodeN.Text = "";
                btnChange.BackColor = Color.Green;
                txtChangeBarcodeO.Visible = true;
                txtChangeBarcodeO.Enabled = true;
                lblChangeBarcodeO.Visible = true;
                btnChangeConfirmO.Visible = true;
                btnChangeConfirmO.Enabled = true;
                txtChangeBarcodeN.Visible = true;
                lblChangeBarcodeN.Visible = true;
                btnChangeConfirmN.Visible = true;
                btnDone.Enabled = false;
                checkBox1.Checked = false;
                lblChangeBarcodeO.Text = "Biten Bileşenin Barkodunu Okutunuz !";
                lblChangeBarcodeN.Text = "Yeni Bileşenin Barkodunu Okutunuz !";
                lblChangeBarcodeN.ForeColor = Color.White;
                lblChangeBarcodeO.ForeColor = Color.White;
                btnChangeConfirmO.ForeColor = Color.White;
                btnChangeConfirmN.ForeColor = Color.White;
                lockToObject(txtFirstRecordBarcode, btnFirstRecord);
                lockToObject(txtChangeBarcodeN, btnFirstRecord);

            }
            btnChangeConfirmO.BackColor = Color.ForestGreen;
            btnChangeConfirmN.BackColor = Color.ForestGreen;
            txtChangeBarcodeO.Focus();
        }

        private void btnChangeConfirmN_Click(object sender, EventArgs e)
        {
        }
        private void ConfirmN()
        {
            if (indexView == 0)
            {
                indexExcelCell = 8;
            }
            else if (indexView == 1)
            {
                indexExcelCell = 12;
            }
            else if (indexView == 2)
            {
                indexExcelCell = 16;
            }
            else if (indexView == 3)
            {
                indexExcelCell = 20;
            }
            else if (indexView == 4)
            {
                indexExcelCell = 24;
            }
            else if (indexView == 5)
            {
                indexExcelCell = 28;
            }
            else if (indexView == 6)
            {
                indexExcelCell = 32;
            }
            else if (indexView == 7)
            {
                indexExcelCell = 36;
            }
            else if (indexView == 8)
            {
                indexExcelCell = 40;
            }
            else if (indexView == 9)
            {
                indexExcelCell = 44;
            }
            else if (indexView == 10)
            {
                indexExcelCell = 48;
            }
            else if (indexView == 11)
            {
                indexExcelCell = 52;
            }
            else if (indexView == 12)
            {
                indexExcelCell = 56;
            }
            else if (indexView == 13)
            {
                indexExcelCell = 60;
            }
            else if (indexView == 14)
            {
                indexExcelCell = 64;
            }
            else if (indexView == 15)
            {
                indexExcelCell = 68;
            }
            if (barcodeSplit(txtChangeBarcodeN.Text))
            {
                btnChangeConfirmN.Text = epsCode;
                addComponent2Flag = false;
                if (!rCodeControl(rCode) && btnChangeConfirmO.Text==epsCode)
                {
                    addComponent(epsCode, dataGridView_list[indexView + 1], indexExcelCell, indexExcelCell + 3, indexView + 1);
                    letData = ("0" + (indexView + 1).ToString());
                    openToObject(txtChangeBarcodeO, btnChangeConfirmO);
                    lockToObject(txtChangeBarcodeN, btnFirstRecord);
                    setColorFindCells(dataGridView_list[indexView2], indexView3, indexView2 + 1, Color.ForestGreen, Color.White);
                    lblChangeBarcodeN.Text = "Yeni Malzemenin Barkodunu Okutunuz !";
                    lblChangeBarcodeN.ForeColor = Color.White;
                    timer3.Start();
                    changeFlag = true;
                    changeMessage = "Değişim Onaylandı";
                    btnChangeConfirmO.BackColor = Color.ForestGreen;
                    btnChangeConfirmN.BackColor = Color.ForestGreen;
                    txtChangeBarcodeO.Focus();

                }
                else
                {
                    lblChangeBarcodeN.Text = "Eşleşmeyen Malzeme Takıldı !";
                    lblChangeBarcodeN.ForeColor = Color.Red;
                    timer3.Start();
                    changeFlag = false;
                    changeMessage = "Malzeme Onaylanmadı";
                    txtChangeBarcodeN.Text = "";
                    btnChangeConfirmO.BackColor = Color.Red;
                    btnChangeConfirmN.BackColor = Color.Red;
                    txtChangeBarcodeN.Focus();
                }

            }
            else
            {
                lblChangeBarcodeN.Text = "Malzeme Bulunamadı !";
                lblChangeBarcodeN.ForeColor = Color.Red;
                timer3.Start();
                changeFlag = false;
                changeMessage = "Malzeme Onaylanmadı";
                txtChangeBarcodeN.Text = "";
                btnChangeConfirmO.BackColor = Color.Red;
                btnChangeConfirmN.BackColor = Color.Red;
                txtChangeBarcodeN.Focus();
            }
          
        }
        private void btnChangeConfirmO_Click(object sender, EventArgs e)
        {
        }
        private bool rCodeControl(string rcodeData)
        {
            bool flag = false;
            foreach (var item in rcode_list)
            {
                if (calculateComp(item) == rCode)
                {
                    flag = true;
                }

            }
            return flag;
        }
        private void ConfirmO()
        {
            bool flag = false;
            if (barcodeSplit(txtChangeBarcodeO.Text))
            {
                flag = true;
            }
          
            btnChangeConfirmO.Text = epsCode;
            if (flag&&findComponent())
            {
                lblChangeBarcodeO.Text = "Biten Malzemenin Barkodunu Okutunuz !";
                lblChangeBarcodeO.ForeColor = Color.White;
                templotNo = lotNo;
                tempEpsCode = epsCode;
                timer3.Start();
                changeFlag = true;
                changeMessage = "Malzeme Onaylandı";
                btnChangeConfirmO.BackColor = Color.ForestGreen;
            }
            else
            {
                lblChangeBarcodeO.Text = "Değişiklik Yapılamaz !";
                lblChangeBarcodeO.ForeColor = Color.Red;
                timer3.Start();
                changeFlag = false;
                changeMessage = "Malzeme Onaylanmadı";
                txtChangeBarcodeO.Text = "";
                btnChangeConfirmO.BackColor = Color.Red;
            }
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            settingChangeBarcodeVisible();
            lblDoneMessRemove();
        }

        public void parameterAdd(params dynamic[] duck)
        {
            foreach (var item in duck)
            {
                MessageBox.Show(item.ToString());
            }
        }
        private void btnDone_Click(object sender, EventArgs e)
        {
            if (isOpened)
            {
                sheet.Save();
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                workFinish();
                moveToFinishedFile(excelFile, finishedFile);
                rcode_list.Clear();
                if (Setting.Default.publicFolderUsingSit)
                {
                    try
                    {
                        string t20DestFolder = Setting.Default.publicFolder + t20ExcelName;
                        File.Copy(finishedFile, t20DestFolder);
                    }
                    catch (Exception)
                    {

                        MessageBox.Show("Ortak klasör bulunmadı !");
                    }
                }


            }
            else if (!searchAlgorithm)
            {
                workFinish();
            }
            else
            {
                lblDoneMess(false, "İşlem Başarısız");
            }
        }

        private void workFinish()
        {

            lblDoneMess(true, "Kayıt Başarılı !");
            dataGridViewMaterial.Rows.Clear();
            foreach (var item in dataGridView_list)
            {
                item.Rows.Clear();
            }

            disableButton(3, false, btnFirstRecord, btnChange, btnDone);
            disableButton(3, true, btnProductConfirm, btnRecover, btnFirstRecordConfirm);
            btnFirstRecordConfirm.Visible = true;
            btnProductConfirm.BackColor = Color.FromArgb(0, 0, 64);
            progressBarProductCode.Value = 0;
            lblProductMessage.Text = "";
            lblPrdOrder.Text = "İş Emri : ";
            lblYMaterial.Text = "Y. M. : ";
            progressBarProductCode.Visible = false;
            isOpened = false;
            flagRecover = false;
            if (serialPortAlarm.IsOpen)
            {
                byteArray[0] = 83;
                serialPortAlarm.Write(byteArray, 0, 1);
              //  serialPortAlarm.Write("S");
                serialBufferClear();
            }

        }

        private void lblDoneMessRemove()
        {
            pictureBoxConfirm.Visible = false;
            pictureBoxFailure.Visible = false;
            lblDoneMessage.Visible = false;
        }

        private void lblDoneMess(Boolean flag, string message)
        {
            if (flag)
            {
                pictureBoxConfirm.Visible = true;
                lblDoneMessage.Visible = true;
                lblDoneMessage.Text = message;
                lblDoneMessage.ForeColor = Color.Green;
                pictureBoxFailure.Visible = false;
            }
            else
            {
                pictureBoxFailure.Visible = true;
                lblDoneMessage.Visible = true;
                lblDoneMessage.Text = message;
                lblDoneMessage.ForeColor = Color.Red;
                pictureBoxConfirm.Visible = false;
            }
        }
        private bool barcodeSplit(string barcodeText)
        {
            barcodeData = barcodeText;
            string[] words = barcodeData.Split(',');
            try
            {
                if (barcodeText != "" && 50 <= barcodeData.Length && barcodeData.Length <= 90 && barcodeText.Substring(0, 1) == "R")
                {
                    for (int i = 0; i < words.Length; i++)
                    {
                        if (i == 0)
                        {
                            rCode = words[i];
                        }
                        if (i == 1)
                        {

                            epsCode = words[i];
                            epsCode = epsCode.Substring(findFirstCharacter(epsCode, "E"), epsCode.Length - 1);
                        }
                        else if (i == 2)
                        {
                            lotNo = words[i];
                            // findFirstCharacter(epsCode, 0);
                            lotNo = lotNo.Substring(1, lotNo.Length - 1);
                        }
                        else if (i == 4)
                        {
                            quantity = words[i];
                            findFirstCharacter(epsCode, 0);
                            quantity = quantity.Substring(findFirstCharacter(quantity, 0), quantity.Length - 3);
                        }
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;

            }
           
        }
        private void btnFirstRecordConfirm_Click(object sender, EventArgs e)
        {
            if (sqlConnection.State == ConnectionState.Open)
            {
                if (txtProductCode.Text != "")
                {

                    lblDoneMessRemove();
                    materialCompare();

                }
                else
                {
                    lblProductMessage.Text = "Lütfen Ürün Kodunu Giriniz !";
                    lblProductMessage.ForeColor = Color.Red;
                }
            }
            else
            {
                MessageBox.Show("Sql bağlatınsı kuruluyor! Lütfen tekrar deneyiniz...");
            }
            button1.Focus();

        }
        private void timerProduct(System.Windows.Forms.Timer timer, ProgressBar progressBar)
        {
            timer.Start();
            progressBar.Visible = true;
        }


        private void disableButton(byte buttonCount, Boolean flag, params dynamic[] parameter)
        {
            for (int i = 0; i < buttonCount; i++)
            {
                parameter[i].Enabled = flag;
            }

        }

        private void btnProductConfirm_Click(object sender, EventArgs e)
        {

            if (txtProductCode.Text != "")
            {
                btnProductConfirm.Enabled = false;
                btnProductConfirm.BackColor = Color.FromArgb(190, 151, 144);
                sqlToListData(txtProductCode.Text);
                if (sqlConnection.State == ConnectionState.Open)
                {
                    if (flagProduct)
                    {
                        titleExcel();
                        getExcelFilesURL("Ref_Oto_Dizgi_Izlenebilirlik_Listesi_Referans", "Ref_Oto_Dizgi_Izlenebilirlik_Listesi");
                        dataWriteToExcel(excelFile);
                        timer3.Start();
                        changeFlag = true;
                        changeMessage = "Ürün Eklendi";
                        searchAlgorithm = true;

                        dataGridViewListCall(4, dataGridView1, dataGridView2, dataGridView3, dataGridView4);
                        disableButton(3, true, btnFirstRecord, btnChange, btnDone);
                        disableButton(2, false, btnRecover, btnFirstRecordConfirm);
                        isOpened = true;
                        if (sqlConnection.State == ConnectionState.Open)
                        {
                            sqlConnection.Close();
                        }
                    }
                    timerProduct(timer1, progressBarProductCode);

                }
                else
                {
                    MessageBox.Show("Ağ bağlantı sorunu");
                }
            }
            else
            {
                lblProductMessage.Text = "Lütfen Ürün Kodunu Giriniz !";
                lblProductMessage.ForeColor = Color.Red;
            }

            lblDoneMessRemove();
        }
        private void serialBufferClear()
        {
            serialPortAlarm.DiscardInBuffer();
            serialPortAlarm.DiscardOutBuffer();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection.State == ConnectionState.Open)
            {
                sqlConnection.Close();

            }
            if (serialPortAlarm.IsOpen)
            {
                serialPortAlarm.Close();
            }

        }

    }

}
