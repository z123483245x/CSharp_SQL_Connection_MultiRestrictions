using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.IO;

namespace CSharp_SQL3
{

    public partial class Form1 : Form
    {
        SaveFileDialog sdlgCSV = new SaveFileDialog();
        DataTable dataTable = new DataTable();
        Label lblSearch = new Label();
        TextBox tbSearch = new TextBox();
        Button btnExport = new System.Windows.Forms.Button();
        Button btnSearch = new Button();
        DataGridView dgv1 = new DataGridView();
        TextBox tbCheck = new TextBox();
        public Form1()
        {
            InitializeComponent();
            this.Text = "生產報表";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterScreen;

            lblSearch.Text = "請輸入客戶編號、生產機種的料號號碼或工單號碼";
            lblSearch.Location = new Point(40, 30);
            lblSearch.AutoSize = true;

            tbSearch.Location = new Point(40, 50);
            tbSearch.Size = new Size(280, 00);


            btnSearch.Location = new Point(360, 13);
            btnSearch.Text = "查詢";
            btnSearch.Size = new Size(80, 30);
            btnSearch.BackColor = Color.SlateGray;
            btnSearch.FlatStyle = FlatStyle.Popup;
            btnSearch.Click += btnSearch_Click;

            btnExport.Location = new Point(360, 50);
            btnExport.Text = "匯出";
            btnExport.Size = new Size(80, 30);
            btnExport.FlatStyle = FlatStyle.Popup;
            btnExport.BackColor = Color.SlateGray;
            btnExport.Click += btnExport_Click;

            dgv1.Size = new Size(440, 200);
            dgv1.Location = new Point(23, 105);
            dgv1.BorderStyle = BorderStyle.FixedSingle;
            dgv1.BackColor = Color.AntiqueWhite;
            dgv1.RowHeadersVisible = false;
            dgv1.AutoGenerateColumns = true;
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            tbCheck.BackColor = Color.Bisque;
            tbCheck.Size = new Size(500, 0);
            tbCheck.Dock = DockStyle.Bottom;
            tbCheck.BorderStyle = BorderStyle.Fixed3D;

            tbSearch.KeyPress += tbSearch_KeyPress;

            this.Controls.Add(dgv1);
            this.Controls.Add(btnExport);
            this.Controls.Add(lblSearch);
            this.Controls.Add(tbSearch);
            this.Controls.Add(btnSearch);
            this.Controls.Add(tbCheck);
        }

        private void tbSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnSearch.PerformClick();
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            dataTable.Clear();
            string connectionString = "DATA SOURCE=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.1)(PORT=1521)))(CONNECT_DATA=(SID = MIS)));PERSIST SECURITY INFO=True;USER ID=MIS;PASSWORD=OracleMis;";
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sqlQuery = "SELECT mse0017.f002 as 客戶別 , MISTEST.mse0008.f002 as 生產機種料號 , MISTEST.mse0017.f003 as 工單號碼"+
                        " , MISTEST.mse0017.f005 as 工單號碼生產數量 , MISTEST.mse0017.f008 as 有效否 , MISTEST.mse0017.f012 as 結案否"+
                        " , TO_CHAR(MISTEST.mse0017.f011,'yyyy/MM/dd HH24:MI:SS') as 工單開立時間"+" , MISTEST.sys0003.f002 as 建立人員的工號"+
                        " FROM MISTEST.mse0017 JOIN MISTEST.mse0004 ON MISTEST.mse0017.f002 LIKE MISTEST.mse0004.f001 JOIN MISTEST.mse0008 "+
                        " ON MISTEST.mse0017.f005 LIKE"+" MISTEST.mse0008.f001 JOIN MISTEST.sys0003 ON MISTEST.mse0008.f011 LIKE MISTEST.sys0003.f001 "+
                        " WHERE MISTEST.mse0017.f002 LIKE :WorkOrderNumber OR mse0008.f002 LIKE :WorkOrderNumber OR mse0017.f003 LIKE :WorkOrderNumber ";
                    OracleCommand command = new OracleCommand(sqlQuery, connection);
                    command.Parameters.Add(new OracleParameter(":WorkOrderNumber", OracleDbType.Varchar2)).Value = tbSearch.Text;
                    OracleDataReader reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        dataTable.Load(reader);
                        dgv1.DataSource = dataTable;
                        reader.Close();
                        tbCheck.Text = "查詢完成";
                    }
                    else if (tbSearch.Text == "")
                    {
                        tbCheck.Text = "查詢失敗";
                        MessageBox.Show("請輸入客戶編號、生產機種的料號號碼或工單號碼");
                    }
                    else
                    {
                        tbCheck.Text = "查詢失敗";
                        MessageBox.Show("目前查詢的工單不存在");
                    }
                    reader.Dispose();
                }
                catch (Exception ex)
                {
                    tbCheck.Text = "查詢失敗";
                    MessageBox.Show("連線或查詢過程中發生錯誤 : " + ex.Message);
                    connection.Close();
                    connection.Dispose();
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            sdlgCSV.Filter = "CSV文件|*.csv";
            sdlgCSV.Title = "匯出CSV報表";

            if (dataTable.Rows.Count > 0)
            {
                if (sdlgCSV.ShowDialog() == DialogResult.OK)
                {
                    string filePath = sdlgCSV.FileName;
                    try
                    {
                        StringBuilder stringBulider = new StringBuilder();
                        foreach (DataColumn dataColumn in dataTable.Columns)
                        {
                            stringBulider.Append(dataColumn.ColumnName + ",");
                        }
                        stringBulider.AppendLine();
                        foreach (DataRow dataRow in dataTable.Rows)
                        {
                            foreach (var item in dataRow.ItemArray)
                            {
                                stringBulider.Append(item.ToString() + ",");
                            }
                            stringBulider.AppendLine();
                        }
                        File.WriteAllText(filePath, stringBulider.ToString(), Encoding.UTF8);
                        tbCheck.Text = "匯出成功";
                        MessageBox.Show("csv檔案已匯出");
                    }
                    catch (Exception ex)
                    {
                        tbCheck.Text = "匯出失敗";
                        MessageBox.Show("csv檔案匯出時發生錯誤 : " + ex.Message);
                    }
                }
            }
            else
            {
                tbCheck.Text = "匯出失敗";
                MessageBox.Show("沒有可匯出的資料");
            }
        }


    }
}

