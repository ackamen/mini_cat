using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace 使用oledb读取excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void 上传_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    foreach (var item in openFileDialog1.FileNames)
                    {
                        //后缀名   解析器  版本
                        //xls       Excel 8.0          97/00/02/03 
                        //xlsx      Excel 12.0 Xml     07/10/13 
                        //xlsb      Excel 12.0         07/10/13 
                        //xlsm      Excel 12.0 Macro   07/10/13 
                        var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", item);
                        var adapter = new OleDbDataAdapter("SELECT * FROM [sheet$]", connectionString);
                        var ds = new DataSet();
                        adapter.Fill(ds, "tableName");
                        DataTable data = ds.Tables["tableName"];

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
