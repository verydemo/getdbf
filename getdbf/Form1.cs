using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace getdbf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string defaultDir = Path.GetDirectoryName(ofd.FileName);
                string Tablename = Path.GetFileNameWithoutExtension(ofd.FileName);
                var dta=GetDbfDataByODBC(Tablename, defaultDir);
                //生成的Table结构
                string[] cstr1 = new string[] { "PAPERCODE", "STUID", "ITEMNO", "SCORE", "PCID" };
                //需转换的列  "T4", "T51", "T52", "T6" 转成 "ITEMNO", "SCORE"
                string[] cstr2 = new string[] { "T4^t4", "T51^t51", "T52^t52", "T6^t6" };
                Dictionary<string,string> t= cstr2.ToDictionary(a => a.Split('^')[0], a => a.Split('^')[1]);
                //不需转换的列
                string[] cstr3 = new string[] { "STUID", "PCID", "PAPERCODE" };
                //var t = GetDynamicClassBydt(dta);
                var dta1=rowchangeColumn(dta, cstr1, t, cstr3);
                if (File.Exists("2.txt")) File.Delete("2.txt");
                StreamWriter sw1 = File.AppendText("2.txt");
                foreach (DataRow a in dta1.Rows)
                {
                    string str= "";
                    int num = 0;

                    foreach (DataColumn b in dta1.Columns)
                    {
                        num++;
                        if (num == 1) str = a[b.ColumnName].ToString();
                        else str = str + ',' + a[b.ColumnName].ToString();
                    }
                    sw1.WriteLine(str);
                }
                sw1.Close();
                //dataGridView1.DataSource = dta1;
            }
        }

        /// <summary>
        /// 读取dbf文件中的数据 如所在目录为E:\manman\aa.dbf
        /// </summary>
        /// <param name="tableName">要读取的表名（也就是文件名）aa</param>
        /// <param name="defaultDir">文件所在的路径不包括文件名E:\manman</param>
        /// <returns></returns>
        private DataTable GetDbfDataByODBC(string tableName, string defaultDir)
        {
            OdbcConnection oConn = new System.Data.Odbc.OdbcConnection();
            oConn.ConnectionString = "Driver={Microsoft dBase Driver (*.dbf)};DefaultDir=" + defaultDir;
            oConn.Open();
            OdbcDataAdapter odbcDataAdapt = new OdbcDataAdapter("select * from  " + tableName, oConn.ConnectionString);
            DataTable dtData = new DataTable();
            dtData.TableName = "gb";
            odbcDataAdapt.Fill(dtData);
            oConn.Close();
            return dtData;
        }

        //行转列
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="cstr1">生成的Table结构</param>
        /// <param name="cstr2">需转换的列</param>
        /// <param name="cstr3">不需转换的列</param>
        /// <returns></returns>
        private DataTable rowchangeColumn(DataTable dt, string[] cstr1, Dictionary<string, string> t, string[] cstr3)
        {
            DataTable new_DataTable = new DataTable();
            foreach (string c in cstr1)
            {
                new_DataTable.Columns.Add(c);
            }
            foreach (DataRow c in dt.Rows)
            {               
                foreach (DataColumn a2 in dt.Columns)
                {
                    if (t.ContainsKey(a2.ColumnName))
                    {
                        DataRow row = new_DataTable.NewRow();
                        foreach (DataColumn a1 in dt.Columns)
                        {
                            if (cstr3.Contains(a1.ColumnName))
                            {
                                row[a1.ColumnName] = c[a1.ColumnName];
                            }
                        }
                        row["ITEMNO"] = t[a2.ColumnName];
                        row["SCORE"] = c[a2.ColumnName];
                        new_DataTable.Rows.Add(row);
                    }
                }
            }
            return new_DataTable;
        }

        /// <summary>
        /// 使用dynamic根据DataTable的列名自动添加属性并赋值
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns> 
        public static Object GetDynamicClassBydt(DataTable dt)
        {
            dynamic d = new System.Dynamic.ExpandoObject();
            var a = d as System.Collections.Generic.ICollection<System.Collections.Generic.KeyValuePair<string, object>>;
            //创建属性，并赋值。

                //(d as System.Collections.Generic.ICollection<System.Collections.Generic.KeyValuePair<string, object>>).Add(new System.Collections.Generic.KeyValuePair<string, object>(cl.ColumnName, dt.Rows[0][cl.ColumnName].ToString()));
                
           //for (int i=0; i<dt.Rows.Count;i++)
           //{
                foreach (DataColumn cl in dt.Columns)
                {
                    a.Add(new System.Collections.Generic.KeyValuePair<string, object>(cl.ColumnName, dt.Rows[0][cl.ColumnName].ToString()));
                }
           //}
            return a;
        }

    }
}
