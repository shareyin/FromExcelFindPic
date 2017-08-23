using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace FormExcelFindPic
{
    public partial class labelCat : Form
    {
        public labelCat()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog filedialog = new OpenFileDialog();
            string FileName = "";

            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                FileName = filedialog.FileName;
                //((DataGridViewTextBoxColumn)dataGridView1.Columns[5]).MaxInputLength=1024;
                dataGridView1.DataSource = GetExcelData(FileName);
                dataGridView1.DataMember = "[Sheet1$]";

                for (int count = 0; (count <= (dataGridView1.Rows.Count - 1)); count++)
                {
                    dataGridView1.Rows[count].HeaderCell.Value = (count + 1).ToString();
                }
            }
        }
        //获取Excel
        public static DataSet GetExcelData(string str)
        {
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + str + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'";
            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = " SELECT * FROM [Sheet1$]";
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            DataSet myDataSet = new DataSet();
            myCommand.Fill(myDataSet, "[Sheet1$]");
            myConn.Close();
            return myDataSet;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int index = this.dataGridView1.CurrentRow.Index;
                string s_mJGCarType = (this.dataGridView1.Rows[index].Cells[2].Value).ToString();//序号
                string s_mCamPlateNum = (this.dataGridView1.Rows[index].Cells[5].Value).ToString();//车型
                string s_mOBUCarType = (this.dataGridView1.Rows[index].Cells[10].Value).ToString();//车牌颜色
                string s_mOBUPlateNum = (this.dataGridView1.Rows[index].Cells[9].Value).ToString();//车牌号码
                string s_mImagePath = (this.dataGridView1.Rows[index].Cells[7].Value).ToString();//图片路径

                Showpic(s_mJGCarType, s_mCamPlateNum, s_mOBUCarType, s_mOBUPlateNum, s_mImagePath);
                //OpenIe(s_mImagePath);
                //Process.Start(s_mImagePath);
                //GridDouclickShowPicDialog gdspd = new GridDouclickShowPicDialog(vehtype, getplateno, imagepath, forcetime, ovehtype, oplatenumber, laser_vehlenth, laser_vehheight, this);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
                Console.WriteLine(ex.ToString());
            }
        }
        //显示图片
        public void Showpic(string s_mJGCarType, string s_mCamPlateNum, string s_mOBUCarType, string s_mOBUPlateNum, string s_mImagePath)
        {
            string fileName = System.IO.Path.GetFileName(s_mImagePath);
            if (textBoxfile.Text != "")
            {
                if (radioButton1.Checked)
                {
                    s_mImagePath = textBoxfile.Text +"\\"+ fileName;
                }
                else
                {
                    s_mImagePath = textBoxfile.Text + s_mImagePath; 
                }
                
            }
            if (s_mImagePath != null || s_mImagePath != "" || s_mImagePath == "未知")
            {
                try
                {
                    pictureBox1.Load(s_mImagePath);
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("该记录无对应图片");
                }
            }
            else
            {
                MessageBox.Show("该记录无对应图片");
            }
            labelCarNum.Text = s_mCamPlateNum;
            labelCarType.Text = s_mJGCarType;
            labelOBUNum.Text = s_mOBUPlateNum;
            labelOBUType.Text = s_mOBUCarType;
            //MessageBox.Show(s_mImagePath);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            textBoxfile.Text = path;
            //return path;
        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            string PicPath = pictureBox1.ImageLocation;
            Process.Start(PicPath);
        }
    }
}
