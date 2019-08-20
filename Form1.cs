using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace opexcl
{
    public partial class Form1 : Form
    {
        string path3,addr,addr1;
        way way1 = new way();
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
            path3 = System.IO.Directory.GetCurrentDirectory();//获取软件路径
            addr = path3 + "\\数据文件.xls";
            addr1= path3 + "\\复制数据文件.xls";
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            
            if (way1.getexcl(addr) == 1)
                label1.Text = "已有";
            else label1.Text = "没有";
                
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (label1.Text == "已有")
                return;
            way1.createexcl(addr);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            try
            {
                File.Delete(addr);
            }catch(Exception)
            {
                label1.Text = "删除成功";
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            way1.openingexcl(addr);
            label1.Text = "excl打开成功";

        }

        private void Button5_Click(object sender, EventArgs e)
        {
            dt= way1.getwritetodt();
            dataGridView1.DataSource = dt;
            label1.Text = "数据读取成功";

        }

        private void Button6_Click(object sender, EventArgs e)
        {
            way1.saveexcl(addr);
            label1.Text = "已关闭并保存excl表";
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            label7.Text = way1.geth();
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            label8.Text = way1.getl();
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            List<string> temp = new List<string>();
            foreach (var item in this.listBox1.Items)
            {
                temp.Add(item.ToString());
            }
            //MessageBox.Show(a.ToString());
            way1.WriteTitle(temp);
            label1.Text = "标题添加成功";
        }

        private void Button12_Click(object sender, EventArgs e)
        {
            //首先判断列表框中的项是否大于0
            if(listBox1.Items.Count > 0)
            {
                //清空所有项
                listBox1.Items.Clear();
            }
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            if(textBox1.Text=="")
            {
                MessageBox.Show("标题不能为空");
                return; 
            }
            listBox1.Items.Add(textBox1.Text);
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            if(listBox1.SelectedIndex.ToString()=="-1")
            {
                MessageBox.Show("为选中任何标题");
                return;
            }
            listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void Button13_Click(object sender, EventArgs e)
        {
            way1.closing();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void Button14_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("标题不能为空");
                return;
            }
            listBox2.Items.Add(textBox2.Text);
        }

        private void Button15_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex.ToString() == "-1")
            {
                MessageBox.Show("为选中任何标题");
                return;
            }
            listBox2.Items.Remove(listBox2.SelectedItem);
        }

        private void Button16_Click(object sender, EventArgs e)
        {
            //首先判断列表框中的项是否大于0
            if (listBox2.Items.Count > 0)
            {
                //清空所有项
                listBox2.Items.Clear();
            }
        }

        private void Button17_Click(object sender, EventArgs e)
        {
            List<string> temp = new List<string>();
            foreach (var item in this.listBox2.Items)
            {
                temp.Add(item.ToString());
            }
            //MessageBox.Show(a.ToString());
            way1.WriteData(temp);
            label1.Text = "数据添加成功";
        }

        private void Button18_Click(object sender, EventArgs e)
        {
            MessageBox.Show(way1.SelectH(textBox3.Text, comboBox1.Text).ToString());
        }

        private void Button20_Click(object sender, EventArgs e)
        {
            way1.RemoveH(Convert.ToInt32(textBox4.Text));
            label9.Text = "删除行成功";
        }

        private void Button19_Click(object sender, EventArgs e)
        {
            way1.DataTableToExcel(addr1,dt,"Sheet2",false);
        }

        //删除文件
        public static void DelectDir(string srcPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }
    }
}
