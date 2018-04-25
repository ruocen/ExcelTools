using System;
using System.Collections.Generic;
using System.Windows.Forms;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections;
using NPOI.XSSF.UserModel;
using System.Linq;

namespace ExcelTools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<ExcelRows> left = new List<ExcelRows>();
        List<ExcelRows> right = new List<ExcelRows>();
        List<string> leftHeaders = new List<string>();
        List<string> rightHeaders = new List<string>();

        Dictionary<string, int> link = new Dictionary<string, int>();
        string leftFileName = "";
        string rightFileName = "";
        string leftfileExt = "";
        string rightfileExt = "";

        string sExportFilePath = "";
        private void button1_Click(object sender, EventArgs e)
        {
            left.Clear();
            OpenFileDialog fileDialog = File();
            if (fileDialog.FileName == "") return;
            label3.Text = fileDialog.FileName;

            ExcelHelper.Import(fileDialog.FileName, fileDialog.SafeFileNames[0], out leftFileName, out leftfileExt, out leftHeaders, out left, out link, "left");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            right.Clear();
            link.Clear();
            OpenFileDialog fileDialog = File();
            if (fileDialog.FileName == "") return;
            label4.Text = fileDialog.FileName;

            ExcelHelper.Import(fileDialog.FileName, fileDialog.SafeFileNames[0], out rightFileName, out rightfileExt, out rightHeaders, out right, out link, "right");
        }
        /// <summary>
        /// 合并
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            if (rightfileExt != leftfileExt)
            {
                MessageBox.Show("格式不一样,请选择相同格式文件");
            }
            if (left.Count == 0)
            {
                MessageBox.Show("请选择左边文件");
                return;
            }
            if (right.Count == 0)
            {
                MessageBox.Show("请选择右边文件");
                return;
            }
            IWorkbook workbook;
            string fileExt = "";
            if (leftfileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook();
                fileExt = ".xlsx";
            }
            else if (leftfileExt == ".xls")
            {
                workbook = new HSSFWorkbook();
                fileExt = ".xls";
            }
            else
            {
                workbook = null;
                return;
            }
            ISheet sheet = workbook.CreateSheet(leftFileName + "|" + rightFileName);
            //表头
            IRow row = sheet.CreateRow(0);
            int j = 0;
            List<string> headers = leftHeaders.Concat(rightHeaders).ToList();
            System.Collections.Generic.HashSet<string> newheaders = new System.Collections.Generic.HashSet<string>(headers);
            foreach (var header in newheaders)
            {
                ICell cell = row.CreateCell(j);
                cell.SetCellValue(header);
                j++;
            }
            var used = new Dictionary<string, bool>();
            //数据
            int leftNum = 0;
            foreach (var rowleft in left)
            {
                IRow row1 = sheet.CreateRow(leftNum + 1);
                if (link.ContainsKey(rowleft.Code))
                {
                    var i = link[rowleft.Code];
                    var r = right[i];
                    ICell cell = null;
                    for (int k = 0; k < rowleft.ExcelCells.Count; k++)
                    {
                        cell = row1.CreateCell(k);
                        cell.SetCellValue(rowleft.ExcelCells[k]);
                    }
                    for (int n = 1; n < r.ExcelCells.Count; n++)
                    {
                        cell = row1.CreateCell(n + 2);
                        cell.SetCellValue(r.ExcelCells[n]);
                    }
                    used[rowleft.Code] = true;
                }
                else
                {
                    //
                }
                leftNum++;
            }

            foreach (var rowrigth in right)
            {
                if (used.ContainsKey(rowrigth.Code) == false)
                {
                    string noMatch = "";
                    for (int i = 0; i < rowrigth.ExcelCells.Count; i++)
                    {
                        noMatch += "|" + rowrigth.ExcelCells[i];
                    }
                    textBox2.AppendText(noMatch.Substring(1));
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.ScrollToCaret();
                }
            }
            foreach (var rowleft in left)
            {
                if (used.ContainsKey(rowleft.Code) == false)
                {
                    string noMatch = "";
                    for (int i = 0; i < rowleft.ExcelCells.Count; i++)
                    {
                        noMatch += "|" + rowleft.ExcelCells[i];
                    }
                    textBox1.AppendText(noMatch.Substring(1));
                    textBox1.AppendText(Environment.NewLine);
                    textBox1.ScrollToCaret();
                }
            }
            mergeSuccess.Text = "合并成功!";

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            string sExportFileName = "";
            
            string sFileName = leftFileName + "_" + rightFileName;
            string sWebBasePath = AppDomain.CurrentDomain.BaseDirectory;
            string sExportDir = sWebBasePath + "Export";
            sExportFileName = sFileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + fileExt;
            sExportFilePath = sExportDir + "\\" + sExportFileName;
            if (!Directory.Exists(sExportDir))
                Directory.CreateDirectory(sExportDir);
            using (FileStream fs = new FileStream(sExportFilePath, FileMode.Create, FileAccess.Write))
            {
                textBox3.Text = sExportFilePath;
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
                fs.Close();
                if (MessageBox.Show(this, "是否打开文件？", "打开", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(sExportFilePath);
                }
            }
        }


        private OpenFileDialog File()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string extension = Path.GetExtension(fileDialog.FileName);//文件后缀
                string[] str = new string[] { ".xls",".xlsx" };
                if (!((IList)str).Contains(extension))
                {
                    MessageBox.Show("仅能上传.xls和.xlsx格式的文件！");
                    fileDialog.FileName = "";
                    return fileDialog;
                }

                FileInfo fileInfo = new FileInfo(fileDialog.FileName);
                if (fileInfo.Length > 1024 * 1000000)
                {
                    MessageBox.Show("上传的文件不能大于20K");
                    fileDialog.FileName = "";
                    return fileDialog;
                }
                return fileDialog;
            }
            fileDialog.FileName = "";
            return fileDialog;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            psi.Arguments = "/e,/select," + sExportFilePath;
            System.Diagnostics.Process.Start(psi);
        }
    }
}
