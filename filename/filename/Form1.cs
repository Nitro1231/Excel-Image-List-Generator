using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace filename {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            Microsoft.Office.Interop.Excel.Application ap = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = ap.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;
            try {
                DirectoryInfo d = new DirectoryInfo(@"C:\NitroSoft\Screenshot\Parts");//Assuming Test is your Folder
                FileInfo[] Files = d.GetFiles("*"); //Getting Text files

                double height = 0;
                for(int i = 2; i < Files.Length + 2; i++) {
                    Image img = Image.FromFile(Files[i - 2].FullName);
                    double wss = img.Width, hs = img.Height;
                    while (wss >= 250) {
                        wss /= 1.01;
                        hs /= 1.01;
                    }
                    int w = (int)wss, h = (int)hs;

                    if (i % 2 == 0) {
                        ws.Range[$"b{i}"].Value = Files[i - 2].Name;
                        float Left = (float)(double)ws.Range[$"b{i + 1}"].Left;
                        float Top  = (float)(double)ws.Range[$"b{i + 1}"].Top;
                        ws.Shapes.AddPicture(Files[i - 2].FullName, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, w, h);
                        ws.Range[$"b{i + 1}"].ColumnWidth = w / 6;
                        ws.Range[$"b{i + 1}"].RowHeight = hs + 2;
                        height = hs + 2;
                    } else {
                        ws.Range[$"c{i - 1}"].Value = Files[i - 2].Name;
                        float Left = (float)(double)ws.Range[$"c{i}"].Left;
                        float Top = (float)(double)ws.Range[$"c{i}"].Top;
                        ws.Shapes.AddPicture(Files[i - 2].FullName, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, w, h);
                        ws.Range[$"c{i}"].ColumnWidth = w / 6;
                        if (hs + 2 < height)
                            ws.Range[$"c{i}"].RowHeight = height;
                        else
                            ws.Range[$"c{i}"].RowHeight = hs + 2;
                    }
                }

                string filepath = @"C:\NitroSoft\Screenshot\ex.xlsx";
                if (File.Exists(filepath)) File.Delete(filepath);
                wb.SaveAs(filepath);
            } catch (Exception ex) {
                if (wb != null) wb.Close();
                if (ap != null) ap.Quit();
                // process error message
                label1.Text = ex.ToString();
                return;
            }

            if (wb != null) wb.Close();
            label1.Text = "job finished.";
        }
    }
}
