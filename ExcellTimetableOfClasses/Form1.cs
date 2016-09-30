using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellTimetableOfClasses
{

    public partial class Form1 : Form
    {

        Excel.Application ObjExcel = new Excel.Application();
        Excel.Application ObjExcel2 = new Excel.Application();
        char[] alpha = "ABCDEF".ToCharArray(); //ABCDEFGHIJKLMNOPQRSTUVWXYZ

        public Form1()
        {
            InitializeComponent();
        }

        private void btnGetAll_Click(object sender, EventArgs e)
        {
            if (txtNewFile.Text == "")
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                    return;
                else
                    txtNewFile.Text = openFileDialog1.FileName; //новый

            tabControl1.SelectedTab = tabPage2;
            richTextBox2.Clear();
            char[] charsToTrim = { '*', ' ', '_', '\n' };
            char[] ColumnsToSplit = { 'A', 'B', 'E', 'F' };

            //int[] Indexes = { };
            string group = "";
            int CountToExit = 0;
            string[] ArrayOfChars = { };
            //-----------------------------------------------------
            //var e;
            /*try
            {*/
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", false, false, 0, true, false, false);
            var ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[numEdt.Value];
            Excel.Worksheet sheet = ObjExcel.ActiveWorkbook.ActiveSheet;
            /*}
            catch {
                MessageBox.Show("Произошла ошибка при попытке открыть файл");
            }*/
            progressBar1.Value = 0;

            var ResultCell = "";
            for (int i = 7; i < 200; i++)
            {
                progressBar1.PerformStep();
                var IsGroup = (ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim) != ""
                               && ObjWorkSheet.Range["B" + i.ToString(), "B" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                               && ObjWorkSheet.Range["C" + i.ToString(), "C" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                               && group != ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim));
                if (IsGroup || CountToExit > 9)
                {
                    richTextBox2.Text = richTextBox2.Text + group;
                    if (richTextBox2.Text.Length > 0)
                        group = "\nГруппа " + ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim) + '\n';
                    else
                        group = "Группа " + ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim) + '\n';


                    var words = ResultCell.Trim(charsToTrim).Split('_');
                    if (words.Length > 1)
                    {
                        IEnumerable<string> query = from word in words
                                                    orderby word.Substring(3, 2), word.Substring(0, 2), word.Substring(6, 2)
                                                    select word;

                        foreach (string str in query)
                            richTextBox2.Text = richTextBox2.Text + str + "\n";
                    }
                    ResultCell = "";
                    if (IsGroup)
                        continue;
                }
                if (ObjWorkSheet.Range['A' + i.ToString(), 'A' + i.ToString()].Text.ToString() == "Число")
                    continue;
                if (CountToExit > 9)
                {
                    break;
                }
                if (ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                     && ObjWorkSheet.Range["B" + i.ToString(), "B" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                     && ObjWorkSheet.Range["C" + i.ToString(), "C" + i.ToString()].Text.ToString().Trim(charsToTrim) == "")
                {
                    CountToExit++;
                    if (CountToExit > 5)
                        progressBar1.Value = 180 + (2 * CountToExit);
                    continue;
                }

                var CellA = ObjWorkSheet.Range['A' + i.ToString(), 'A' + i.ToString()].Text.ToString().Trim(charsToTrim);
                var NumOfAEnter = CellA.Split('\n').Length;

                /*Всякие проверки окончены. Погнали*/
                for (var j = 0; j < NumOfAEnter; j++)
                {
                    foreach (var d in alpha)
                    {
                        var Cell = ObjWorkSheet.Range[d + i.ToString(), d + i.ToString()].Text.ToString().Trim(charsToTrim);
                        ArrayOfChars = Cell.Split('\n');

                        if (ArrayOfChars.Length == NumOfAEnter && ColumnsToSplit.Contains(d))
                        {
                            var Cell1 = ArrayOfChars[j].Replace("//", "/").Replace("_", "/").Trim(charsToTrim);
                            if (d == 'B' && Cell1.Length > 12)
                            {
                                ResultCell = ResultCell + Cell1.Substring(Cell1.Length - 11) + " ";
                            }
                            else
                            {
                                ResultCell = ResultCell + Cell1 + " ";
                            }
                        }
                        else
                            ResultCell = ResultCell +
                                         Cell.Replace("\n", "/").Replace("//", "/").Replace("_", "/").Trim(charsToTrim) +
                                         " ";

                        /*закончили формирование строки*/
                    }
                    ResultCell = ResultCell + "_";
                }
                /*ArrayOfResult[z].DateIndex = 0;
                ArrayOfResult[z].Row = ArrayOfResult[z].Row + ResultCell;
                t++;
                richTextBox2.Text = richTextBox2.Text + ResultCell;
                ResultCell = "";*/
            }
        }

        private void btnDiffer_Click(object sender, EventArgs e)
        {
            if (txtNewFile.Text == "")
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                    return;
            if (txtOldFile.Text == "")
                if (openFileDialog2.ShowDialog() != DialogResult.OK)
                    return;
            txtOldFile.Text = openFileDialog2.FileName; //старый
            txtNewFile.Text = openFileDialog1.FileName; //новый

            tabControl1.SelectedTab = tabPage1;
            char[] charsToTrim = { '*', ' ', '\n' };
            richTextBox1.Clear();
            progressBar1.Value = 0;
            //Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", false, false, 0, true, false, false);
            var ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[numEdt.Value];

            ////Excel.Application ObjExcel2 = new Excel.Application();
            Excel.Workbook ObjWorkBook2 = ObjExcel2.Workbooks.Open(openFileDialog2.FileName, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", false, false, 0, true, false, false);
            var ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook2.Sheets[numEdt.Value];
            Application.DoEvents();
            richTextBox1.Text = "Старый файл:" + openFileDialog2.FileName;
            richTextBox1.Text = richTextBox1.Text + "\nНовый файл:" + openFileDialog1.FileName;
            string Group = "";
            int NeedToSetGroup = 0;
            bool FindSomething = false;
            int CountToExit = 0;
            bool IsDiff = false;

            Application.DoEvents();

            for (var i = 7; i < 200; i++)
            {
                progressBar1.PerformStep();
                //Excel.Range range1 = ObjWorkSheet.get_Range('A' + i.ToString(), 'A' + i.ToString());
                //Excel.Range range2 = ObjWorkSheet2.get_Range('A' + i.ToString(), 'A' + i.ToString());
                if (ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim) != ""
                    && ObjWorkSheet.Range["B" + i.ToString(), "B" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                    && ObjWorkSheet.Range["C" + i.ToString(), "C" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                    && Group != ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim))
                {
                    NeedToSetGroup = 1;
                    Group = "\nГруппа " + ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim);
                }
                if (CountToExit > 9)
                {
                    break;
                }
                if (ObjWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                     && ObjWorkSheet.Range["B" + i.ToString(), "B" + i.ToString()].Text.ToString().Trim(charsToTrim) == ""
                     && ObjWorkSheet.Range["C" + i.ToString(), "C" + i.ToString()].Text.ToString().Trim(charsToTrim) == "")
                {
                    CountToExit++;
                    if (CountToExit > 5)
                        progressBar1.Value = 180 + (2 * CountToExit);
                    continue;
                }

                foreach (var d in alpha)
                {
                    if (ObjWorkSheet.Range[d + i.ToString(), d + i.ToString()].Text.ToString().Trim(charsToTrim) != ObjWorkSheet2.Range[d + i.ToString(), d + i.ToString()].Text.ToString().Trim(charsToTrim)
                        /*&& (ObjWorkSheet.get_Range(d + i.ToString(), d + i.ToString()).Text.ToString().Trim(charsToTrim) != ""
                        || ObjWorkSheet2.get_Range(d + i.ToString(), d + i.ToString()).Text.ToString().Trim(charsToTrim) != "")*/)
                    {
                        IsDiff = true;
                        break;
                    }
                }
                if (IsDiff && !FindSomething)
                {
                    richTextBox1.Text = richTextBox1.Text + "\n\nПри сравнении файлов были найдены следующие отличия: ";
                }
                if (IsDiff)
                {
                    IsDiff = false;
                    FindSomething = true;
                    string OldS = "";
                    string NewS = "";
                    foreach (var g in alpha)
                    {
                        if (ObjWorkSheet.Range[g + i.ToString(), g + i.ToString()].Text.ToString().Trim(charsToTrim).Replace(" \n", "\n").Replace("\n ", "\n").Replace("\n", "\\") != ""
                            || ObjWorkSheet2.Range[g + i.ToString(), g + i.ToString()].Text.ToString().Trim(charsToTrim).Replace(" \n", "\n").Replace("\n ", "\n").Replace("\n", "\\") != "")
                        {
                            OldS = OldS + ObjWorkSheet2.Range[g + i.ToString(), g + i.ToString()].Text.ToString().Trim(charsToTrim).Replace(" \n", "\n").Replace("\n ", "\n").Replace("\n", "\\") + "; ";
                            NewS = NewS + ObjWorkSheet.Range[g + i.ToString(), g + i.ToString()].Text.ToString().Trim(charsToTrim).Replace(" \n", "\n").Replace("\n ", "\n").Replace("\n", "\\") + "; ";
                        }
                    }
                    if (NeedToSetGroup == 1)
                    {
                        NeedToSetGroup = 0;
                        richTextBox1.Text = richTextBox1.Text + Group + "\n";
                    }
                    richTextBox1.Text = richTextBox1.Text + "   Было: " + OldS.Replace("/", "\\").Replace("\\\\", "\\") + "\n" + "   Стало: " + NewS.Replace("/", "\\").Replace("\\\\", "\\") + "\n";
                }

            }
            Application.DoEvents();
            if (!FindSomething)
                richTextBox1.Text = "Отличий не найдено";
            else
                richTextBox1.Text = richTextBox1.Text + "\nСравнение окончено";
            Application.DoEvents();
        }

        public void btnConnect_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK || openFileDialog2.ShowDialog() != DialogResult.OK)
            {
                btnDiffer.Enabled = false;
                btnGetAll.Enabled = false;
            }
            else
            {
                btnDiffer.Enabled = true;
                btnGetAll.Enabled = true;
                txtOldFile.Text = openFileDialog2.FileName; //старый
                txtNewFile.Text = openFileDialog1.FileName; //новый
            }
            Application.DoEvents();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            ObjExcel.Quit();
            ObjExcel2.Quit();
            Close();
        }

         private void btnWikiStyle_Click(object sender, EventArgs e)
        {
            if (txtNewFile.Text == "")
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                    return;
                else
                    txtNewFile.Text = openFileDialog1.FileName; //новый

            tabControl1.SelectedTab = tabPage2;
            richTextBox2.Clear();

            var dt = GetDataFromXls(txtNewFile.Text);
            var toOut = "";
            var group = "";

            foreach (DataRow row in dt.Rows)
            {
                if (group != row["Group"].ToString())
                {
                    group = row["Group"].ToString();
                    if (string.IsNullOrEmpty(toOut))
                        toOut = toOut + "\n{{Hider|" + group + "\n{|\n|-";
                    else
                        toOut = toOut + "\n|}\n}}\n{{Hider|" + group + "\n{|\n|-";
                }

                toOut = toOut + "\n|-" + (chbDate.Checked ? "\n|" + row["Date"].ToString() : "")
                + (chbTime.Checked ? "\n|" + row["Time"].ToString() : "")
                + (chbSubj.Checked ? "\n|" + row["Subject"].ToString() : "")
                + (chbTeacher.Checked ? "\n|" + row["Teacher"].ToString() : "")
                + (chbClass.Checked ? "\n|" + row["Room"].ToString() : "")
                + (chbOther.Checked ? "\n|" + row["Note"].ToString() : "");
            }
            richTextBox2.Text = toOut + "\n|}\n}}";

            //return;
        }

        public DataTable GetDataFromXls(string fileName)
        {
            char[] charsToTrim = { '*', ' ', '_', '\n' };
            char[] charsToTrim2 = { '*', ' ', '_', '\n', '-', '.', '-', ':' };
            char[] columnsToSplit = { 'A', 'B', 'E', 'F' };
            Regex pattern = new Regex("[* _\n-.-:]");

            //int[] Indexes = { };
            string group = "";
            int countToExit = 0;
            string[] arrayOfChars = { };
            DataTable resultTable = new DataTable();
            resultTable.Columns.Add("Group", typeof(string));
            resultTable.Columns.Add("Date", typeof(string));
            resultTable.Columns.Add("Time", typeof(string));
            resultTable.Columns.Add("Subject", typeof(string));
            resultTable.Columns.Add("Teacher", typeof(string));
            resultTable.Columns.Add("Room", typeof(string));
            resultTable.Columns.Add("Note", typeof(string));
            resultTable.Columns.Add("SortColumn", typeof(string));
            //-----------------------------------------------------
            //var e;
            var objWorkSheet = new Excel.Worksheet();
            try
            {
                Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", false, false, 0, true, false, false);
                objWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[numEdt.Value];
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при попытке открыть файл\n{0}", e.ToString());
            }
            progressBar1.Value = 0;

            //var ResultCell = "";
            for (int i = 4; i < 120; i++)
            {
                progressBar1.PerformStep();
                var IsGroup = objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim) != ""
                               && objWorkSheet.Range["B" + i, "B" + i].Text.ToString().Trim(charsToTrim) == ""
                               && objWorkSheet.Range["C" + i, "C" + i].Text.ToString().Trim(charsToTrim) == ""
                               && group != objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim);
                if (IsGroup || countToExit > 9)
                {
                    //richTextBox2.Text = richTextBox2.Text + group;
                    group = "Группа " + objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim);
                    if (IsGroup)
                        continue;
                }
                if (objWorkSheet.Range["A" + i, "A" + i].Text.ToString() == "Число")
                    continue;
                if (countToExit > 9)
                {
                    break;
                }
                if (objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim) == ""
                     && objWorkSheet.Range["B" + i, "B" + i].Text.ToString().Trim(charsToTrim) == ""
                     && objWorkSheet.Range["C" + i, "C" + i].Text.ToString().Trim(charsToTrim) == "")
                {
                    countToExit++;
                    if (countToExit > 5)
                        progressBar1.Value = 180 + (2 * countToExit);
                    continue;
                }

                var CellA = objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim);
                var CellB = objWorkSheet.Range["B" + i, "B" + i].Text.ToString().Trim(charsToTrim);
                var CellE = objWorkSheet.Range["E" + i, "E" + i].Text.ToString().Trim(charsToTrim);
                var CellAArr = objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim).Split('\n');
                var CellBArr = objWorkSheet.Range["B" + i, "B" + i].Text.ToString().Trim(charsToTrim).Split('\n');
                var CellEArr = objWorkSheet.Range["E" + i, "E" + i].Text.ToString().Trim(charsToTrim).Split('\n');

                if (CellAArr.Length == 1)
                {
                    var row = resultTable.NewRow();
                    row["Group"] = group;
                    row["Date"] = objWorkSheet.Range["A" + i, "A" + i].Text.ToString().Trim(charsToTrim);
                    row["Time"] = objWorkSheet.Range["B" + i, "B" + i].Text.ToString().Trim(charsToTrim);
                    row["Subject"] = objWorkSheet.Range["C" + i, "C" + i].Text.ToString().Trim(charsToTrim);
                    row["Teacher"] = objWorkSheet.Range["D" + i, "D" + i].Text.ToString().Trim(charsToTrim);
                    row["Room"] = objWorkSheet.Range["E" + i, "E" + i].Text.ToString().Trim(charsToTrim);
                    row["Note"] = objWorkSheet.Range["F" + i, "F" + i].Text.ToString().Trim(charsToTrim);
                    var sort = row["Group"].ToString().Remove(0, 7).Trim(charsToTrim2) + '.' + row["Date"].ToString().Split('.')[1] + row["Date"].ToString().Split('.')[0] + "_" + row["Time"].ToString().Trim(charsToTrim2);
                    row["SortColumn"] = pattern.Replace(sort, "");

                    resultTable.Rows.Add(row);
                }
                else
                {
                    for (var j = 0; j < CellAArr.Length; j++)
                    {
                        var row = resultTable.NewRow();
                        row["Group"] = group;
                        row["Date"] = CellAArr[j];
                        row["Time"] = (CellBArr.Length == CellAArr.Length ? CellBArr[j] : CellB);
                        row["Subject"] = objWorkSheet.Range["C" + i, "C" + i].Text.ToString().Trim(charsToTrim);
                        row["Teacher"] = objWorkSheet.Range["D" + i, "D" + i].Text.ToString().Trim(charsToTrim);
                        row["Room"] = (CellEArr.Length == CellAArr.Length ? CellEArr[j] : CellE);
                        row["Note"] = objWorkSheet.Range["F" + i, "F" + i].Text.ToString().Trim(charsToTrim);
                        var sort = row["Group"].ToString().Remove(0, 7).Trim(charsToTrim2) + '.' + row["Date"].ToString().Split('.')[1] + row["Date"].ToString().Split('.')[0] + "_" + row["Time"].ToString().Trim(charsToTrim2);
                        row["SortColumn"] = pattern.Replace(sort, "");
                        resultTable.Rows.Add(row);
                    }
                }
            }
            //DataView dv = resultTable.DefaultView;
            resultTable.DefaultView.Sort = "SortColumn asc";
            DataTable sortedDT = resultTable.DefaultView.ToTable();
            return sortedDT;
        }
    }
}
