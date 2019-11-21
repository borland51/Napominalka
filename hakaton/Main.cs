using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;

namespace hakaton
{
    public partial class Main : Form
    {
        List<int> GetDays = new List<int>();
        List<DataGridView> sGrids = new List<DataGridView>();
        int SelectedGrid = -1;
        static public bool isLoaded = false;
        string sMess = null;

        const int SET_INTERVAL = 60 * 60 * 1000;

        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            timer1.Interval = SET_INTERVAL;

            if (!TrshConfig.GetConfig())
                MessageBox.Show("Это первый запуск программы, пожалуйста настройте её.");
            else if (TrshConfig.SettType == 0)
                MessageBox.Show("В настройках программы указан неправильный путь к файлу/директории, пожалуйста перенастройте.");
            else
            {
                LoadForm();
                return;
            }

            button3_Click(sender, e);
            LoadForm();
        }

        public void LoadForm(bool msg = true)
        {
            timer1.Enabled = false;

            SelectedGrid = -1;
            SetGrid(dataGridView1);

            if(sGrids.Count > 0)
            {
                for (int i = 1; i < sGrids.Count; i++)
                {
                    this.Controls.Remove(sGrids[i]);
                    sGrids[i].Dispose();
                }

                sGrids.Clear();
            }
            GetDays.Clear();

            sGrids.Add(dataGridView1);
            for (int i = 0; i < TrshConfig.SettDays.Count; i++)
            {
                if(i > 0)
                    CreateGrid();

                GetDays.Add(TrshConfig.SettDays[i]);
            }

            button1.Visible = false;
            button2.Visible = false;

            LoadExcel(msg);
            isLoaded = true;
            timer1.Enabled = true;
        }

        void LoadExcel(bool msg = true)
        {
            sMess = null;
            switch (TrshConfig.GetType(TrshConfig.SettFile))
            {
                case 0:
                    MessageBox.Show("В настройках программы указан неправильный путь к файлу/директории, пожалуйста перенастройте.");
                    return;
                case 1:
                    GetExcelInDataGrid(TrshConfig.SettFile);
                    break;
                case 2:
                    string[] dirs = Directory.GetFiles(TrshConfig.SettFile, "*.xls*");
                    foreach (string file in dirs)
                    {
                        var f = new FileInfo(file);
                        if(f.Name[0] != '~')
                            GetExcelInDataGrid(file);
                    }
                    break;
            }

            int i = -1;
            while(++i < sGrids.Count)
            {
                if (sGrids[i].Rows.Count == 0)
                {
                    if (sGrids[i] != dataGridView1)
                        sGrids[i].Dispose();
                    else
                        sGrids[i].Visible = false;

                    GetDays.RemoveAt(i);
                    sGrids.RemoveAt(i);

                    i--;
                    continue;
                }
                else
                {
                    this.Controls.Add(sGrids[i]);
                    sMess += "Через " + GetDays[i].ToString() + " " + get_wordend(GetDays[i], "день", "дня", "дней") +
                            (sGrids[i].RowCount == 1 ? " истечет " : " истекут ") + sGrids[i].RowCount.ToString() + " " + get_wordend(sGrids[i].RowCount, "договор", "договора", "договоров") + "\n";
                }
            }

            switch (sGrids.Count)
            {
                case 0:
                    MessageBox.Show("Никаких данных не найдено!");
                    return;
                case 1: break;
                default:
                 button1.Visible = true;
                 button2.Visible = true;
                 break;
            }

            SelectedGrid = 0;
            UpdateLabel();

            if(msg)
                MessageBox.Show(sMess);
        }

        void GetExcelInDataGrid(string fileName)
        {
            Excel.Application ObjWorkExcel = null;
            Excel.Workbook ObjWorkBook = null;
            double percent = 0, addPerc = 0;

            Loading frm = null;

            if (WindowState != FormWindowState.Minimized)
            {
                frm = new Loading();
                frm.Show();
            }

            try
            {
                ObjWorkExcel = new Excel.Application(); //открыть эксель
                ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //лист

                dynamic getSort = ObjWorkSheet.UsedRange;
                getSort.Sort(getSort.Columns[4], Excel.XlSortOrder.xlAscending);

                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                addPerc = 100.0 / lastCell.Row;

                DateTime nDate, date = DateTime.Today;
                string[] list = new string [lastCell.Column];

                for (int i = 0; i < lastCell.Row; i++)
                {
                    percent += addPerc;
                    frm.SetProgress(Math.Round(percent));

                    for (int j = 0; j < lastCell.Column; j++)
                        list[j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();

                    if (!DateTime.TryParse(list[3], out nDate))
                        continue;

                    double iDays = (nDate - date).TotalDays;
                    for (int k = 0; k < TrshConfig.SettDays.Count; k++)
                    {
                        if (iDays != TrshConfig.SettDays[k])
                            continue;

                        sGrids[k].Rows.Add(list);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (ObjWorkBook != null)
                    ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя

                if (ObjWorkExcel != null)
                    ObjWorkExcel.Quit(); // выйти из экселя

                if(frm != null)
                    frm.Close();
            }
        }

        DataGridView CreateGrid()
        {
            DataGridView grid = new DataGridView();
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.RowTemplate.ReadOnly = false;
            grid.Anchor = dataGridView1.Anchor;
            grid.Size = dataGridView1.Size;
            grid.Font = dataGridView1.Font;
            grid.BackgroundColor = dataGridView1.BackgroundColor;
            grid.ForeColor = dataGridView1.ForeColor;
            grid.Location = dataGridView1.Location;
            grid.Name = "DataGridView" + sGrids.Count;
            
            SetGrid(grid);

            sGrids.Add(grid);
            return grid;
        }

        void SetGrid(DataGridView grid)
        {
            grid.Rows.Clear();

            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;

            grid.ColumnCount = 4;
            grid.ColumnHeadersVisible = true;

            grid.Columns[0].Name = "№";
            grid.Columns[1].Name = "Наименование органазации";
            grid.Columns[2].Name = "№, дата заключения договора";
            grid.Columns[3].Name = "Срок окончания действия договора";

            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9.75F, FontStyle.Bold);

            for (int i = 0; i < 4; i++)
               grid.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Settings frm = new Settings(this);
            frm.ShowDialog();
            frm.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (sGrids.Count == 0 || SelectedGrid == -1)
                return;
            
            SelectedGrid = ++SelectedGrid >= sGrids.Count ? 0 : SelectedGrid;
            sGrids[((SelectedGrid > 0) ? SelectedGrid : sGrids.Count) - 1].Visible = false;
            sGrids[SelectedGrid].Visible = true;

            UpdateLabel();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (sGrids.Count == 0 || SelectedGrid == -1)
                return;

            SelectedGrid = --SelectedGrid < 0 ? 0 : SelectedGrid;
            sGrids[SelectedGrid + 1].Visible = false;
            sGrids[SelectedGrid].Visible = true;

            UpdateLabel();
        }

        void UpdateLabel()
        {
            label1.Text = "Срок действия " + get_wordend(sGrids[SelectedGrid].RowCount, "договора", "договоров", "договоров") + " истечёт через " + GetDays[SelectedGrid].ToString() + " " + get_wordend(GetDays[SelectedGrid], "день", "дня", "дней") + ".";
            UpdateLabe2();
        }

        void UpdateLabe2()
        {
            label2.Text = "Количество: " + sGrids[SelectedGrid].RowCount.ToString();
        }

        string get_wordend(int day, string first, string second, string third)
        {
            switch (day % 10)
            {
                case 1:
                    return first;
                case 2: 
                case 3: 
                case 4:
                    return second;
                default:
                    return third;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application ex = null;
            try
            {
                if(SelectedGrid == -1)
                    return;
                        
                DataGridView grid = sGrids[SelectedGrid];

                saveFileDialog1.FileName = GetDays[SelectedGrid].ToString() + "days";
                saveFileDialog1.DefaultExt = ".xlsx";
                saveFileDialog1.Filter = "Книга Excel (*.xlsx)|*.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;

                if (grid.Rows.Count == 0)
                {
                    MessageBox.Show("Нет данных для вывода в excel!", "Ошибка!");
                    return;
                }

                //Объявляем приложение
                ex = new Excel.Application();
                //Количество листов в рабочей книге
                ex.SheetsInNewWorkbook = 1;
                //Добавить рабочую книгу
                Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
                //Отключить отображение окон с сообщениями
                ex.DisplayAlerts = false;
                //Получаем первый лист документа (счет начинается с 1)
                Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
                //Название листа (вкладки снизу)
                sheet.Name = "За " + GetDays[SelectedGrid].ToString() + get_wordend(GetDays[SelectedGrid], "день", "дня", "дней");

                string[] arr1 = { "A", "D", "G", "I" }, arr2 = { "C", "F", "H", "L" };
                goto_range(1, 1, sheet, arr1, arr2);

                sheet.Cells[1, 1] = "№";
                sheet.Cells[1, 4] = "Название организации";
                sheet.Cells[1, 7] = "№, дата заключения договора";
                sheet.Cells[1, 9] = "Срок окончания действия договора ";

                int row = 2;
                for (int i = 0; i < grid.Rows.Count; i++)
                {
                    goto_range(row, row, sheet, arr1, arr2);

                    sheet.Cells[row, 1] = i+1;
                    sheet.Cells[row, 4] = grid.Rows[i].Cells[1].Value.ToString();
                    sheet.Cells[row, 7] = grid.Rows[i].Cells[2].Value.ToString();
                    sheet.Cells[row, 9] = grid.Rows[i].Cells[3].Value.ToString();

                    row++;
                }

                /*SetLine(row, row, sheet, arr1[0], arr2[arr2.Length - 1]);
                goto_range(++row, row, sheet, arr1, arr2);
                sheet.Cells[row, 1] = $"Всего: {grid.Rows.Count}";*/

                border_set(sheet, arr1[0], arr2[arr2.Length - 1], row-1);

                ex.Application.ActiveWorkbook.SaveAs(saveFileDialog1.FileName, Type.Missing,
                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MessageBox.Show("Экспорт завершен успешно!");
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Произошла ошибка!");
            }
            finally
            {
                if(ex != null)
                    ex.Quit();
            }
        }

        void SetLine(int start, int end, Excel.Worksheet sheet, string st, string et)
        {
            Excel.Range rangeCell;
            rangeCell = sheet.get_Range($"{st}{start}", $"{et}{end}");
            rangeCell.Merge(Type.Missing);
        }

        void goto_range(int start, int end, Excel.Worksheet sheet, string[] st, string[] et)
        {
            for (int i = 0; i < st.Length; i++)
            {
                Excel.Range rangeCell = sheet.get_Range($"{st[i]}{start}", $"{et[i]}{end}");
                rangeCell.Merge(Type.Missing);
            }
        }

        void border_set(Excel.Worksheet sheet, string st, string et, int row)
        {
            Excel.Range range;
            range = sheet.get_Range($"{st}1", $"{et}{row}");

            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            notifyIcon1.Visible = false;
            this.ShowInTaskbar = true;
            WindowState = FormWindowState.Normal;
        }

        private void Main_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                notifyIcon1.Visible = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            LoadForm(false);

            if (sMess == null)
                return;

            if (WindowState == FormWindowState.Minimized)
                notifyIcon1.ShowBalloonTip(10000, "Обновлено!","Приложение получило новые данные по договорам, зайдите в приложение...",ToolTipIcon.Info);

            MessageBox.Show(sMess);
            
        }

        string GetFileHash(string filepath)
        {
            var md5 = MD5.Create();
            var stream = File.OpenRead(filepath);
            string hash = Encoding.Default.GetString(md5.ComputeHash(stream));
            stream.Close();

            return hash;
        }
    }
}
