using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Diagnostics;


namespace Cursach
{
    public partial class tabPage5 : Form
    {
        Implementation im;
        public tabPage5()
        {

            im = new Implementation();
            InitializeComponent();

            nameshop.Text = "50-i";
            obl.Text = "Одесская";
            city.Text = "Одесса";
            streat.Text = "50";
            nameofgoods.Text = "50";
            code.Text = "3";
            quantity.Text = "3";
            price.Text = "3";
            number.Text = "3";

            List<string> oblast = new List<string>
            {
                "Одесская","Николаевская","Херсонская"
            };
            obl.Items.AddRange(oblast);
        }
        int m = 0, t;
        private void Button1_Click(object sender, EventArgs e)
        {
            t = 1;
            im.Nameshop = nameshop.Text;
            if (nameshop.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не ввели название магазина!!!");
            }
            if (im.Proverca1() == 1)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод названия магазина!!!");
            }

            im.Obl = obl.Text;
            if (obl.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не указали область!!!");
            }
            if (im.Proverca2() != 0)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод области!!!");
            }

            im.City = city.Text;
            if (city.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не указали город!!!");
            }
            if (im.Proverca3() != 0)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод названия города!!!");
            }

            im.Streat = streat.Text;
            if (streat.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не указали улицу!!!");
            }
            if (im.Proverca4() != 0)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод названия улицы!!!");
            }

            im.Number = Convert.ToDouble(number.Text == "" ? "0" : number.Text);
            if (number.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не указали номер здания!!!");
            }
            if ((im.Number = Convert.ToDouble(number.Text == "" ? "0" : number.Text)) < 0)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод номера дома!!!");
            }

            im.Code = code.Text;
            if (code.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не ввели код товара!!!");
            }
            if (im.Proverca6() != 0)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод кода товара!!!");
            }
           
            im.Nameofgoods = nameofgoods.Text;
            if (nameofgoods.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не ввели название товара!!!");
            }
            if (im.Proverca5() != 0)
            {
                m = 1;
                DialogResult result = MessageBox.Show("Некоректный ввод названия товара!!!");
            }

            im.Quantity = Convert.ToDouble(quantity.Text == "" ? "0" : quantity.Text);
            if (quantity.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не указали количество товара!!!");
            }           

            im.Price = Convert.ToDouble(price.Text == "" ? "0" : price.Text);
            if (price.Text == "")
            {
                m = 1;
                DialogResult result = MessageBox.Show("Вы не указали цену за одну единицу товара!!!");
            }           

            amount.Text = Convert.ToString(im.Amount());
        }
        private void Button2_Click(object sender, EventArgs e)
        {
           
            if (m == 0 && t == 1)
            {
                Tabl.Rows.Add(nameshop.Text, obl.Text, city.Text, streat.Text, number.Text, nameofgoods.Text, code.Text, quantity.Text, price.Text, amount.Text);
                spravochnik.Text = "Данные успешно добавлены!!!\n";
            }
            else
                spravochnik.Text = "\nДанные не были добавлены в табллицу. Отсуствует либо некоректно введины один или несколько элементов!!!" + "\n";
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            string str1 = Microsoft.VisualBasic.Interaction.InputBox("Введите название магагзина:");
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                Tabl.Rows[i].Visible = (Tabl.Rows[i].Cells[5].Value.ToString() == str1);
            }
        }
        double sum = 0;
        private void B2_Click(object sender, EventArgs e)
        {
            string str1 = Microsoft.VisualBasic.Interaction.InputBox("Введите название магагзина:");
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                Tabl.Rows[i].Visible = (Tabl.Rows[i].Cells[5].Value.ToString() == str1);
                sum += Convert.ToDouble(Tabl.Rows[i].Cells[9].Value.ToString() == "" ? "0" : Tabl.Rows[i].Cells[9].Value.ToString());
            }
            DialogResult result = MessageBox.Show("Объщая сумма по всем магазинам " + str1 + " = " + Convert.ToString(sum));
        }
        private void B3_Click(object sender, EventArgs e)
        {
            string str1 = Microsoft.VisualBasic.Interaction.InputBox("Введите название магагзина:");
            double Min = 0;
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                Tabl.Rows[i].Visible = (Tabl.Rows[i].Cells[5].Value.ToString() == str1);
                Min = Convert.ToDouble(Tabl.Rows[i].Cells[8].Value.ToString() == "" ? "0" : Tabl.Rows[i].Cells[8].Value.ToString());
                if ((Tabl.Rows[i].Cells[5].Value.ToString() == str1) && (Min != 0))
                {
                    break;
                }
            }
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                if (Min > Convert.ToDouble(Tabl.Rows[i].Cells[8].Value.ToString() == "" ? "0" : Tabl.Rows[i].Cells[8].Value.ToString()))
                {
                    Min = Convert.ToDouble(Tabl.Rows[i].Cells[8].Value.ToString() == "" ? "0" : Tabl.Rows[i].Cells[8].Value.ToString());
                }
            }
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                Tabl.Rows[i].Visible = (Min == Convert.ToDouble(Tabl.Rows[i].Cells[8].Value.ToString() == "" ? "0" : Tabl.Rows[i].Cells[8].Value.ToString()));
            }
        }
        private void B4_Click(object sender, EventArgs e)
        {
            string str1 = Microsoft.VisualBasic.Interaction.InputBox("Введите цену за единицу товара:");
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                if (str1 == Tabl.Rows[i].Cells[8].Value.ToString())
                {
                    Tabl.Rows.RemoveAt(i);
                }
            }
        }
        private void Button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < Tabl.RowCount - 1; ++i)
            {
                Tabl.Rows[i].Visible = true;
            }
        }
        private void Button8_Click(object sender, EventArgs e)
        {
            string str1 = Microsoft.VisualBasic.Interaction.InputBox("Введите пароль:");
            if (str1 == password.Text && password.Text != "")
            {
                Tabl.ReadOnly = false;
                b4.Enabled = true;
            } 
            else
                MessageBox.Show("Неверный пароль!!!");
        }
        private void Button9_Click(object sender, EventArgs e)
        {
            int pin;
            Random rand = new Random();
            pin = rand.Next(9999);
            password.Text = Convert.ToString(pin);
            spravochnik.Text += "Password was generated: " + password.Text + "\n";
        }
        void Check_input_digit(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.Equals('\b')) return;
            e.Handled = !char.IsDigit(e.KeyChar);
            if (!(Char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                    e.Handled = true;
            }
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)        
        {
            if (Tabl.RowCount != 0)
            {
                for (int i = Tabl.RowCount-2; i > -1; --i)
                {
                    Tabl.Rows.RemoveAt(i);
                }
            }
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLS)|*.XLS ";

            if (opf.ShowDialog() == DialogResult.OK)
            {
                DataTable tb = new DataTable();
                string filename = opf.FileName;
                Excel.Application ExcelApp = new Excel.Application();
                Excel._Workbook ExcelWorkBook;
                Excel.Worksheet ExcelWorkSheet;

              
                ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    "\t", false, false, 0, true, 1, 0);

                ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                for (int i = 1; i <= ExcelApp.Rows.Count; i++)
                {
                    if (ExcelApp.Cells[i, 1].Value != null)
                    {                       
                        Tabl.Rows.Add(ExcelApp.Cells[i, 1].Value, ExcelApp.Cells[i, 2].Value, ExcelApp.Cells[i, 3].Value,
                            ExcelApp.Cells[i, 4].Value, ExcelApp.Cells[i, 5].Value, ExcelApp.Cells[i, 6].Value, ExcelApp.Cells[i, 7].Value,
                            ExcelApp.Cells[i, 8].Value, ExcelApp.Cells[i, 9].Value, ExcelApp.Cells[i, 10].Value);
                    }
                    else
                    {
                        break;
                    }
                }

            }
        }
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;           

            for (int i = 1; i < Tabl.RowCount+1; i++)
            {
                for (int j = 1; j < Tabl.ColumnCount+1; j++)
                {
                    worksheet.Rows[i].Columns[j] = Tabl.Rows[i-1].Cells[j-1].Value;
                }
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel (*.xls)|*.xls";
            
            string path = null;
            saveDialog.ShowDialog();
            path = saveDialog.FileName;
            app.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            app.Quit();
            spravochnik.Text += "Файл суспешно сохранён в xls.";
            MessageBox.Show("Файл сохранён");            
        }
        private void HelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\\Users\\Будяну Андрей\\Desktop\\проги\\лабы\\Cursach\\Cursach\\reference.chm");
        }
    }
}
