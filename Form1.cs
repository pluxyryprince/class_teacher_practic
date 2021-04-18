using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;


namespace классный_руководитель
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        void exportData(DataGridView dataGrid)
        {
            ExcelObj.Application ExcelApp = new ExcelObj.Application();
            ExcelObj.Workbook ExcelWorkBook;
            ExcelObj.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(Missing.Value);
            ExcelWorkSheet = ExcelWorkBook.Worksheets.get_Item(1) as ExcelObj.Worksheet;

            for (int i = 0; i < dataGrid.Rows.Count; i++)
            {
                for (int j = 0; j < dataGrid.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGrid.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage10_Click(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDataSet3.teacher". При необходимости она может быть перемещена или удалена.
            this.teacherTableAdapter1.Fill(this.class_teacherDataSet3.teacher);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDataSet2._class". При необходимости она может быть перемещена или удалена.
            this.classTableAdapter1.Fill(this.class_teacherDataSet2._class);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDataSet1.journal". При необходимости она может быть перемещена или удалена.
            this.journalTableAdapter1.Fill(this.class_teacherDataSet1.journal);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDataSet.ratings". При необходимости она может быть перемещена или удалена.
            this.ratingsTableAdapter1.Fill(this.class_teacherDataSet.ratings);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.benefits". При необходимости она может быть перемещена или удалена.
            this.benefitsTableAdapter.Fill(this.class_teacherDatabase.benefits);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.events_and_class_hours". При необходимости она может быть перемещена или удалена.
            this.events_and_class_hoursTableAdapter.Fill(this.class_teacherDatabase.events_and_class_hours);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.parents_conferences". При необходимости она может быть перемещена или удалена.
            this.parents_conferencesTableAdapter.Fill(this.class_teacherDatabase.parents_conferences);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.parent". При необходимости она может быть перемещена или удалена.
            this.parentTableAdapter.Fill(this.class_teacherDatabase.parent);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.subjects". При необходимости она может быть перемещена или удалена.
            this.subjectsTableAdapter.Fill(this.class_teacherDatabase.subjects);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.teacher". При необходимости она может быть перемещена или удалена.
            this.teacherTableAdapter.Fill(this.class_teacherDatabase.teacher);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase._class". При необходимости она может быть перемещена или удалена.
            this.classTableAdapter.Fill(this.class_teacherDatabase._class);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.journal". При необходимости она может быть перемещена или удалена.
            this.journalTableAdapter.Fill(this.class_teacherDatabase.journal);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.ratings". При необходимости она может быть перемещена или удалена.
            this.ratingsTableAdapter.Fill(this.class_teacherDatabase.ratings);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "class_teacherDatabase.student". При необходимости она может быть перемещена или удалена.
            this.studentTableAdapter.Fill(this.class_teacherDatabase.student);

        }

        private void button1_Click(object sender, EventArgs e)//обновить данные в таблице студенты
        {
            try
            {
                studentTableAdapter.Update(class_teacherDatabase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }
       
        private void button3_Click(object sender, EventArgs e)//удалить данные из таблицы студенты
        {
            try
            {
                int delet = dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)//импорт данных в excel из таблицы студенты
        {
            try
            {
                exportData(dataGridView1);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            } 
        }

        private void button6_Click(object sender, EventArgs e)//обновить данные в таблице оценки
        {
            try
            {
                ratingsTableAdapter1.Update(class_teacherDataSet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)//удалить данные из таблицы оценки
        {
            try
            {
                int delet = dataGridView2.SelectedCells[0].RowIndex;
                dataGridView2.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button5_Click(object sender, EventArgs e)//экспорт данных их таблицы оценки
        {
            try
            {
                exportData(dataGridView2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button9_Click(object sender, EventArgs e)//обновить данные в таблице журнал
        {
            try
            {
                journalTableAdapter1.Update(class_teacherDataSet1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button7_Click(object sender, EventArgs e)//удалить данные из таблицы журнал
        {
            try
            {
                int delet = dataGridView3.SelectedCells[0].RowIndex;
                dataGridView3.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button8_Click(object sender, EventArgs e)//экспорт данных из таблицы журнал
        {
            try
            {
                exportData(dataGridView3);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button12_Click(object sender, EventArgs e)//обновление данных в таблице группы
        {
            try
            {
                classTableAdapter1.Update(class_teacherDataSet2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button10_Click(object sender, EventArgs e)//удаление данных их таблицы группы
        {
            try
            {
                int delet = dataGridView4.SelectedCells[0].RowIndex;
                dataGridView4.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button11_Click(object sender, EventArgs e)//экспорт данных из таблицы группы
        {
            try
            {
                exportData(dataGridView4);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button15_Click(object sender, EventArgs e)//обновление данных в таблице преподаватели
        {
            try
            {
                teacherTableAdapter1.Update(class_teacherDataSet3);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button13_Click(object sender, EventArgs e)//удаление данных из таблицы преподаватели
        {
            try
            {
                int delet = dataGridView5.SelectedCells[0].RowIndex;
                dataGridView5.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button14_Click(object sender, EventArgs e)//импорт данных их таблицы преподаватели
        {
            try
            {
                exportData(dataGridView5);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button18_Click(object sender, EventArgs e)//обновление данных в таблице предметы
        {
            try
            {
                subjectsTableAdapter.Update(class_teacherDatabase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button16_Click(object sender, EventArgs e)//удаление данных их таблицы предметы
        {
            try
            {
                int delet = dataGridView6.SelectedCells[0].RowIndex;
                dataGridView6.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button17_Click(object sender, EventArgs e)//импорт данных из таблицы предметы
        {
            try
            {
                exportData(dataGridView6);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button21_Click(object sender, EventArgs e)//обновление данных в таблице родители
        {
            try
            {
                parentTableAdapter.Update(class_teacherDatabase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button19_Click(object sender, EventArgs e)//удаление данных из таблицы родители
        {
            try
            {
                int delet = dataGridView7.SelectedCells[0].RowIndex;
                dataGridView7.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button20_Click(object sender, EventArgs e)//импорт данных из таблицы родители
        {
            try
            {
                exportData(dataGridView7);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button24_Click(object sender, EventArgs e)//обновление данных в таблице родительские собрания
        {
            try
            {
                parents_conferencesTableAdapter.Update(class_teacherDatabase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button22_Click(object sender, EventArgs e)//удаление данных из таблицы родительские собрания
        {
            try
            {
                int delet = dataGridView8.SelectedCells[0].RowIndex;
                dataGridView8.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button23_Click(object sender, EventArgs e)//экспорт данных из таблицы родительские собрания
        {
            try
            {
                exportData(dataGridView8);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button27_Click(object sender, EventArgs e)//обновление данных в таблице план проведения мероприятий и классных часов
        {
            try
            {
                events_and_class_hoursTableAdapter.Update(class_teacherDatabase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button25_Click(object sender, EventArgs e)//удаление данных из таблицы план проведения мероприятий и классных часов
        {
            try
            {
                int delet = dataGridView9.SelectedCells[0].RowIndex;
                dataGridView9.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button26_Click(object sender, EventArgs e)//экспорт данных из таблицы план проведения мероприятий и классных часов
        {
            try
            {
                exportData(dataGridView9);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button30_Click(object sender, EventArgs e)//обновить данные в таблице льготы
        {
            try
            {
                benefitsTableAdapter.Update(class_teacherDatabase);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }
        private void button28_Click(object sender, EventArgs e)//удаление данных в таблице льготы
        {
            try
            {
                int delet = dataGridView10.SelectedCells[0].RowIndex;
                dataGridView10.Rows.RemoveAt(delet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void button29_Click(object sender, EventArgs e)//экспорт данных из таблицы льготы
        {
            try
            {
                exportData(dataGridView10);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }
    }
}
