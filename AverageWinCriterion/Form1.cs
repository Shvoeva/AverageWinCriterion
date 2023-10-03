using System.Data;

namespace AverageWinCriterion
{
    public partial class Form1 : Form
    {
        private string goal = "";

        private byte expertsNumber = 1;

        private byte controlOptionsNumber = 1;

        private byte situationsNumber = 1;

        private string efficiencyCriterion = "";

        private float budget = 100000;

        DataTable evaluationTabel;

        TabPage[] pages;

        DataGridView[] dataGridViews;

        DataTable[] individualTabels;

        Button[] buttonsClear;

        SplitContainer[] splitContainers;

        DataTable pivotTable;

        DataTable efficiencyTable;

        public Form1()
        {
            InitializeComponent();

            textBox1.Text = "Данная программа предназначена для оценивания альтернатив " +
                "на основе критерия среднего выигрыша (групповое оценивание).\r\n" +
                "\r\n\tИнструкция:\r\n" +
                "\t1) Укажите цель, достижение которой поможет решить выявленную проблему.\r\n" +
                "\t2) Введите количество экспертов, которые будут участвовать в оценивании " +
                "вероятности возникновениявозможных ситуаций.\r\n" +
                "\t3) Запишите количество вариантов управления (альтернативы), которые будут сравниваться между собой.\r\n" +
                "\t4) Добавьте количество возможных ситуаций, которые могут повлиять на выбор.\r\n" +
                "\t5) Напечатайте критерий эффективности, по которому будет проводиться оценка.\r\n" +
                "\t6) Укажите выделенный бюджет в рублях.\r\n" +
                "\t7) Нажмите \"ВВОД\"\r\n" +
                "\t8) Перейдите в раздел непосредственного оценивания.\r\n " +
                "\t9) Введите ситуации, экспертов, их компететность (цифры от 0 до 1, которые в сумме дают 1) и " +
                "оценку ситуаций.\r\n" +
                "\t10) Нажмите кнопку \"ВВОД\", после чего покажутся расчеты обобщенной оценки " +
                "и коэффициента вариации.\r\n" +
                "\t11) Перейдите в раздел расчета критерия среднего выигрыша для каждого эксперта.\r\n" +
                "\t12) Введите варианты управления и их критерии эффективности при разных ситуациях для каждого эксперта.\r\n" +
                "\t13) Перейдите на вкладку сводной таблицы и нажмите кнопку \"ВВОД\", после чего покажутся расчеты\r\n" +
                "эффективности по критериям и оптимальное значение и сводная таблица.\r\n" +
                "\t14) В текущей вкладке выведется сводная таблица, во вкладке «Результат»: таблица эффективности по критерию\r\n" +
                "среднего выигрыша, оптимальное значение.\r\n" +
                "\t15) Во вкладках «Коэффициент вариации» и «Лингвистическое значение коэффициента» отображены соответствующие значения.";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            goal = "";
            expertsNumber = 1;
            controlOptionsNumber = 1;
            situationsNumber = 1;
            efficiencyCriterion = "";
            budget = 100000;

            textBox2.Text = "";
            numericUpDown1.Value = 1;
            numericUpDown2.Value = 1;
            numericUpDown3.Value = 1;
            textBox3.Text = "";

            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;

            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;
            dataGridView5.DataSource = null;
            dataGridView6.DataSource = null;
            while (tabControl1.TabCount > 6)
            {
                tabControl1.TabPages.RemoveAt(6);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            goal = textBox2.Text;
            expertsNumber = Decimal.ToByte(numericUpDown1.Value);
            controlOptionsNumber = Decimal.ToByte(numericUpDown2.Value);
            situationsNumber = Decimal.ToByte(numericUpDown3.Value);
            efficiencyCriterion = textBox3.Text;
            budget = (float)numericUpDown6.Value;

            button3.Enabled = true;
            button4.Enabled = true;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;
            dataGridView5.DataSource = null;
            dataGridView6.DataSource = null;
            while (tabControl1.TabCount > 6)
            {
                tabControl1.TabPages.RemoveAt(6);
            }

            evaluationTabel = new DataTable();

            DataColumn[] column = new DataColumn[situationsNumber + 2];
            column[0] = new DataColumn("Эксперты", Type.GetType("System.String"));
            column[1] = new DataColumn("Компетентность", Type.GetType("System.Single"));
            for (int i = 2; i < situationsNumber + 2; i++)
            {
                column[i] = new DataColumn("Ситуация " + (i - 1).ToString(), Type.GetType("System.Single"));
            }
            for (int i = 0; i < situationsNumber + 2; i++)
            {
                evaluationTabel.Columns.Add(column[i]);
            }

            for (int i = 0; i < expertsNumber; i++)
            {
                evaluationTabel.Rows.Add(new object[] { ("Эксперт " + (i + 1).ToString()).ToString() });
            }
            evaluationTabel.Rows.Add(new object[] { "Обобщенная оценка" });
            evaluationTabel.Rows.Add(new object[] { "Коэффициент вариации" });

            dataGridView1.DataSource = evaluationTabel;
            dataGridView1.Rows[expertsNumber].ReadOnly = true;
            dataGridView1.Rows[expertsNumber + 1].ReadOnly = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button2_Click(sender, e);

            button5.Enabled = false;

            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;
            dataGridView5.DataSource = null;
            dataGridView6.DataSource = null;
            while (tabControl1.TabCount > 6)
            {
                tabControl1.TabPages.RemoveAt(6);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;

            for (int i = 0; i < situationsNumber; i++)
            {
                evaluationTabel.Rows[expertsNumber][i + 2] = 0;
                for (int j = 0; j < expertsNumber; j++)
                {
                    evaluationTabel.Rows[expertsNumber][i + 2] =
                        (float)evaluationTabel.Rows[expertsNumber][i + 2] + 
                        (float)evaluationTabel.Rows[j][1] * 
                        (float)evaluationTabel.Rows[j][i + 2];
                }
            }

            var coefficientTabel = new DataTable();

            DataColumn[] column = new DataColumn[situationsNumber + 2];
            column[0] = new DataColumn("Эксперты", Type.GetType("System.String"));
            column[1] = new DataColumn("Компетентность", Type.GetType("System.String"));
            for (int i = 2; i < situationsNumber + 2; i++)
            {
                column[i] = new DataColumn("Ситуация " + (i - 1).ToString(), Type.GetType("System.String"));
            }
            for (int i = 0; i < situationsNumber + 2; i++)
            {
                coefficientTabel.Columns.Add(column[i]);
            }
            coefficientTabel.Rows.Add(new object[] { "Коэффициент вариации, лингвистичесеок значение" });

            for (int i = 0; i < situationsNumber; i++)
            {
                evaluationTabel.Rows[expertsNumber + 1][i + 2] = 0;
                for (int j = 0; j < expertsNumber; j++)
                {
                    evaluationTabel.Rows[expertsNumber + 1][i + 2] =
                        (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] +
                        (float)evaluationTabel.Rows[j][1] *
                        Math.Pow((float)evaluationTabel.Rows[j][i + 2] - 
                        (float)evaluationTabel.Rows[expertsNumber][i + 2], 2);
                }
                evaluationTabel.Rows[expertsNumber + 1][i + 2] = 
                    Math.Sqrt((float)evaluationTabel.Rows[expertsNumber + 1][i + 2]);
                evaluationTabel.Rows[expertsNumber + 1][i + 2] = 
                    (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] * 100;
                evaluationTabel.Rows[expertsNumber + 1][i + 2] =
                    (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] /
                    (float)evaluationTabel.Rows[expertsNumber][i + 2];

                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] <= 11)
                {
                    coefficientTabel.Rows[0][i + 2] = "Очень высокая";
                }
                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] > 11 
                    && (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] <= 22)
                {
                    coefficientTabel.Rows[0][i + 2] = "Высокая";
                }
                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] > 22 
                    && (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] <= 33)
                {
                    coefficientTabel.Rows[0][i + 2] = "Умеренная";
                }
                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] > 33)
                {
                    coefficientTabel.Rows[0][i + 2] = "Недостаточная (слабая)";
                }
            }

            dataGridView4.DataSource = coefficientTabel;

            while (tabControl1.TabCount > 6)
            {
                tabControl1.TabPages.RemoveAt(6);
            }

            pages = new TabPage[expertsNumber];
            dataGridViews = new DataGridView[expertsNumber];
            individualTabels = new DataTable[expertsNumber];
            buttonsClear = new Button[expertsNumber];
            splitContainers = new SplitContainer[expertsNumber];
            for (int i = 0; i < expertsNumber; i++)
            {
                dataGridViews[i] = new DataGridView();
                dataGridViews[i].Dock = DockStyle.Fill;
                dataGridViews[i].ReadOnly = false;
                dataGridViews[i].AllowUserToAddRows = false;
                dataGridViews[i].AllowUserToDeleteRows = false;

                buttonsClear[i] = new Button();
                buttonsClear[i].Dock = DockStyle.Fill;
                buttonsClear[i].Text = "ОЧИСТИТЬ";
                buttonsClear[i].Font = new Font(buttonsClear[i].Font.Name,
                    button1.Font.Size, button1.Font.Style);
                buttonsClear[i].MouseClick += OnMouseClick;

                splitContainers[i] = new SplitContainer();
                splitContainers[i].Orientation = 0;
                splitContainers[i].Dock = DockStyle.Fill;
                splitContainers[i].Panel2MinSize = 10;
                splitContainers[i].SplitterDistance = 1000;
                splitContainers[i].Panel1.Controls.Add(dataGridViews[i]);
                splitContainers[i].Panel2.Controls.Add(buttonsClear[i]);

                pages[i] = new TabPage();
                pages[i].Text = (string)evaluationTabel.Rows[i][0];
                tabControl1.TabPages.Add(pages[i]);
                pages[i].Controls.Add(dataGridViews[i]);

                individualTabels[i] = new DataTable();

                DataColumn[] indivTabelsColumn = new DataColumn[situationsNumber + 1];
                indivTabelsColumn[0] = new DataColumn("Варианты управления", Type.GetType("System.String"));
                individualTabels[i].Columns.Add(indivTabelsColumn[0]);
                for (int j = 1; j < situationsNumber + 1; j++)
                {
                    indivTabelsColumn[j] = new DataColumn(evaluationTabel.Columns[j + 1].ColumnName, Type.GetType("System.Single"));
                    individualTabels[i].Columns.Add(indivTabelsColumn[j]);
                }

                for (int j = 0; j < controlOptionsNumber; j++)
                {
                    individualTabels[i].Rows.Add(new object[] { ("Вариант управления " + (j + 1).ToString()).ToString() });
                }

                dataGridViews[i].DataSource = individualTabels[i];
            }
        }

        private void OnMouseClick(object sender, EventArgs e)
        {
            var button = (Button)sender;
            var index = buttonsClear.ToList().IndexOf(button);

            individualTabels[index] = new DataTable();

            DataColumn[] indivTabelsColumn = new DataColumn[situationsNumber + 1];
            indivTabelsColumn[0] = new DataColumn("Варианты управления", Type.GetType("System.String"));
            individualTabels[index].Columns.Add(indivTabelsColumn[0]);
            for (int j = 1; j < situationsNumber + 1; j++)
            {
                indivTabelsColumn[j] = new DataColumn(evaluationTabel.Columns[j + 1].ColumnName, Type.GetType("System.Single"));
                individualTabels[index].Columns.Add(indivTabelsColumn[j]);
            }

            for (int j = 0; j < situationsNumber; j++)
            {
                individualTabels[index].Rows.Add(new object[] { ("Вариант управления " + (j + 1).ToString()).ToString() });
            }

            dataGridViews[index].DataSource = individualTabels[index];

            dataGridView3.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView5.DataSource = null;
            dataGridView6.DataSource = null;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            pivotTable = new DataTable();
            DataTable consistencyTabel = new DataTable();
            DataTable linguisticСonsistencyTabel = new DataTable();

            DataColumn[] column = new DataColumn[situationsNumber + 1];
            column[0] = new DataColumn("Варианты управления", Type.GetType("System.String"));
            pivotTable.Columns.Add(column[0]);
            for (int i = 1; i < situationsNumber + 1; i++)
            {
                column[i] = new DataColumn(evaluationTabel.Columns[i + 1].ColumnName, Type.GetType("System.Single"));
                pivotTable.Columns.Add(column[i]);
            }

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                pivotTable.Rows.Add(new object[] { individualTabels[0].Rows[i][0] });
            }

            DataColumn[] consistencyСolumn = new DataColumn[situationsNumber + 1];
            consistencyСolumn[0] = new DataColumn("Варианты управления", Type.GetType("System.String"));
            consistencyTabel.Columns.Add(consistencyСolumn[0]);
            for (int i = 1; i < situationsNumber + 1; i++)
            {
                consistencyСolumn[i] = new DataColumn(evaluationTabel.Columns[i + 1].ColumnName, Type.GetType("System.Single"));
                consistencyTabel.Columns.Add(consistencyСolumn[i]);
            }

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                consistencyTabel.Rows.Add(new object[] { individualTabels[0].Rows[i][0] });
            }

            DataColumn[] linguisticColumn = new DataColumn[situationsNumber + 1];
            linguisticColumn[0] = new DataColumn("Варианты управления", Type.GetType("System.String"));
            linguisticСonsistencyTabel.Columns.Add(linguisticColumn[0]);
            for (int i = 1; i < situationsNumber + 1; i++)
            {
                linguisticColumn[i] = new DataColumn(evaluationTabel.Columns[i + 1].ColumnName, Type.GetType("System.String"));
                linguisticСonsistencyTabel.Columns.Add(linguisticColumn[i]);
            }

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                linguisticСonsistencyTabel.Rows.Add(new object[] { individualTabels[0].Rows[i][0] });
            }

            dataGridView2.DataSource = pivotTable;
            dataGridView5.DataSource = consistencyTabel;
            dataGridView6.DataSource = linguisticСonsistencyTabel;

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                for (int j = 1; j < situationsNumber + 1; j++)
                {
                    pivotTable.Rows[i][j] = 0;
                    for (int k = 0; k < expertsNumber; k++)
                    {
                        pivotTable.Rows[i][j] = (float)pivotTable.Rows[i][j] + (float)(individualTabels[k].Rows[i][j]) * (float)evaluationTabel.Rows[k][1];
                    }
                    pivotTable.Rows[i][j] = budget - (float)pivotTable.Rows[i][j];
                }
            }

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                for (int j = 1; j < situationsNumber + 1; j++)
                {
                    consistencyTabel.Rows[i][j] = 0;
                    for (int k = 0; k < expertsNumber; k++)
                    {
                        consistencyTabel.Rows[i][j] =
                            (float)consistencyTabel.Rows[i][j] +
                            (float)evaluationTabel.Rows[k][1] *
                            Math.Pow(budget - (float)individualTabels[k].Rows[i][j] -
                            (float)pivotTable.Rows[i][j], 2);
                    }

                    consistencyTabel.Rows[i][j] = 
                        Math.Sqrt((float)consistencyTabel.Rows[i][j]);
                    consistencyTabel.Rows[i][j] =
                        (float)consistencyTabel.Rows[i][j] * 100;
                    consistencyTabel.Rows[i][j] =
                        (float)consistencyTabel.Rows[i][j] /
                        (float)pivotTable.Rows[i][j];

                    if ((float)consistencyTabel.Rows[i][j] <= 11)
                    {
                        linguisticСonsistencyTabel.Rows[i][j] = "Очень высокая";
                    }
                    if ((float)consistencyTabel.Rows[i][j] > 11
                        && (float)consistencyTabel.Rows[i][j] <= 22)
                    {
                        linguisticСonsistencyTabel.Rows[i][j] = "Высокая";
                    }
                    if ((float)consistencyTabel.Rows[i][j] > 22
                        && (float)consistencyTabel.Rows[i][j] <= 33)
                    {
                        linguisticСonsistencyTabel.Rows[i][j] = "Умеренная";
                    }
                    if ((float)consistencyTabel.Rows[i][j] > 33)
                    {
                        linguisticСonsistencyTabel.Rows[i][j] = "Недостаточная (слабая)";
                    }
                }
            }

            efficiencyTable = new DataTable();

            DataColumn[] efficiencyColumn = new DataColumn[2];
            efficiencyColumn[0] = new DataColumn("Варианты управления", Type.GetType("System.String"));
            efficiencyTable.Columns.Add(efficiencyColumn[0]);
            efficiencyColumn[1] = new DataColumn("Эффективность по критериям", Type.GetType("System.Single"));
            efficiencyTable.Columns.Add(efficiencyColumn[1]);

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                float sum = 0;
                for (int j = 1; j < situationsNumber + 1; j++)
                {
                    sum += (float)pivotTable.Rows[i][j] * (float)evaluationTabel.Rows[expertsNumber][j + 1];
                }
                efficiencyTable.Rows.Add(new object[] { individualTabels[0].Rows[i][0], sum});
            }

            dataGridView3.DataSource = efficiencyTable;

            float maxValue = (float)efficiencyTable.Rows[0][1];
            label7.Text = (string)efficiencyTable.Rows[0][0];
            for (int i = 1; i < controlOptionsNumber; i++)
            {
                if ((float)efficiencyTable.Rows[i][1] > maxValue)
                {
                    maxValue = (float)efficiencyTable.Rows[i][1];
                    label7.Text = (string)efficiencyTable.Rows[i][0];
                }
            }
        }
    }
}