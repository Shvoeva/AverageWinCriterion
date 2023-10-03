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

            textBox1.Text = "������ ��������� ������������� ��� ���������� ����������� " +
                "�� ������ �������� �������� �������� (��������� ����������).\r\n" +
                "\r\n\t����������:\r\n" +
                "\t1) ������� ����, ���������� ������� ������� ������ ���������� ��������.\r\n" +
                "\t2) ������� ���������� ���������, ������� ����� ����������� � ���������� " +
                "����������� ���������������������� ��������.\r\n" +
                "\t3) �������� ���������� ��������� ���������� (������������), ������� ����� ������������ ����� �����.\r\n" +
                "\t4) �������� ���������� ��������� ��������, ������� ����� �������� �� �����.\r\n" +
                "\t5) ����������� �������� �������������, �� �������� ����� ����������� ������.\r\n" +
                "\t6) ������� ���������� ������ � ������.\r\n" +
                "\t7) ������� \"����\"\r\n" +
                "\t8) ��������� � ������ ����������������� ����������.\r\n " +
                "\t9) ������� ��������, ���������, �� ������������� (����� �� 0 �� 1, ������� � ����� ���� 1) � " +
                "������ ��������.\r\n" +
                "\t10) ������� ������ \"����\", ����� ���� ��������� ������� ���������� ������ " +
                "� ������������ ��������.\r\n" +
                "\t11) ��������� � ������ ������� �������� �������� �������� ��� ������� ��������.\r\n" +
                "\t12) ������� �������� ���������� � �� �������� ������������� ��� ������ ��������� ��� ������� ��������.\r\n" +
                "\t13) ��������� �� ������� ������� ������� � ������� ������ \"����\", ����� ���� ��������� �������\r\n" +
                "������������� �� ��������� � ����������� �������� � ������� �������.\r\n" +
                "\t14) � ������� ������� ��������� ������� �������, �� ������� ����������: ������� ������������� �� ��������\r\n" +
                "�������� ��������, ����������� ��������.\r\n" +
                "\t15) �� �������� ������������ �������� � ���������������� �������� ������������ ���������� ��������������� ��������.";
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
            column[0] = new DataColumn("��������", Type.GetType("System.String"));
            column[1] = new DataColumn("��������������", Type.GetType("System.Single"));
            for (int i = 2; i < situationsNumber + 2; i++)
            {
                column[i] = new DataColumn("�������� " + (i - 1).ToString(), Type.GetType("System.Single"));
            }
            for (int i = 0; i < situationsNumber + 2; i++)
            {
                evaluationTabel.Columns.Add(column[i]);
            }

            for (int i = 0; i < expertsNumber; i++)
            {
                evaluationTabel.Rows.Add(new object[] { ("������� " + (i + 1).ToString()).ToString() });
            }
            evaluationTabel.Rows.Add(new object[] { "���������� ������" });
            evaluationTabel.Rows.Add(new object[] { "����������� ��������" });

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
            column[0] = new DataColumn("��������", Type.GetType("System.String"));
            column[1] = new DataColumn("��������������", Type.GetType("System.String"));
            for (int i = 2; i < situationsNumber + 2; i++)
            {
                column[i] = new DataColumn("�������� " + (i - 1).ToString(), Type.GetType("System.String"));
            }
            for (int i = 0; i < situationsNumber + 2; i++)
            {
                coefficientTabel.Columns.Add(column[i]);
            }
            coefficientTabel.Rows.Add(new object[] { "����������� ��������, ��������������� ��������" });

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
                    coefficientTabel.Rows[0][i + 2] = "����� �������";
                }
                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] > 11 
                    && (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] <= 22)
                {
                    coefficientTabel.Rows[0][i + 2] = "�������";
                }
                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] > 22 
                    && (float)evaluationTabel.Rows[expertsNumber + 1][i + 2] <= 33)
                {
                    coefficientTabel.Rows[0][i + 2] = "���������";
                }
                if ((float)evaluationTabel.Rows[expertsNumber + 1][i + 2] > 33)
                {
                    coefficientTabel.Rows[0][i + 2] = "������������� (������)";
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
                buttonsClear[i].Text = "��������";
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
                indivTabelsColumn[0] = new DataColumn("�������� ����������", Type.GetType("System.String"));
                individualTabels[i].Columns.Add(indivTabelsColumn[0]);
                for (int j = 1; j < situationsNumber + 1; j++)
                {
                    indivTabelsColumn[j] = new DataColumn(evaluationTabel.Columns[j + 1].ColumnName, Type.GetType("System.Single"));
                    individualTabels[i].Columns.Add(indivTabelsColumn[j]);
                }

                for (int j = 0; j < controlOptionsNumber; j++)
                {
                    individualTabels[i].Rows.Add(new object[] { ("������� ���������� " + (j + 1).ToString()).ToString() });
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
            indivTabelsColumn[0] = new DataColumn("�������� ����������", Type.GetType("System.String"));
            individualTabels[index].Columns.Add(indivTabelsColumn[0]);
            for (int j = 1; j < situationsNumber + 1; j++)
            {
                indivTabelsColumn[j] = new DataColumn(evaluationTabel.Columns[j + 1].ColumnName, Type.GetType("System.Single"));
                individualTabels[index].Columns.Add(indivTabelsColumn[j]);
            }

            for (int j = 0; j < situationsNumber; j++)
            {
                individualTabels[index].Rows.Add(new object[] { ("������� ���������� " + (j + 1).ToString()).ToString() });
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
            DataTable linguistic�onsistencyTabel = new DataTable();

            DataColumn[] column = new DataColumn[situationsNumber + 1];
            column[0] = new DataColumn("�������� ����������", Type.GetType("System.String"));
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

            DataColumn[] consistency�olumn = new DataColumn[situationsNumber + 1];
            consistency�olumn[0] = new DataColumn("�������� ����������", Type.GetType("System.String"));
            consistencyTabel.Columns.Add(consistency�olumn[0]);
            for (int i = 1; i < situationsNumber + 1; i++)
            {
                consistency�olumn[i] = new DataColumn(evaluationTabel.Columns[i + 1].ColumnName, Type.GetType("System.Single"));
                consistencyTabel.Columns.Add(consistency�olumn[i]);
            }

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                consistencyTabel.Rows.Add(new object[] { individualTabels[0].Rows[i][0] });
            }

            DataColumn[] linguisticColumn = new DataColumn[situationsNumber + 1];
            linguisticColumn[0] = new DataColumn("�������� ����������", Type.GetType("System.String"));
            linguistic�onsistencyTabel.Columns.Add(linguisticColumn[0]);
            for (int i = 1; i < situationsNumber + 1; i++)
            {
                linguisticColumn[i] = new DataColumn(evaluationTabel.Columns[i + 1].ColumnName, Type.GetType("System.String"));
                linguistic�onsistencyTabel.Columns.Add(linguisticColumn[i]);
            }

            for (int i = 0; i < controlOptionsNumber; i++)
            {
                linguistic�onsistencyTabel.Rows.Add(new object[] { individualTabels[0].Rows[i][0] });
            }

            dataGridView2.DataSource = pivotTable;
            dataGridView5.DataSource = consistencyTabel;
            dataGridView6.DataSource = linguistic�onsistencyTabel;

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
                        linguistic�onsistencyTabel.Rows[i][j] = "����� �������";
                    }
                    if ((float)consistencyTabel.Rows[i][j] > 11
                        && (float)consistencyTabel.Rows[i][j] <= 22)
                    {
                        linguistic�onsistencyTabel.Rows[i][j] = "�������";
                    }
                    if ((float)consistencyTabel.Rows[i][j] > 22
                        && (float)consistencyTabel.Rows[i][j] <= 33)
                    {
                        linguistic�onsistencyTabel.Rows[i][j] = "���������";
                    }
                    if ((float)consistencyTabel.Rows[i][j] > 33)
                    {
                        linguistic�onsistencyTabel.Rows[i][j] = "������������� (������)";
                    }
                }
            }

            efficiencyTable = new DataTable();

            DataColumn[] efficiencyColumn = new DataColumn[2];
            efficiencyColumn[0] = new DataColumn("�������� ����������", Type.GetType("System.String"));
            efficiencyTable.Columns.Add(efficiencyColumn[0]);
            efficiencyColumn[1] = new DataColumn("������������� �� ���������", Type.GetType("System.Single"));
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