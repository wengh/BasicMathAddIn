using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace BasicMathAddIn
{
    public partial class Ribbon1
    {
        Random rand = new Random();

        private void Ribbon1_Load (object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click (object sender, RibbonControlEventArgs e)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Content.Font.NameAscii = "Arial";
                doc.Content.Font.Size = 28;
                doc.PageSetup.TextColumns.SetCount(2);

                var selected = new List<int>();
                if (checkBox1.Checked)
                {
                    selected.Add(0);
                }
                if (checkBox2.Checked)
                {
                    selected.Add(1);
                }
                if (checkBox3.Checked)
                {
                    selected.Add(2);
                }
                if (checkBox4.Checked)
                {
                    selected.Add(3);
                }
                if (selected.Count == 0)
                {
                    MessageBox.Show("请选择至少一种题目");
                    return;
                }
                int pages = Convert.ToInt16(editBox2.Text);
                int length = 30 * pages;
                int maxA = 1 + Convert.ToInt16(editBox1.Text);
                int minA = Math.Min(Convert.ToInt16(editBox3.Text), maxA);
                int maxS = 1 + Convert.ToInt16(editBox4.Text);
                int minS = Math.Min(Convert.ToInt16(editBox5.Text), maxS);
                int maxM = 1 + Convert.ToInt16(editBox6.Text);
                int minM = Math.Min(Convert.ToInt16(editBox7.Text), maxS);
                int maxD = 1 + Convert.ToInt16(editBox8.Text);
                int minD = Math.Max(Math.Min(Convert.ToInt16(editBox9.Text), maxS), 1);
                var lst = new List<string>();

                for (int i = 0; i < length; i++)
                {
                    int type;
                    type = selected[rand.Next(selected.Count)];
                    int result, val1, val2;
                    switch (type)
                    {
                        case 0:
                            result = rand.Next(minA, maxA);
                            val1 = rand.Next(result + 1);
                            val2 = result - val1;
                            lst.Add(string.Format("{0} + {1} = ", val1, val2));
                            break;

                        case 1:
                            val1 = rand.Next(minS, maxS);
                            val2 = rand.Next(val1 + 1);
                            lst.Add(string.Format("{0} - {1} = ", val1, val2));
                            break;
                        case 2:
                            val1 = rand.Next(minM, maxM);
                            val2 = rand.Next(minM, maxM);
                            lst.Add(string.Format("{0} × {1} = ", val1, val2));
                            break;
                        case 3:
                            val1 = rand.Next(minD, maxD);
                            val2 = rand.Next(minD, maxD);
                            result = val1 * val2;
                            lst.Add(string.Format("{0} ÷ {1} = ", result, val1));
                            break;
                    }
                }
                string s = string.Join("\n", lst);
                if (dropDown1.SelectedItem.Label == "覆盖")
                {
                    doc.Content.Text = s;
                }
                else
                {
                    doc.Content.Text += s;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
        }
    }
}
