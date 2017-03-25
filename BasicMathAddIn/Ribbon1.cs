using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace BasicMathAddIn
{
    public partial class Ribbon1
    {
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

                bool add = checkBox1.Checked;
                bool sub = checkBox2.Checked;
                if (!(add || sub))
                {
                    return;
                }
                int pages = Convert.ToInt16(editBox2.Text);
                int length = 30 * pages;
                int maxA = 1 + Convert.ToInt16(editBox1.Text);
                int minA = Math.Min(Convert.ToInt16(editBox3.Text), maxA);
                int maxS = 1 + Convert.ToInt16(editBox4.Text);
                int minS = Math.Min(Convert.ToInt16(editBox5.Text), maxS);
                var lst = new List<string>();
                var rand = new Random();

                for (int i = 0; i < length; i++)
                {
                    int type;
                    if (add && sub)
                    {
                        type = rand.Next(0, 2);
                    }
                    else
                    {
                        type = add ? 0 : 1;
                    }

                    if (type == 0)
                    {
                        int result = rand.Next(minA, maxA);
                        int val1 = rand.Next(result + 1);
                        int val2 = result - val1;
                        lst.Add(string.Format("{0} + {1} = ", val1, val2));
                    }
                    else
                    {
                        int val1 = rand.Next(minS, maxS);
                        int val2 = rand.Next(val1 + 1);
                        lst.Add(string.Format("{0} - {1} = ", val1, val2));
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
