using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace BasicMathAddIn
{
    public partial class Ribbon1
    {
        MathSettingsManager manager;

        private void Ribbon1_Load (object sender, RibbonUIEventArgs e)
        {
            Properties.Settings.Default.Reload();
            try
            {
                manager = MathSettingsManager.Deserialize(Properties.Settings.Default.Prefs);
            }
            catch
            {
                manager = new MathSettingsManager();
                Properties.Settings.Default.Prefs = manager.Serialize();
                Properties.Settings.Default.Save();
                MessageBox.Show(Properties.Settings.Default.Prefs);
            }
            Refresh();
        }

        private void button1_Click (object sender, RibbonControlEventArgs e)
        {
            try
            {
                int seed = new Random().Next();
                var rand = new Random(seed);
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Content.Font.NameAscii = "Arial";
                doc.Content.Font.Size = 28;
                doc.PageSetup.TextColumns.SetCount(2);
                manager.lastSeed = seed;
                Refresh();

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

        private void button5_Click (object sender, RibbonControlEventArgs e)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Content.Font.NameAscii = "Arial";
                doc.Content.Font.Size = 28;
                doc.PageSetup.TextColumns.SetCount(2);
                var rand = new Random(manager.lastSeed);

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
                            lst.Add(string.Format("{0} + {1} = {2}", val1, val2, result));
                            break;

                        case 1:
                            val1 = rand.Next(minS, maxS);
                            val2 = rand.Next(val1 + 1);
                            lst.Add(string.Format("{0} - {1} = {2}", val1, val2, val1 - val2));
                            break;
                        case 2:
                            val1 = rand.Next(minM, maxM);
                            val2 = rand.Next(minM, maxM);
                            lst.Add(string.Format("{0} × {1} = {2}", val1, val2, val1 * val2));
                            break;
                        case 3:
                            val1 = rand.Next(minD, maxD);
                            val2 = rand.Next(minD, maxD);
                            result = val1 * val2;
                            lst.Add(string.Format("{0} ÷ {1} = {2}", result, val1, val2));
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

        private void Save (object sender, RibbonControlEventArgs e)
        {
            var s = new MathSettings();

            s.Pages = editBox2.Text;
            s.Mode = dropDown1.SelectedItemIndex;
            s.Add = checkBox1.Checked;
            s.Sub = checkBox2.Checked;
            s.Mul = checkBox3.Checked;
            s.Div = checkBox4.Checked;

            s.MaxA = editBox1.Text;
            s.MinA = editBox3.Text;
            s.MaxS = editBox4.Text;
            s.MinS = editBox5.Text;
            s.MaxM = editBox6.Text;
            s.MinM = editBox7.Text;
            s.MaxD = editBox8.Text;
            s.MinD = editBox9.Text;

            manager.Add(editBox10.Text, s);
            Refresh();
        }

        public void Refresh ()
        {
            if (manager != null)
            {
                Properties.Settings.Default.Prefs = manager.Serialize();
                Properties.Settings.Default.Save();
                Properties.Settings.Default.Upgrade();
                dropDown2.Items.Clear();
                foreach (string key in manager.Keys)
                {
                    RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                    item.Label = key;
                    dropDown2.Items.Add(item);
                }
            }
            else
            {
                MessageBox.Show("");
            }
        }

        private void LoadPref (object sender, RibbonControlEventArgs e)
        {
            string name;
            try
            {
                name = dropDown2.SelectedItem.Label;
            }
            catch (NullReferenceException)
            {
                return;
            }
            var s = manager[name];

            editBox2.Text = s.Pages;
            dropDown1.SelectedItemIndex = s.Mode;
            checkBox1.Checked = s.Add;
            checkBox2.Checked = s.Sub;
            checkBox3.Checked = s.Mul;
            checkBox4.Checked = s.Div;

            editBox1.Text = s.MaxA;
            editBox3.Text = s.MinA;
            editBox4.Text = s.MaxS;
            editBox5.Text = s.MinS;
            editBox6.Text = s.MaxM;
            editBox7.Text = s.MinM;
            editBox8.Text = s.MaxD;
            editBox9.Text = s.MinD;

            editBox10.Text = name;
        }

        private void RemovePref (object sender, RibbonControlEventArgs e)
        {
            string name;
            try
            {
                name = dropDown2.SelectedItem.Label;
            }
            catch (NullReferenceException)
            {
                return;
            }
            if (MessageBox.Show("是否删除“" + name + "”？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                manager.Remove(name);
                Refresh();
            }
        }
    }

    [Serializable]
    public class MathSettings
    {
        public string Pages;
        public int Mode;
        public bool Add;
        public bool Sub;
        public bool Mul;
        public bool Div;

        public string MaxA;
        public string MinA;
        public string MaxS;
        public string MinS;
        public string MaxM;
        public string MinM;
        public string MaxD;
        public string MinD;
    }

    [Serializable]
    public class MathSettingsManager : Dictionary<string, MathSettings>
    {
        public int lastSeed;

        public string Serialize ()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                new BinaryFormatter().Serialize(stream, this);
                return Convert.ToBase64String(stream.ToArray());
            }
        }

        public static MathSettingsManager Deserialize (string str)
        {
            byte[] bytes = Convert.FromBase64String(str);

            using (MemoryStream stream = new MemoryStream(bytes))
            {
                return new BinaryFormatter().Deserialize(stream) as MathSettingsManager;
            }
        }

        public MathSettingsManager () { }

        protected MathSettingsManager (SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
