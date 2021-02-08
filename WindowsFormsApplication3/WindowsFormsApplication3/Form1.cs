using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net.Mail;
using System.Media;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dataGridView1.Columns.Add("UID", "UID");
            this.dataGridView1.Columns.Add("Order", "Order");
        }

        private void playSimpleSound()
        {
            SoundPlayer simpleSound = new SoundPlayer(@"c:\Windows\Media\Speech Sleep.wav");
            simpleSound.Play();
        }

        private void playErrorSound()
        {
            SoundPlayer simpleSound = new SoundPlayer(@"C:\Sounds\RobotBlip.wav");
            simpleSound.Play();
        }

        void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.Text.Contains("RC") || textBox1.Text.Contains("32"))
                {
                    textBox2.Focus();
                    label7.Text = textBox1.Text;
                    label7.ForeColor = System.Drawing.Color.Green;
                    e.SuppressKeyPress = true;
                    playSimpleSound();
                }
                else
                {
                    label7.Text = "Bad Scan";
                    playErrorSound();
                    label7.ForeColor = System.Drawing.Color.Red;
                    textBox1.Clear();
                }
            }
        }

        void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!textBox2.Text.Contains("RC"))
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(this.dataGridView1, textBox1.Text, textBox2.Text, DateTime.Now);
                    this.dataGridView1.Rows.Add(row);
                    e.SuppressKeyPress = true;
                    label7.Text = textBox2.Text;
                    label7.ForeColor = System.Drawing.Color.Green;
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox1.Focus();
                    playSimpleSound();
                }
                else
                {
                    label7.Text = "Bad Scan";
                    playErrorSound();
                    label7.ForeColor = System.Drawing.Color.Red;
                    textBox2.Clear();
                }
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            string path = "\\Scanned Returns\\";
            string sDrivePath = "S:\\" + path;
            string date = DateTime.Now.ToString("d/MMM/yyyy");
            string Today = DateTime.Now.ToString("MMM-y");
            string filepath = sDrivePath + "Scanned Returns " + Today + ".xlsx";
            if (!File.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
            copyAllToClipboard();
            Excel.Application xlexcel;
            Excel.Workbook xlworkbook;
            Excel.Worksheet xlworksheet;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            if (!File.Exists(filepath))
            {
                xlworkbook = xlexcel.Workbooks.Add(Type.Missing);
                xlworkbook.SaveAs(filepath);
                xlworkbook.Close();
            }
            xlworkbook = xlexcel.Workbooks.Open(filepath);
            xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
            xlworksheet.Columns[1].ColumnWidth = 18;
            xlworksheet.Columns[2].ColumnWidth = 18;
            Excel.Range XlRange = (Excel.Range)xlworksheet.Cells[xlworksheet.Rows.Count, 1];
            long lastRow = (long)XlRange.get_End(Excel.XlDirection.xlUp).Row;
            Excel.Range Trailer = (Excel.Range)xlworksheet.Cells[lastRow + 1, 1];
            Excel.Range today = (Excel.Range)xlworksheet.Cells[lastRow + 1, 2];
            Excel.Range Stock = (Excel.Range)xlworksheet.Cells[lastRow + 2, 1];
            Trailer.Select();
            Trailer.Value2 = "MT" + textBox3.Text;
            today.Select();
            today.Value2 = date.ToString();
            Stock.Select();
            xlworksheet.PasteSpecial(Stock, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            xlworkbook.Save();
            xlworkbook.Close();
            xlexcel.Quit();
            if (File.Exists(filepath))
            {
                this.Close();
            }
        }

        private void email()
        {
            try
            {
                SmtpClient client = new SmtpClient("smtp.office365.com", 587);
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential(textBox8.Text, textBox4.Text);
                MailAddress from = new MailAddress(textBox8.Text, String.Empty, System.Text.Encoding.UTF8);
                MailAddress to = new MailAddress(comboBox1.Text);
                MailMessage message = new MailMessage(from, to);
                MailMessage mail = new MailMessage();
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Body = textBox6.Text;
                message.Attachments.Add(new Attachment(textBox5.Text));
                message.Subject = "Scanned returns ";
                message.SubjectEncoding = System.Text.Encoding.UTF8;
                client.Send(message);
                MessageBox.Show("Email Sent Successfully");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }

        }


        private void copyAllToClipboard()
        {
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox4.PasswordChar = default(char);
            }
            else
            {
                textBox4.UseSystemPasswordChar = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                DialogResult result = openFileDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    textBox5.Text = openFileDialog1.FileName;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            email();
        }

        private void release()
        {
            Excel.Application xlexcel;
            xlexcel = new Excel.Application();
            xlexcel.Quit();
            xlexcel = null;
        }
    }
}
