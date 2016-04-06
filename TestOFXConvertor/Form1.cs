using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace TestOFXConvertor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // load ofx file
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();

            string filename = ofd.FileName;

            string text = null;

            using (var fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                using (var sr = new StreamReader(fs))
                {
                    text = sr.ReadToEnd();
                }
            }
            
            // strip header
            int startPos = text.IndexOf("<OFX>");
            text = text.Substring(startPos);
            
            // strip illegal xml section
            var r = new Regex("<SONRS>[\n0-9A-Z<>\\./]*</SONRS>");
            text = r.Replace(text, "");

            // load into object graph
            var doc = XDocument.Parse(text);

            var items =
                (from el in doc.Descendants("STMTTRN")
                select new {
                    type = el.Element("TRNTYPE").Value,
                    date = el.Element("DTPOSTED").Value,
                    amount = el.Element("TRNAMT").Value,
                    FITID = el.Element("FITID").Value,
                    name = el.Element("NAME").Value,
                    memo = el.Element("MEMO")?.Value,
                }).ToList();

            MessageBox.Show("Done");

        }
    }

}
