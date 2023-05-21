using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace lab26
{
    public partial class comment_txt : Form
    {
        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        Word.Document doc = new Word.Document();

        public comment_txt()
        {
            InitializeComponent();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                doc.Close();
                app.Quit();
                doc = null;
                app = null;
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] Findtxt = new string[] { "Name1", "Name2", "Fax1", "Fax2", "Tel1", "Tel2", "Theme", "date", "comm"};

            object fileName = @"C:\Users\User\Desktop\Fax1.dotx";
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            for(int i = 0; i < Findtxt.Length; i++)
            {

                object findText = Findtxt[i];
                object replaceWith = Controls.Find(Findtxt[i], true)[0].Text;
                object replace = 2;

                app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                ref replace, ref missing, ref missing, ref missing, ref missing);
            }

            app.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\Users\User\Desktop\Fax1.dotx");
        }
    }
}
