using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string filename;
        string filelocation;
        public Form1()
        {
            InitializeComponent();
        }

        private void bnt_ok_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                pdftext.Text += ReadPdfFile(textBox1.Text);

                string pathFile = System.IO.Path.GetDirectoryName(filelocation);
                string newFileName = System.IO.Path.GetFileNameWithoutExtension(filelocation) + ".txt";

                filelocation = pathFile+System.IO.Path.DirectorySeparatorChar+ newFileName;
                  System.IO.File.WriteAllText(filelocation, pdftext.Text);
                MessageBox.Show("Archivo Guardado!", "de PDF a TEXT logrado!");
               
            }

            

        }

        private void btn_cargar_Click(object sender, EventArgs e)
        {
            
        }

        private void openFile()
        {
           
            openFileDialog1.Filter = "eBook (*.pdf)|*.pdf";
            openFileDialog1.ShowDialog();

            filelocation = openFileDialog1.InitialDirectory + openFileDialog1.FileName;
            textBox1.Text = filelocation;
            filename = filelocation;
            if (openFileDialog1.FileName != null)
            {
                // use the LoadFile(ByVal fileName As String) function for open the pdf in control
                axAcroPDF1.LoadFile(openFileDialog1.FileName);
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            openFile();
        }

        public string ReadPdfFile(string fileName)
        {
            StringBuilder text = new StringBuilder();

            if (File.Exists(fileName))
            {
                PdfReader pdfReader = new PdfReader(fileName);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {

        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        public static void OpenAndAddTextToWordDocument(string filepath, string txt)
        {
            //    // abriendo un documento.
            //    //WordprocessingDocument wordprocessingDocument =
            //    // WordprocessingDocument.Open(filepath, true);
            //    WordprocessingDocument wordprocessingDocument =
            //          WordprocessingDocument.Open(filepath, false);
            //         // asignando
            //        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            //    // agregando el nuevo testo.
            //    Paragraph para = body.AppendChild(new Paragraph());
            //    Run run = para.AppendChild(new Run());
            //    run.AppendChild(new Text(txt));

            //    // cerrando el documento.
            //    wordprocessingDocument.Close();

           

            
        }

        private void btn_about_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.Show();
        }
    }
}
