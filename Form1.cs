using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace asposeWords
{
    public partial class Form1 : Form
    {
        public enum StyleNames { None, Heading1, Heading2, Quote }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ProcessMarkdown();
        }

        
        private void ProcessMarkdown()
        {
            var docPath = "C:\\temp\\aspose\\";
            var fileToLoad = "snippet.md";

            var doc = new Document(Path.Combine(docPath, fileToLoad));


            var options = new FindReplaceOptions();
            //options.ApplyFont.Bold = false;
            options.ApplyFont.Italic = true;

            _ = doc.Range.Replace("demonstration", "code demonstration", options);



            var paragraph0 = CreateParagraph(doc, StyleNames.Heading1);
            _ = paragraph0.AppendChild(CreateRun(doc, "This is a Heading 1"));

            var paragraph1 = CreateParagraph(doc);
            _ = paragraph1.AppendChild(CreateRun(doc, "This is some bold text", true));

            var paragraph2 = CreateParagraph(doc, StyleNames.Heading2);
            _ = paragraph2.AppendChild(CreateRun(doc, "This is a Heading 2"));

            var paragraph3 = CreateParagraph(doc);
            _ = paragraph3.AppendChild(CreateRun(doc, "This is some Italic text", false, true));

            var paragraph4 = CreateParagraph(doc, StyleNames.Quote);
            _ = paragraph4.AppendChild(CreateRun(doc, "This is a quote or something profound"));



            _ = doc.Save(Path.Combine(docPath, "snippetModified.md"), SaveFormat.Markdown);

            //doc.Protect(ProtectionType.ReadOnly, "password");
            //_ = doc.Save(Path.Combine(docPath, "snippetModified.docx"), SaveFormat.Docx);


            //doc.Print("Foxit Reader PDF Printer");                                   
        }

        private Paragraph CreateParagraph(Document doc, StyleNames styleName = StyleNames.None)
        {
            var section = new Section(doc);
            _ = doc.AppendChild(section);
            var body = new Body(doc);
            _ = section.AppendChild(body);
            var paragraph = new Paragraph(doc);
            _ = body.AppendChild(paragraph);

            switch (styleName)
            {
                case StyleNames.Heading1:
                    _ = paragraph.ParagraphFormat.StyleName = "Heading 1";
                    break;
                case StyleNames.Quote:
                    _ = paragraph.ParagraphFormat.StyleName = "Quote";
                    break;
                case StyleNames.Heading2:
                    _ = paragraph.ParagraphFormat.StyleName = "Heading 2";
                    break;
                default:
                    break;
            }            

            return paragraph;
        }

        private Run CreateRun(Document doc, string text, bool isBold = false, bool isItalic = false)
        {
            var textRun = new Run(doc);
            textRun.Text = text;
            textRun.Font.Bold = isBold;
            textRun.Font.Italic = isItalic;            
            return textRun;
        }        
    }
}
