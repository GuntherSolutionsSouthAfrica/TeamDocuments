using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace Team_Documents___Pyed_Piper_WF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string SourceFileName;
        public string DestinationFileName = @"C:\\hmz\\ProcessedDocuments";
        private void Form1_Load(object sender, EventArgs e)
        {
            string Username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString();
            lblLoggedOnUser.Text = lblLoggedOnUser.Text + " " + Username;
        }

        private void FormatDocumentMargins(string DocToFormat)
        {
            var application = new Word.Application();            
                Word.Document document = application.Documents.Open(DocToFormat);

            try
            {            // format entire document first , remember 1 cm = 28.35 points


                document.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                document.PageSetup.BottomMargin = (float)85;
                document.PageSetup.TopMargin = (float)99;
                document.PageSetup.LeftMargin = (float)57;
                document.PageSetup.RightMargin = (float)57;
                document.PageSetup.Gutter = 0;
                document.PageSetup.GutterPos = Word.WdGutterStyle.wdGutterPosLeft;

                document.PageSetup.HeaderDistance = (float)35;
                document.PageSetup.FooterDistance = (float)31;

                    //document.SaveAs2(@"C:\hmz\ProcessedDocuments\InterOP_Processed_" + FileName);
                    document.Save();
                    document.Close();
                    application.Quit();
            }
            catch(Exception we)
            {
                MessageBox.Show(we.Message.ToString());
                document.Close();
                application.Quit();
            }


        }
        private void FormatHeaders_LibertyBrand(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            string imgHeader1 = @"C:\hmz\LibertyLogo.png";
            string HeaderLine = @"C:\hmz\Stanlib_HeaderLine.png";

            foreach (Word.Section section in document.Sections)
            {
                section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                lblProcessing.Text = "Processing Headers...";
                //wdHeaderFooterFirstPage would be for first page..

              
                header.Shapes.AddPicture(imgHeader1, 0, 1, -15, -15, 39,45);
                header.Shapes.AddPicture(HeaderLine, 0, 1, -15, 50, 520, 1);
                //header.Shapes.AddLine(28, 90, 558, 90);

            }

            document.Save();
            document.Close();
            application.Quit();

        }
        private void FormatHeaders(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            string imgHeader1 = @"C:\hmz\Stanlib_Logo_Board_Reports.png";
            string HeaderLine = @"C:\hmz\Stanlib_HeaderLine.png";

            foreach (Word.Section section in document.Sections)
            {
                section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                lblProcessing.Text = "Processing Headers...";
                //wdHeaderFooterFirstPage would be for first page..

                header.Shapes.AddPicture(imgHeader1, 0, 1, -15, -15, 137, 51);
                header.Shapes.AddPicture(HeaderLine, 0, 1, -15, 50, 520, 1);
                //header.Shapes.AddLine(28, 90, 558, 90);

            }

            document.Save();
            document.Close();
            application.Quit();

        }
        private void FormatFooters_LibertyBrand(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);


            string FooterLine = @"C:\hmz\Stanlib_FooterLine.png";
            try
            {
                foreach (Word.Section section in document.Sections)
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    lblProcessing.Text = "Processing Footers...";

                    footer.Shapes.AddPicture(FooterLine, 0, 1, -10, -20, 520, 1);

                    footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberRight, false);
                    footer.PageNumbers.StartingNumber = 0;
                    footer.Range.InsertBefore("Page  ");
                    footer.Range.Font.Size = 8;
                    footer.Range.Font.Name = "Arial";
                    footer.Range.Font.Bold = 1;

                }
                document.Save();
                document.Close();
                application.Quit();
            }
            catch (Exception el)
            {
                MessageBox.Show(el.Message.ToString());
                document.Close();
                application.Quit();
            }
            /// fix footers

        }

        private void FormatFooters(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            
            string FooterLine = @"C:\hmz\Stanlib_FooterLine.png";
            try
            {
                foreach (Word.Section section in document.Sections)
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    lblProcessing.Text = "Processing Footers...";

                    footer.Shapes.AddPicture(FooterLine, 0, 1, -10, -20, 520, 1);

                    footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberRight, false);
                    footer.PageNumbers.StartingNumber = 0;
                    footer.Range.InsertBefore("Page ");
                    footer.Range.Font.Size = 8;
                    footer.Range.Font.Name = "Arial";
                    footer.Range.Font.Bold = 1;

                }
                        document.Save();
                        document.Close();
                        application.Quit();
            }
            catch (Exception el)
            {
                MessageBox.Show(el.Message.ToString());
                document.Close();
                application.Quit();
            }
            /// fix footers


        }

        private void FormatParagraphHeadings(string DocToFormat)
        {
   
                var application = new Word.Application();
                Word.Document document = application.Documents.Open(DocToFormat);

            try
            {
                int PCount = document.Content.Paragraphs.Count;

                if (PCount > 0)
                {
                    for (int i = 1; i < document.Content.Paragraphs.Count; i++)
                    {
                        Word.Paragraph para = document.Content.Paragraphs[i];
                        Word.Style style = para.get_Style() as Word.Style;
                        lblProcessing.Text = "Processing All Paragraphs...";

                        if (style.NameLocal == "Heading 1")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;
                            para.Range.Font.Size = 12;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;

                            //para.Range.Paragraphs.KeepWithNext = 1;
                            para.Range.Paragraphs.LeftIndent = 0;
                            para.Range.Paragraphs.RightIndent = 0;
                            para.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
                            para.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                            para.Range.Paragraphs.SpaceAfter = 12;
                            para.Range.Paragraphs.SpaceBefore = 0;
                            para.Range.Paragraphs.LineSpacing = 14;
                            para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                            

                        }
                        if (style.NameLocal == "Heading 2")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        }

                        if (style.NameLocal == "Heading 3")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        }

                        if (style.NameLocal == "Heading 4")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        }



                        if (style.NameLocal == "Book Title")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Bold = 0;
                            para.Range.Italic = 0;
                            para.LineSpacing = 14;
                            para.SpaceAfter = 18;
                            para.SpaceBefore = 24;


                            //para.Range.Paragraphs.SpaceAfter = 18;
                            //para.Range.Paragraphs.SpaceBefore = 24;
                            //para.Range.Paragraphs.LineSpacing = 14;
                            //para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;

                        }

                    }

                    document.Save();
                    document.Close();
                    application.Quit();
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString());
                document.Close();
                application.Quit();
            }



        }

        private void FormatParagraphHeadings_LibertyBrand(string DocToFormat)
        {

            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            try
            {
                int PCount = document.Content.Paragraphs.Count;

                if (PCount > 0)
                {
                    for (int i = 1; i < document.Content.Paragraphs.Count; i++)
                    {
                        Word.Paragraph para = document.Content.Paragraphs[i];
                        Word.Style style = para.get_Style() as Word.Style;
                        lblProcessing.Text = "Processing All Paragraphs...";

                        if (style.NameLocal == "Heading 1")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorDarkTeal;
                            para.Range.Font.Size = 12;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;

                            //para.Range.Paragraphs.KeepWithNext = 1;
                            para.Range.Paragraphs.LeftIndent = 0;
                            para.Range.Paragraphs.RightIndent = 0;
                            para.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
                            para.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                            para.Range.Paragraphs.SpaceAfter = 12;
                            para.Range.Paragraphs.SpaceBefore = 0;
                            para.Range.Paragraphs.LineSpacing = 14;
                            para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                        }
                        if (style.NameLocal == "Heading 2")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlueGray;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        }

                        if (style.NameLocal == "Heading 3")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlueGray;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        }

                        if (style.NameLocal == "Heading 4")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlueGray;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Font.Bold = 1;
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        }



                        if (style.NameLocal == "Book Title")
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlueGray;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Bold = 0;
                            para.Range.Italic = 0;
                            para.LineSpacing = 14;
                            para.SpaceAfter = 18;
                            para.SpaceBefore = 24;


                            //para.Range.Paragraphs.SpaceAfter = 18;
                            //para.Range.Paragraphs.SpaceBefore = 24;
                            //para.Range.Paragraphs.LineSpacing = 14;
                            //para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;

                        }

                    }

                    document.Save();
                    document.Close();
                    application.Quit();
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString());
                document.Close();
                application.Quit();
            }



        }

        private void FormatNormalParagraphs_LibertyBrand(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            int PCount = document.Content.Paragraphs.Count;

            if (PCount > 0)
            {
                for (int i = 1; i < document.Content.Paragraphs.Count; i++)
                {
                    Word.Paragraph para = document.Content.Paragraphs[i];
                    Word.Style style = para.get_Style() as Word.Style;

                    lblProcessing.Text = "Processing All Paragraphs...";
                    document.RemoveNumbers();

                    if (style.NameLocal == "Normal")
                    {
                        para.Range.Font.Color = Word.WdColor.wdColorBlack;
                        para.Range.Font.Size = 11;
                        para.Range.Font.Name = "Arial";
                        para.Range.Bold = 0;

                        para.Range.Paragraphs.LeftIndent = 1;
                        para.Range.Paragraphs.RightIndent = 0;
                        para.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                        para.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        para.Range.Paragraphs.SpaceAfter = 10;
                        para.Range.Paragraphs.SpaceBefore = 0;
                        para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;


                    }
                }
            }
            document.Save();
            document.Close();
            application.Quit();
        }
        private void FormatNormalParagraphs(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            int PCount = document.Content.Paragraphs.Count;

            if (PCount > 0)
            {
                for (int i = 1; i < document.Content.Paragraphs.Count; i++)
                {
                    Word.Paragraph para = document.Content.Paragraphs[i];
                    Word.Style style = para.get_Style() as Word.Style;
                    Word.Range range = para.Range;
                    
                    lblProcessing.Text = "Processing All Paragraphs...";
                    //document.RemoveNumbers();
                    
                    if (style.NameLocal == "Normal")
                    {
                        
                        {
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;
                            para.Range.Font.Size = 11;
                            para.Range.Font.Name = "Arial";
                            para.Range.Bold = 0;

                            para.Range.Paragraphs.LeftIndent = 1;
                            para.Range.Paragraphs.RightIndent = 0;
                            para.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                            para.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                            para.Range.Paragraphs.SpaceAfter = 10;
                            para.Range.Paragraphs.SpaceBefore = 0;
                            para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                        }


                        
                    }
                }
            }
            document.Save();
            document.Close();
            application.Quit();

        }

        private void FormatAllTables(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            int PCount = document.Content.Tables.Count;
            Word.Tables tables = document.Tables;
            Word.Table table;

            if (PCount > 0)
            {
                for (int i = 1; i <= document.Content.Tables.Count; i++)
                {
                    
                    lblProcessing.Text = "Processing All Tables...";
                    table = tables[i];
                    int cols = table.Columns.Count;
                    for (int ii = 1; ii < table.Columns.Count + 1; ii++)
                    {
                            Word.Cell cell = table.Cell(1, ii);

                            cell.Range.Font.Color = Word.WdColor.wdColorWhite;
                            cell.Range.Font.Size = 11;
                            cell.Range.Font.Name = "Arial";
                            cell.Range.Bold = 1;
                    }
                    int rows = table.Rows.Count;
                    for(int r = 2; r < table.Rows.Count + 1; r++)
                    {
                        for (int ii = 1; ii < table.Columns.Count + 1; ii++)
                        {
                            Word.Cell cell = table.Cell(r, ii);

                            cell.Range.Font.Color = Word.WdColor.wdColorBlack;
                            cell.Range.Font.Size = 11;
                            cell.Range.Font.Name = "Arial";
                            cell.Range.Bold = 1;
                        }

                    }
                   
                }
            }
            document.Save();
            document.Close();
            application.Quit();
        }
        private void FormatBullets(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            int PCount = document.Content.Paragraphs.Count;

            if (PCount > 0)
            {
                for (int i = 1; i < document.Content.Paragraphs.Count; i++)
                {
                    Word.Paragraph para = document.Content.Paragraphs[i];
                    Word.Style style = para.get_Style() as Word.Style;

                    lblProcessing.Text = "Processing bullets...";

                    if (style.NameLocal == "List Paragraph" || style.NameLocal == "Bullet 01" || style.NameLocal == "Bullet 02")
                    {
                        para.Range.Font.Color = Word.WdColor.wdColorBlack;
                        para.Range.Font.Size = 11;
                        para.Range.Font.Name = "Arial";
                        para.Range.Bold = 0;

                        para.Range.Paragraphs.SpaceAfter = 6;
                        para.Range.Paragraphs.SpaceBefore = 6;
                        para.Range.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                        //para.Range.Paragraphs.LeftIndent = 1;
                        //para.Range.Paragraphs.RightIndent = 0;
                        //para.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                        //para.Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;

                    }
                }
            }
            document.Save();
            document.Close();
            application.Quit();


        }

        private void FormatTableOfContents_LibertyBrand(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            int tCount = document.TablesOfContents.Count;
            if (tCount > 0)
            {
                try
                {
                    lblProcessing.Text = "Processing Table of Contents...";
                    //Table of Contents
                    document.TablesOfContents[1].IncludePageNumbers = true;
                    document.TablesOfContents[1].Range.Font.Name = "Arial";
                    document.TablesOfContents[1].Range.Font.Size = 11;
                    document.TablesOfContents[1].Range.Font.Bold = 1;
                    document.TablesOfContents[1].Range.Font.Color = Word.WdColor.wdColorBlueGray;
                    document.TablesOfContents[1].RightAlignPageNumbers = true;

                    document.TablesOfContents[1].Range.Italic = 0;
                    document.TablesOfContents[1].Range.Paragraphs.LineSpacing = 14;
                    document.TablesOfContents[1].Range.Paragraphs.SpaceAfter = 18;
                    document.TablesOfContents[1].Range.Paragraphs.SpaceBefore = 24;
                }
                catch (Exception er)
                {
                    MessageBox.Show(er.Message.ToString());
                }
            }


            document.Save();
            document.Close();
            application.Quit();
        }
        private void FormatTableOfContents(string DocToFormat)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Open(DocToFormat);

            int tCount = document.TablesOfContents.Count;
            if (tCount > 0)
            {
                try
                {
                    lblProcessing.Text = "Processing Table of Contents...";
                    //Table of Contents
                    document.TablesOfContents[1].IncludePageNumbers = true;
                    document.TablesOfContents[1].Range.Font.Name = "Arial";
                    document.TablesOfContents[1].Range.Font.Size = 12;
                    document.TablesOfContents[1].Range.Font.Bold = 1;
                    document.TablesOfContents[1].Range.Font.Color = Word.WdColor.wdColorBlack;
                    document.TablesOfContents[1].RightAlignPageNumbers = true;

                    document.TablesOfContents[1].Range.Italic = 0;
                    document.TablesOfContents[1].Range.Paragraphs.LineSpacing = 14;
                    document.TablesOfContents[1].Range.Paragraphs.SpaceAfter = 18;
                    document.TablesOfContents[1].Range.Paragraphs.SpaceBefore = 24;
                }
                catch (Exception er)
                {
                    MessageBox.Show(er.Message.ToString());
                }
            }


            document.Save();
            document.Close();
            application.Quit();
        }

        private void MoveFileToProcessingFolder(string SourceFileName, string DestinationFileName)
        {
            try {
                        if(File.Exists(DestinationFileName))
                            {
                                File.Delete(DestinationFileName);
                            }
                        File.Copy(SourceFileName, DestinationFileName);
                }
            catch (Exception ee)
                {
                MessageBox.Show(ee.Message.ToString());
                DestinationFileName = @"C:\\hmz\\ProcessedDocuments";
               }
            
        }
        private void btnUploadFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdg = new OpenFileDialog();
            fdg.Title = "Team Documents - Pyed Piper : Open File";
            fdg.Filter = "docx |*.docx|doc |*.doc";
            
            try
            {
                if (fdg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFileName.Text = fdg.FileName.ToString();
                    SourceFileName = fdg.SafeFileName.ToString();

                    DestinationFileName = DestinationFileName + "\\" + fdg.SafeFileName.ToString();
                    MoveFileToProcessingFolder(fdg.FileName.ToString(), DestinationFileName);
                }
            }
            catch (Exception es)
            {
                MessageBox.Show(es.Message.ToString());
                txtFileName.Text = "";
                DestinationFileName = @"C:\\hmz\\ProcessedDocuments";
            }
        }

        private void btnProcessFile_Click(object sender, EventArgs e)
        {

            if ((string)comboBox1.SelectedItem == "STANLIB BRAND (FULL i.e.with Headers, Footers & TOC)")
            {
                    panel2.Visible = false;
                    lblProcessing.Visible = true;

                    FormatDocumentMargins(DestinationFileName);
                    FormatHeaders(DestinationFileName);
                    FormatFooters(DestinationFileName);
                    FormatNormalParagraphs(DestinationFileName);
                    FormatParagraphHeadings(DestinationFileName);
                    FormatBullets(DestinationFileName);
                    FormatAllTables(DestinationFileName);
                    FormatTableOfContents(DestinationFileName);
                    //reset destination folder
                    txtFileName.Text = "";
                    DestinationFileName = @"C:\\hmz\\ProcessedDocuments";

                    lblProcessing.Visible = false;
                    lblProcessing.Text = "";
                    panel2.Visible = true;

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = DestinationFileName,
                    UseShellExecute = true,
                    Verb = "open"
                });

            }
            if ((string)comboBox1.SelectedItem == "STANLIB BRAND (PARTIAL)")
            {
                panel2.Visible = false;
                lblProcessing.Visible = true;

                FormatDocumentMargins(DestinationFileName);
                //FormatHeaders(DestinationFileName);
                //FormatFooters(DestinationFileName);
                FormatNormalParagraphs(DestinationFileName);
                FormatParagraphHeadings(DestinationFileName);
                FormatBullets(DestinationFileName);
                FormatAllTables(DestinationFileName);
                //FormatTableOfContents(DestinationFileName);
                //reset destination folder
                txtFileName.Text = "";
                DestinationFileName = @"C:\\hmz\\ProcessedDocuments";

                lblProcessing.Visible = false;
                lblProcessing.Text = "";
                panel2.Visible = true;

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = DestinationFileName,
                    UseShellExecute = true,
                    Verb = "open"
                });

            }
            if ((string)comboBox1.SelectedItem == "LIBERTY BRAND  (FULL i.e. with Headers , Footers & TOC)")
            {
                panel2.Visible = false;
                lblProcessing.Visible = true;

                FormatDocumentMargins(DestinationFileName);
                FormatHeaders_LibertyBrand(DestinationFileName);
                FormatFooters_LibertyBrand(DestinationFileName);
                FormatNormalParagraphs(DestinationFileName);
                FormatParagraphHeadings_LibertyBrand(DestinationFileName);
                FormatBullets(DestinationFileName);
                FormatTableOfContents_LibertyBrand(DestinationFileName);
                //reset destination folder
                txtFileName.Text = "";
                DestinationFileName = @"C:\\hmz\\ProcessedDocuments";

                lblProcessing.Visible = false;
                lblProcessing.Text = "";
                panel2.Visible = true;

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = DestinationFileName,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
                if ((string)comboBox1.SelectedItem == "LIBERTY BRAND (PARTIAL)")
                {
                    panel2.Visible = false;
                    lblProcessing.Visible = true;

                    FormatDocumentMargins(DestinationFileName);
                    //FormatHeaders_LibertyBrand(DestinationFileName);
                    //FormatFooters_LibertyBrand(DestinationFileName);
                    FormatNormalParagraphs(DestinationFileName);
                    FormatParagraphHeadings_LibertyBrand(DestinationFileName);
                    FormatBullets(DestinationFileName);
                    //FormatTableOfContents_LibertyBrand(DestinationFileName);
                    //reset destination folder
                    txtFileName.Text = "";
                    DestinationFileName = @"C:\\hmz\\ProcessedDocuments";

                    lblProcessing.Visible = false;
                    lblProcessing.Text = "";
                    panel2.Visible = true;

                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = DestinationFileName,
                        UseShellExecute = true,
                        Verb = "open"
                    });

                }

            }

        }
}
