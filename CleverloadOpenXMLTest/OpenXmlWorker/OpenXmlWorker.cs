using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Xps.Packaging;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Color = System.Drawing.Color;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using UnderlineValues = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Runtime.InteropServices;

namespace CleverloadOpenXMLTest
{
    internal class OpenXmlWorker
    {
        private string _tempDir = "\\coxmlTemp";

        public string LastXpsFileName { get; private set; }
        public string LastWordFileName { get; private set; }
        public string LastConvertedWordFileName { get; private set; }
        public XpsDocument XpsDocument { get; private set; }

        private static OpenXmlWorker _instance;

        public static OpenXmlWorker Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OpenXmlWorker();
                return _instance;
            }
        }

        private OpenXmlWorker()
        {
            if (!Directory.Exists(Environment.CurrentDirectory + _tempDir))
                Directory.CreateDirectory(Environment.CurrentDirectory + _tempDir);
        }

        private void AddPersonalRow(Body body, OpenXmlElement after, string firstName, string lastName, string city)
        {
            Paragraph par = new Paragraph();
            var run = new Run();
            run.RunProperties = new RunProperties();
            run.RunProperties.FontSize = new FontSize();
            run.RunProperties.FontSize.Val = new StringValue("28");
            run.AppendChild(new Break());
            run.AppendChild(new TabChar());
            run.AppendChild(new Text($"{firstName} {lastName}, {city}"));
            par.AppendChild(run);
            body.InsertAfter(par, after);
        }

        private void ApplyColorsFormat(Body body)
        {
            foreach (var runProp in body.Descendants<RunProperties>())
            {
                if (runProp.Color == null || runProp.Color.Val == "000000") continue;
                Color propColor = System.Drawing.ColorTranslator.FromHtml("#" + runProp.Color.Val);
                var hue = propColor.GetHue();
                if (hue >= 210 && hue <= 270)
                {
                    runProp.Color.Val = System.Drawing.ColorTranslator.ToHtml(Color.Green).Replace("#", "");
                }
                else if ((hue >= 0 && hue <= 30) || (hue >= 330 && hue <= 360))
                {
                    var child = new Underline();
                    child.Val = UnderlineValues.Single;
                    runProp.AppendChild(child);
                }
            }
        }

        private void AddSecondPageTable(Body body)
        {
            Table personalInfoTable = new Table();
            TableProperties tableProps = new TableProperties();
            AddTableBorders(tableProps);
            TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            tableProps.Append(tableStyle, tableWidth);
            TableGrid tableGrid = new TableGrid();
            personalInfoTable.AppendChild(tableProps);
            for (int x = 0; x < 6; x++)
                tableGrid.AppendChild(new GridColumn());
            personalInfoTable.AppendChild(tableGrid);
            for (int x = 0; x < 7; x++)
            {
                TableRow pesonalTableRow = new TableRow();
                List<TableCell> cells = new List<TableCell>();
                for (int i = 0; i < 6; i++)
                {
                    TableCell cell = new TableCell();
                    TableCellProperties tableCellProperties = new TableCellProperties();
                    tableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment();
                    tableCellProperties.TableCellVerticalAlignment.Val = TableVerticalAlignmentValues.Center;
                    VerticalMerge verticalMerge = new VerticalMerge();
                    HorizontalMerge horizontalMerge = new HorizontalMerge();
                    if (x == 0)
                    {
                        horizontalMerge.Val = i == 0 ? MergedCellValues.Restart : MergedCellValues.Continue;
                        tableCellProperties.AppendChild(horizontalMerge);
                        if (i == 0)
                            SetCellText(cell, "Time Table", true);
                        else
                            SetCellText(cell, "", false);
                    }
                    else if (i == 0)
                    {
                        verticalMerge.Val = x == 1 ? MergedCellValues.Restart : MergedCellValues.Continue;
                        tableCellProperties.AppendChild(verticalMerge);
                        if (x == 1)
                            SetCellText(cell, "Hours", true);
                        else
                            SetCellText(cell, "", false);
                    }
                    else if (x == 1) switch (i)
                        {
                            case 1:
                                SetCellText(cell, "Mon", true);
                                break;
                            case 2:
                                SetCellText(cell, "Tue", true);
                                break;
                            case 3:
                                SetCellText(cell, "Wed", true);
                                break;
                            case 4:
                                SetCellText(cell, "Thu", true);
                                break;
                            case 5:
                                SetCellText(cell, "Fri", true);
                                break;
                        }
                    else if (x == 4)
                    {
                        horizontalMerge.Val = i == 1 ? MergedCellValues.Restart : MergedCellValues.Continue;
                        tableCellProperties.AppendChild(horizontalMerge);
                        if (i == 1)
                            SetCellText(cell, "Lunch", true);
                        else
                            SetCellText(cell, "", false);
                    }
                    else if (x == 2 || x == 5) switch (i)
                        {
                            case 1:
                            case 3:
                                SetCellText(cell, "Science", false);
                                break;
                            case 2:
                            case 4:
                                SetCellText(cell, "Maths", false);
                                break;
                        }
                    else if (x == 3 || x == 6) switch (i)
                        {
                            case 1:
                            case 4:
                                SetCellText(cell, "Social", false);
                                break;
                            case 2:
                                SetCellText(cell, "History", false);
                                break;
                            case 3:
                                SetCellText(cell, "English", false);
                                break;
                        }
                    if (x == 2 && i == 5)
                    {
                        SetCellText(cell, "Arts", false);
                    }
                    else if (x == 3 && i == 5)
                    {
                        SetCellText(cell, "Sports", false);
                    }
                    else if (i == 5 && (x == 5 || x == 6))
                    {
                        if (x == 5)
                        {
                            verticalMerge.Val = MergedCellValues.Restart;
                            SetCellText(cell, "Project", false);
                        }
                        else
                        {
                            verticalMerge.Val = MergedCellValues.Continue;
                            SetCellText(cell, "", false);
                        }
                        tableCellProperties.AppendChild(verticalMerge);
                    }
                    cell.AppendChild(tableCellProperties);
                    cells.Add(cell);
                }
                pesonalTableRow.Append(cells);
                personalInfoTable.AppendChild(pesonalTableRow);
            }
            body.AppendChild(personalInfoTable);
        }

        private void SetCellText(TableCell cell, string text, bool bold)
        {
            ParagraphProperties parProperties = new ParagraphProperties();
            parProperties.Justification = new Justification();
            parProperties.Justification.Val = JustificationValues.Center;
            Paragraph paragraph = new Paragraph();
            paragraph.AppendChild(parProperties);
            var run = new Run(new Text(text));
            run.RunProperties = new RunProperties();
            run.RunProperties.FontSize = new FontSize();
            run.RunProperties.FontSize.Val = new StringValue("28");
            if (bold)
            {
                run.RunProperties.Bold = new Bold();
            }
            paragraph.AppendChild(run);
            cell.AppendChild(paragraph);
        }

        private void AddTableBorders(TableProperties tableProperties)
        {
            TableBorders tblBorders = new TableBorders();
            AddBorderValues(new TopBorder(), tblBorders);
            AddBorderValues(new BottomBorder(), tblBorders);
            AddBorderValues(new RightBorder(), tblBorders);
            AddBorderValues(new LeftBorder(), tblBorders);
            AddBorderValues(new InsideHorizontalBorder(), tblBorders);
            AddBorderValues(new InsideVerticalBorder(), tblBorders);
            tableProperties.AppendChild(tblBorders);
        }

        private void AddBorderValues(BorderType type, TableBorders borders, BorderValues value = BorderValues.Single,
            string color = "00000")
        {
            type.Val = new EnumValue<BorderValues>(BorderValues.Single);
            type.Color = "00000";
            borders.AppendChild(type);
        }

        public void Transform(string firstName, string lastName, string city)
        {
            using (WordprocessingDocument wordprocessingDocument =
                   WordprocessingDocument.Open(LastConvertedWordFileName ?? LastWordFileName, true))
            {
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                var firstParagraph = body.ChildElements.FirstOrDefault(c => c is Paragraph);
                if (firstParagraph != null)
                {
                    AddPersonalRow(body, firstParagraph, firstName, lastName, city);
                }
                ApplyColorsFormat(body);
                Paragraph breakParagraph = new Paragraph();
                ParagraphProperties breakProperties = new ParagraphProperties();
                SectionProperties sectionProperties = new SectionProperties();
                SectionType sectionType = new SectionType() { Val = SectionMarkValues.NextPage };
                sectionProperties.AppendChild(sectionType);
                breakProperties.AppendChild(sectionProperties);
                breakParagraph.AppendChild(breakProperties);
                body.AppendChild(breakParagraph);
                AddSecondPageTable(body);
            }
        }

        public void ConvertWordDocument(string fullFileName, WdSaveFormat format)
        {
            Word.Application wordApp = new Word.Application();
            try
            {
                string wordFileNameNoExt = fullFileName.Remove(fullFileName.LastIndexOf("."));
                wordApp.Visible = false;
                wordApp.WindowState = WdWindowState.wdWindowStateMinimize;
                wordApp.Documents.Open(fullFileName);
                bool oldWord = (WdSaveFormat)wordApp.ActiveDocument.SaveFormat == WdSaveFormat.wdFormatDocument97;
                if (oldWord)
                {
                    LastWordFileName = fullFileName;
                }
                else if (format == WdSaveFormat.wdFormatDocumentDefault)
                {
                    RemoveLastDocxFile();
                    LastWordFileName = fullFileName;
                    return;
                }
                string fName = wordFileNameNoExt;
                int lastSlashIndex = fullFileName.LastIndexOf("\\");
                string shortFileNameNoExt = wordFileNameNoExt.Substring(lastSlashIndex);
                switch (format)
                {
                    case WdSaveFormat.wdFormatDocumentDefault:
                        RemoveLastDocxFile();
                        fName = Environment.CurrentDirectory + _tempDir + shortFileNameNoExt + ".docx";
                        LastConvertedWordFileName = fName;
                        break;
                    case WdSaveFormat.wdFormatDocument97:
                        fName = LastWordFileName;
                        break;
                    case WdSaveFormat.wdFormatXPS:
                        RemoveLastXpsFile();
                        fName = Environment.CurrentDirectory + _tempDir + shortFileNameNoExt + ".xps";
                        LastXpsFileName = fName;
                        break;
                    default:
                        return;
                }
                wordApp.ActiveDocument.SaveAs(fName, format);
                if (format == WdSaveFormat.wdFormatXPS)
                    XpsDocument = new XpsDocument(LastXpsFileName, FileAccess.Read);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed to prepare MS Office Word document",
                        MessageBoxButton.OK, MessageBoxImage.Stop);
            }
            finally
            {
                wordApp.Documents.Close();
                wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges);
                Marshal.ReleaseComObject(wordApp.Documents);
                Marshal.ReleaseComObject(wordApp);
                wordApp = null;
            }
        }

        public void RemoveLastXpsFile()
        {
            if (XpsDocument != null)
            {
                XpsDocument.Close();
                XpsDocument = null;
            }
            if (!string.IsNullOrEmpty(LastXpsFileName) && File.Exists(LastXpsFileName))
            {
                File.Delete(LastXpsFileName);
            }
            LastXpsFileName = null;
        }

        public void RemoveLastDocxFile()
        {
            if (!string.IsNullOrEmpty(LastConvertedWordFileName) && File.Exists(LastConvertedWordFileName))
            {
                File.Delete(LastConvertedWordFileName);
            }
            LastConvertedWordFileName = null;
        }
    }
}
