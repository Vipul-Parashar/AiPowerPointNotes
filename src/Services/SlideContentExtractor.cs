using System;
using System.Collections.Generic;
using System.Text;
using AiSpeakerNotes.Models;

namespace AiSpeakerNotes.Services
{
    public static class SlideContentExtractor
    {
        public static SlideData ExtractSlide(object slideObj)
        {
            if (slideObj == null) return new SlideData();
            dynamic slide = slideObj;

            var data = new SlideData();
            try { data.SlideIndex = (int)slide.SlideIndex; } catch { }
            try { data.SlideId = (int)slide.SlideID; } catch { }

            // 1. Read existing notes
            data.ExistingNotes = NotesWriter.GetNotes(slideObj);

            // 2. Identify title
            int titleShapeId = -1;
            try
            {
                if ((int)slide.Shapes.HasTitle == -1) // msoTrue
                {
                    dynamic titleShape = slide.Shapes.Title;
                    if (titleShape != null && (int)titleShape.HasTextFrame == -1)
                    {
                        titleShapeId = (int)titleShape.Id;
                        string titleText = titleShape.TextFrame.TextRange.Text;
                        if (titleText != null)
                        {
                            data.Title = titleText.Trim();
                        }
                    }
                }
            }
            catch { }

            // 3. Iterate shapes
            try
            {
                foreach (dynamic shape in slide.Shapes)
                {
                    try
                    {
                        int shapeId = (int)shape.Id;
                        if (shapeId == titleShapeId) continue;

                        // Check table
                        int hasTable = 0;
                        try { hasTable = (int)shape.HasTable; } catch { }
                        if (hasTable == -1) // msoTrue
                        {
                            string tableStr = ExtractTableContent(shape.Table);
                            if (!string.IsNullOrWhiteSpace(tableStr))
                            {
                                data.Tables.Add(tableStr);
                            }
                            continue;
                        }

                        // Check TextFrame
                        int hasTextFrame = 0;
                        try { hasTextFrame = (int)shape.HasTextFrame; } catch { }
                        if (hasTextFrame == -1)
                        {
                            int hasText = 0;
                            try { hasText = (int)shape.TextFrame.HasText; } catch { }
                            if (hasText == -1)
                            {
                                dynamic textRange = shape.TextFrame.TextRange;
                                int paragraphCount = (int)textRange.Paragraphs().Count;

                                for (int p = 1; p <= paragraphCount; p++)
                                {
                                    dynamic para = textRange.Paragraphs(p);
                                    string text = para.Text != null ? para.Text.ToString().Trim() : string.Empty;
                                    if (string.IsNullOrWhiteSpace(text)) continue;

                                    // Fallback title detection: placeholder title, or top-most large font text
                                    if (string.IsNullOrWhiteSpace(data.Title))
                                    {
                                        int shapeType = 0;
                                        try { shapeType = (int)shape.Type; } catch { }
                                        if (shapeType == 14) // msoPlaceholder
                                        {
                                            int placeholderType = (int)shape.PlaceholderFormat.Type;
                                            if (placeholderType == 1 || placeholderType == 3) // ppPlaceholderTitle or ppPlaceholderCenterTitle
                                            {
                                                data.Title = text;
                                                continue;
                                            }
                                        }

                                        float shapeTop = 1000;
                                        try { shapeTop = (float)shape.Top; } catch { }
                                        float fontSize = 0;
                                        try { fontSize = (float)para.Font.Size; } catch { }
                                        if (shapeTop < 120 && (fontSize >= 20 || p == 1))
                                        {
                                            data.Title = text;
                                            continue;
                                        }
                                    }

                                    // Bullet or normal text
                                    string[] lines = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (string line in lines)
                                    {
                                        string trimmedLine = line.Trim();
                                        if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                                        int bulletType = 0;
                                        try { bulletType = (int)para.ParagraphFormat.Bullet.Type; } catch { }
                                        if (bulletType != 0) // not ppBulletNone
                                        {
                                            data.BulletPoints.Add(trimmedLine);
                                        }
                                        else
                                        {
                                            data.TextElements.Add(trimmedLine);
                                        }
                                    }
                                }
                            }
                        }

                        // Check Alt Text (Images, Diagrams, Charts)
                        string altText = string.Empty;
                        try { altText = shape.AlternativeText; } catch { }
                        string shapeTitle = string.Empty;
                        try { shapeTitle = shape.Title; } catch { }

                        altText = altText != null ? altText.Trim() : string.Empty;
                        shapeTitle = shapeTitle != null ? shapeTitle.Trim() : string.Empty;

                        if (!string.IsNullOrWhiteSpace(altText))
                        {
                            string desc = !string.IsNullOrWhiteSpace(shapeTitle)
                                ? string.Format("{0}: {1}", shapeTitle, altText)
                                : altText;
                            data.ImageAltTexts.Add(desc);
                        }
                        else if (!string.IsNullOrWhiteSpace(shapeTitle))
                        {
                            data.ImageAltTexts.Add(shapeTitle);
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return data;
        }

        private static string ExtractTableContent(dynamic table)
        {
            var sb = new StringBuilder();
            try
            {
                int rows = (int)table.Rows.Count;
                int cols = (int)table.Columns.Count;

                for (int r = 1; r <= rows; r++)
                {
                    var rowCells = new List<string>();
                    for (int c = 1; c <= cols; c++)
                    {
                        string cellText = string.Empty;
                        try
                        {
                            dynamic cell = table.Cell(r, c);
                            if (cell != null && (int)cell.Shape.HasTextFrame == -1)
                            {
                                string t = cell.Shape.TextFrame.TextRange.Text;
                                if (t != null) cellText = t.Trim();
                            }
                        }
                        catch { }
                        rowCells.Add(cellText.Replace("\r", " ").Replace("\n", " "));
                    }
                    sb.AppendLine("| " + string.Join(" | ", rowCells.ToArray()) + " |");
                }
            }
            catch { }
            return sb.ToString().TrimEnd();
        }

        public static string SummarizeSlide(SlideData slide)
        {
            if (slide == null || slide.IsEmpty) return "Empty slide";

            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(slide.Title))
            {
                parts.Add("Title: '" + slide.Title + "'");
            }
            if (slide.BulletPoints != null && slide.BulletPoints.Count > 0)
            {
                var summaryPoints = new List<string>();
                for (int i = 0; i < Math.Min(2, slide.BulletPoints.Count); i++)
                {
                    string pt = slide.BulletPoints[i];
                    if (pt.Length > 80) pt = pt.Substring(0, 77) + "...";
                    summaryPoints.Add(pt);
                }
                parts.Add("Key points: " + string.Join("; ", summaryPoints.ToArray()));
            }
            else if (slide.TextElements != null && slide.TextElements.Count > 0)
            {
                string txt = slide.TextElements[0];
                if (txt.Length > 80) txt = txt.Substring(0, 77) + "...";
                parts.Add("Text: " + txt);
            }
            else if (slide.ImageAltTexts != null && slide.ImageAltTexts.Count > 0)
            {
                string alt = slide.ImageAltTexts[0];
                if (alt.Length > 80) alt = alt.Substring(0, 77) + "...";
                parts.Add("Visual: " + alt);
            }
            return string.Join(" - ", parts.ToArray());
        }
    }
}
