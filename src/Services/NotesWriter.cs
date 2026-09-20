using System;

namespace AiSpeakerNotes.Services
{
    public static class NotesWriter
    {
        public static string GetNotes(object slideObj)
        {
            try
            {
                if (slideObj == null) return string.Empty;
                dynamic slide = slideObj;
                dynamic notesPage = slide.NotesPage;
                if (notesPage == null) return string.Empty;

                foreach (dynamic shape in notesPage.Shapes)
                {
                    try
                    {
                        // Check if shape has text frame and placeholder type == 2 (ppPlaceholderBody)
                        if (shape.HasTextFrame == -1) // msoTrue
                        {
                            int shapeType = (int)shape.Type;
                            // msoPlaceholder is 14
                            if (shapeType == 14)
                            {
                                int placeholderType = (int)shape.PlaceholderFormat.Type;
                                // ppPlaceholderBody is 2
                                if (placeholderType == 2)
                                {
                                    string text = shape.TextFrame.TextRange.Text;
                                    return text != null ? text.Trim() : string.Empty;
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error reading notes: " + ex.Message);
            }
            return string.Empty;
        }

        public static void SetNotes(object slideObj, string notesText)
        {
            if (slideObj == null) return;

            dynamic slide = slideObj;
            string textToSet = notesText != null ? notesText.Trim() : string.Empty;

            try
            {
                dynamic notesPage = slide.NotesPage;

                foreach (dynamic shape in notesPage.Shapes)
                {
                    try
                    {
                        if (shape.HasTextFrame == -1)
                        {
                            int shapeType = (int)shape.Type;
                            if (shapeType == 14)
                            {
                                int placeholderType = (int)shape.PlaceholderFormat.Type;
                                if (placeholderType == 2)
                                {
                                    shape.TextFrame.TextRange.Text = textToSet;
                                    return;
                                }
                            }
                        }
                    }
                    catch { }
                }

                // Fallback: If body placeholder was not found, add a textbox to NotesPage
                dynamic newShape = notesPage.Shapes.AddTextbox(1, 54, 300, 432, 300); // 1 = msoTextOrientationHorizontal
                newShape.TextFrame.TextRange.Text = textToSet;
            }
            catch (Exception ex)
            {
                int idx = 0;
                try { idx = (int)slide.SlideIndex; } catch { }
                throw new Exception(string.Format("Failed to write speaker notes to Slide {0}: {1}", idx, ex.Message), ex);
            }
        }
    }
}
