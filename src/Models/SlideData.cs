using System.Collections.Generic;

namespace AiSpeakerNotes.Models
{
    public class SlideData
    {
        public int SlideIndex { get; set; }
        public int SlideId { get; set; }
        public string Title { get; set; }
        public List<string> TextElements { get; set; }
        public List<string> BulletPoints { get; set; }
        public List<string> Tables { get; set; }
        public List<string> ImageAltTexts { get; set; }

        public string ExistingNotes { get; set; }

        public bool HasNotes
        {
            get { return !string.IsNullOrWhiteSpace(ExistingNotes); }
        }

        public string PrevSlideSummary { get; set; }
        public string NextSlideSummary { get; set; }

        public bool IsEmpty
        {
            get
            {
                return string.IsNullOrWhiteSpace(Title) &&
                       (TextElements == null || TextElements.Count == 0) &&
                       (BulletPoints == null || BulletPoints.Count == 0) &&
                       (Tables == null || Tables.Count == 0) &&
                       (ImageAltTexts == null || ImageAltTexts.Count == 0);
            }
        }

        public bool IsImageOnly
        {
            get
            {
                return string.IsNullOrWhiteSpace(Title) &&
                       (TextElements == null || TextElements.Count == 0) &&
                       (BulletPoints == null || BulletPoints.Count == 0) &&
                       (Tables == null || Tables.Count == 0) &&
                       (ImageAltTexts != null && ImageAltTexts.Count > 0);
            }
        }

        public SlideData()
        {
            Title = string.Empty;
            TextElements = new List<string>();
            BulletPoints = new List<string>();
            Tables = new List<string>();
            ImageAltTexts = new List<string>();
            ExistingNotes = string.Empty;
            PrevSlideSummary = string.Empty;
            NextSlideSummary = string.Empty;
        }
    }
}
