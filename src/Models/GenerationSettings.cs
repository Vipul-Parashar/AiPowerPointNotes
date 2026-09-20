using System;

namespace AiSpeakerNotes.Models
{
    public class GenerationSettings
    {
        public LlmProvider Provider { get; set; }
        public string ApiKey { get; set; }
        public string Model { get; set; }

        public string LanguageLevel { get; set; }
        public string Length { get; set; }
        public string Audience { get; set; }

        public bool OverwriteExisting { get; set; }
        public bool ScopeAllSlides { get; set; }

        public GenerationSettings()
        {
            Provider = LlmProvider.Gemini;
            ApiKey = string.Empty;
            Model = ProviderDefaults.DefaultGeminiModel;
            LanguageLevel = "Normal";
            Length = "Medium (5-6 sentences)";
            Audience = "college students, computer science";
            OverwriteExisting = false;
            ScopeAllSlides = true;
        }

        public int GetTargetSentenceMin()
        {
            if (Length != null && Length.IndexOf("Short", StringComparison.OrdinalIgnoreCase) >= 0) return 3;
            if (Length != null && Length.IndexOf("Detailed", StringComparison.OrdinalIgnoreCase) >= 0) return 7;
            return 5;
        }

        public int GetTargetSentenceMax()
        {
            if (Length != null && Length.IndexOf("Short", StringComparison.OrdinalIgnoreCase) >= 0) return 4;
            if (Length != null && Length.IndexOf("Detailed", StringComparison.OrdinalIgnoreCase) >= 0) return 8;
            return 6;
        }
    }
}
