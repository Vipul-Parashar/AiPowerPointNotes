using System;
using System.Text;
using AiSpeakerNotes.Models;

namespace AiSpeakerNotes.Services
{
    public static class PromptBuilder
    {
        public static string BuildSystemPrompt(GenerationSettings settings)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are an expert, friendly teacher writing spoken presentation notes for yourself to deliver to your class.");
            sb.AppendLine();
            sb.AppendLine("## CORE RULES & STYLE:");
            sb.AppendLine("- Tone & Language: Speak in plain, clear, conversational English as if talking directly to students who are completely new to the topic.");

            if (string.Equals(settings.LanguageLevel, "Very Simple", StringComparison.OrdinalIgnoreCase))
            {
                sb.AppendLine("- Simplicity: Use very simple everyday words and concepts (suitable for middle school or total beginners). Avoid academic phrasing.");
            }
            else
            {
                sb.AppendLine("- Simplicity: Keep explanations accessible, engaging, and clear. Avoid pretentious vocabulary.");
            }

            sb.AppendLine("- Sentence Structure: Use short, punchy sentences that are natural to say out loud while standing in front of a class. No long run-ons.");
            sb.AppendLine("- Jargon Policy: Strictly avoid unnecessary jargon. If a technical term is essential, define or explain it immediately in one simple line.");
            sb.AppendLine(string.Format("- Length: Output between {0} and {1} sentences total for this slide. Do not write essays.",
                settings.GetTargetSentenceMin(), settings.GetTargetSentenceMax()));
            sb.AppendLine("- Spoken Word: Write exactly what the teacher should say out loud (spoken dialogue). Do not include stage directions, bullet symbols, timestamps, or headers like 'Slide 1: Notes:'.");
            sb.AppendLine("- Narrative Transitions: Start each slide's notes with a natural, conversational bridge connecting it to the previous slide's idea (e.g., 'Now that we understand X, let's look at how Y works...', 'Building on that idea...'). If this is the very first slide, start with a warm welcome and hook.");
            sb.AppendLine("- Real-Life Analogy: When explaining abstract or theoretical concepts, include ONE brief, intuitive real-life example or everyday analogy.");
            sb.AppendLine("- Explain, Don't Repeat: Never just read off or recite bullet points or table cells. Explain the reasoning behind them, why they matter, and what students should take away.");
            sb.AppendLine(string.Format("- Audience / Subject: Tailor the examples and explanations for: {0}.", settings.Audience));
            sb.AppendLine();
            sb.AppendLine("## OUTPUT FORMAT:");
            sb.AppendLine("Return ONLY the plain text of the spoken notes. No quotes around the whole text, no Markdown bolding, no prefixes.");
            return sb.ToString();
        }

        public static string BuildUserPrompt(SlideData slide, GenerationSettings settings)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Format("Please generate the spoken speaker notes for Slide {0}.", slide.SlideIndex));
            sb.AppendLine();

            if (!string.IsNullOrEmpty(slide.PrevSlideSummary))
            {
                sb.AppendLine("### Context from Previous Slide:");
                sb.AppendLine(slide.PrevSlideSummary);
                sb.AppendLine();
            }

            sb.AppendLine("### Current Slide Content:");
            if (!string.IsNullOrWhiteSpace(slide.Title))
            {
                sb.AppendLine("Title: " + slide.Title);
            }
            else
            {
                sb.AppendLine("Title: (No slide title)");
            }

            if (slide.BulletPoints.Count > 0)
            {
                sb.AppendLine("Bullet Points / Key Items:");
                foreach (var bp in slide.BulletPoints)
                {
                    sb.AppendLine("- " + bp);
                }
            }

            if (slide.TextElements.Count > 0)
            {
                sb.AppendLine("Additional Text Elements:");
                foreach (var t in slide.TextElements)
                {
                    sb.AppendLine("• " + t);
                }
            }

            if (slide.Tables.Count > 0)
            {
                sb.AppendLine("Table Data:");
                foreach (var tbl in slide.Tables)
                {
                    sb.AppendLine(tbl);
                }
            }

            if (slide.ImageAltTexts.Count > 0)
            {
                sb.AppendLine("Images / Diagrams (Alt Text & Descriptions):");
                foreach (var alt in slide.ImageAltTexts)
                {
                    sb.AppendLine("- " + alt);
                }
            }

            if (slide.IsImageOnly)
            {
                sb.AppendLine();
                sb.AppendLine("[Note: This slide contains only visual images/diagrams without body text. Please explain the visual concept based on its alt text and neighboring slide context in 3-4 clear spoken sentences.]");
            }

            if (!string.IsNullOrEmpty(slide.NextSlideSummary))
            {
                sb.AppendLine();
                sb.AppendLine("### Context from Next Slide:");
                sb.AppendLine(slide.NextSlideSummary);
            }

            sb.AppendLine();
            sb.AppendLine("Now provide the spoken speaker notes following all style rules:");
            return sb.ToString();
        }
    }
}
