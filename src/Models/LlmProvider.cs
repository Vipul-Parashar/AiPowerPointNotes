using System;

namespace AiSpeakerNotes.Models
{
    public enum LlmProvider
    {
        Gemini,
        OpenAI,
        Claude,
        Mock
    }

    public static class ProviderDefaults
    {
        public const string DefaultGeminiModel = "gemini-3.1-flash-lite";
        public const string DefaultOpenAiModel = "gpt-4o-mini";
        public const string DefaultClaudeModel = "claude-3-5-sonnet-20241022";
        public const string DefaultMockModel = "offline-mock-teacher";

        public static string GetDefaultModel(LlmProvider provider)
        {
            switch (provider)
            {
                case LlmProvider.Gemini:
                    return DefaultGeminiModel;
                case LlmProvider.OpenAI:
                    return DefaultOpenAiModel;
                case LlmProvider.Claude:
                    return DefaultClaudeModel;
                case LlmProvider.Mock:
                    return DefaultMockModel;
                default:
                    return DefaultGeminiModel;
            }
        }
    }
}
