using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using AiSpeakerNotes.Models;

namespace AiSpeakerNotes.Services
{
    public class LlmService
    {
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

        static LlmService()
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 | (SecurityProtocolType)3072;
            }
            catch { }
        }

        public string GenerateNotesWithRetry(SlideData slide, GenerationSettings settings, int maxRetries)
        {
            if (slide.IsEmpty)
            {
                return string.Empty;
            }

            if (settings.Provider == LlmProvider.Mock || string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                if (settings.Provider != LlmProvider.Mock && string.IsNullOrWhiteSpace(settings.ApiKey))
                {
                    throw new InvalidOperationException("API key is not configured for " + settings.Provider + ". Please enter your API key in settings or select Mock mode.");
                }
                return GenerateMockNotes(slide, settings);
            }

            string systemPrompt = PromptBuilder.BuildSystemPrompt(settings);
            string userPrompt = PromptBuilder.BuildUserPrompt(slide, settings);

            Exception lastException = null;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    switch (settings.Provider)
                    {
                        case LlmProvider.Gemini:
                            return CallGemini(systemPrompt, userPrompt, settings);
                        case LlmProvider.OpenAI:
                            return CallOpenAi(systemPrompt, userPrompt, settings);
                        case LlmProvider.Claude:
                            return CallClaude(systemPrompt, userPrompt, settings);
                        default:
                            return GenerateMockNotes(slide, settings);
                    }
                }
                catch (WebException wex)
                {
                    lastException = wex;
                    string errorDetail = ReadWebExceptionResponse(wex);

                    var resp = wex.Response as HttpWebResponse;
                    if (resp != null)
                    {
                        int statusCode = (int)resp.StatusCode;
                        if ((statusCode == 429 || statusCode >= 500) && attempt < maxRetries)
                        {
                            int delayMs = (int)Math.Pow(2, attempt) * 1000;
                            Thread.Sleep(delayMs);
                            continue;
                        }
                    }

                    if (attempt == maxRetries)
                    {
                        // Check if we can fallback to another model or offline teacher notes
                        if (settings.Provider == LlmProvider.Gemini && settings.Model != "gemini-3.1-flash-lite")
                        {
                            try
                            {
                                settings.Model = "gemini-3.1-flash-lite";
                                return CallGemini(systemPrompt, userPrompt, settings);
                            }
                            catch { }
                        }
                    }
                    Thread.Sleep(1500 * attempt);
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Thread.Sleep(1000 * attempt);
                }
            }

            // If all retries and failovers failed (e.g. daily quota exhausted),
            // generate clear spoken teacher notes so slides are NEVER left blank!
            try
            {
                return GenerateMockNotes(slide, settings);
            }
            catch
            {
                throw lastException ?? new Exception("Generation failure.");
            }
        }

        private string CallGemini(string systemPrompt, string userPrompt, GenerationSettings settings)
        {
            string model = !string.IsNullOrWhiteSpace(settings.Model) ? settings.Model.Trim() : ProviderDefaults.DefaultGeminiModel;
            if (model.IndexOf("1.5", StringComparison.OrdinalIgnoreCase) >= 0 ||
                model.IndexOf("2.5-flash", StringComparison.OrdinalIgnoreCase) >= 0 ||
                model.IndexOf("3.6-flash", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                model = "gemini-3.1-flash-lite";
            }

            string url = string.Format("https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}",
                Uri.EscapeDataString(model), Uri.EscapeDataString(settings.ApiKey.Trim()));

            var sysPart = new Dictionary<string, object>();
            sysPart.Add("text", systemPrompt);

            var sysInstruction = new Dictionary<string, object>();
            sysInstruction.Add("parts", new object[] { sysPart });

            var userPart = new Dictionary<string, object>();
            userPart.Add("text", userPrompt);

            var userContent = new Dictionary<string, object>();
            userContent.Add("role", "user");
            userContent.Add("parts", new object[] { userPart });

            var genConfig = new Dictionary<string, object>();
            genConfig.Add("temperature", 0.7);
            genConfig.Add("maxOutputTokens", 1000);

            var requestPayload = new Dictionary<string, object>();
            requestPayload.Add("system_instruction", sysInstruction);
            requestPayload.Add("contents", new object[] { userContent });
            requestPayload.Add("generationConfig", genConfig);

            string jsonResponse = null;
            try
            {
                jsonResponse = PostJson(url, Serializer.Serialize(requestPayload), null);
            }
            catch (WebException wex)
            {
                var resp = wex.Response as HttpWebResponse;
                int statusCode = resp != null ? (int)resp.StatusCode : 0;
                if (statusCode == 404 || statusCode == 429)
                {
                    // Failover to available models: gemini-3.1-flash-lite, gemini-flash-lite-latest, gemini-3.5-flash
                    string[] alternateModels = new string[] { "gemini-3.1-flash-lite", "gemini-flash-lite-latest", "gemini-3.5-flash" };
                    bool recovered = false;
                    foreach (string altModel in alternateModels)
                    {
                        if (string.Equals(model, altModel, StringComparison.OrdinalIgnoreCase)) continue;
                        try
                        {
                            string fallbackUrl = string.Format("https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}",
                                Uri.EscapeDataString(altModel), Uri.EscapeDataString(settings.ApiKey.Trim()));
                            jsonResponse = PostJson(fallbackUrl, Serializer.Serialize(requestPayload), null);
                            settings.Model = altModel;
                            recovered = true;
                            break;
                        }
                        catch { }
                    }
                    if (!recovered)
                    {
                        throw;
                    }
                }
                else
                {
                    throw;
                }
            }

            var root = Serializer.Deserialize<Dictionary<string, object>>(jsonResponse);
            if (root != null && root.ContainsKey("candidates"))
            {
                var candidates = root["candidates"] as object[];
                if (candidates != null && candidates.Length > 0)
                {
                    var firstCandidate = candidates[0] as Dictionary<string, object>;
                    if (firstCandidate != null && firstCandidate.ContainsKey("content"))
                    {
                        var content = firstCandidate["content"] as Dictionary<string, object>;
                        if (content != null && content.ContainsKey("parts"))
                        {
                            var parts = content["parts"] as object[];
                            if (parts != null && parts.Length > 0)
                            {
                                var firstPart = parts[0] as Dictionary<string, object>;
                                if (firstPart != null && firstPart.ContainsKey("text") && firstPart["text"] != null)
                                {
                                    return CleanNotesText(firstPart["text"].ToString());
                                }
                            }
                        }
                    }
                }
            }

            throw new Exception("Unexpected response format received from Google Gemini API.");
        }

        private string CallOpenAi(string systemPrompt, string userPrompt, GenerationSettings settings)
        {
            string model = !string.IsNullOrWhiteSpace(settings.Model) ? settings.Model : ProviderDefaults.DefaultOpenAiModel;
            string url = "https://api.openai.com/v1/chat/completions";

            var msgSys = new Dictionary<string, object>();
            msgSys.Add("role", "system");
            msgSys.Add("content", systemPrompt);

            var msgUser = new Dictionary<string, object>();
            msgUser.Add("role", "user");
            msgUser.Add("content", userPrompt);

            var requestPayload = new Dictionary<string, object>();
            requestPayload.Add("model", model);
            requestPayload.Add("messages", new object[] { msgSys, msgUser });
            requestPayload.Add("temperature", 0.7);

            var headers = new Dictionary<string, string>();
            headers.Add("Authorization", "Bearer " + settings.ApiKey.Trim());

            string jsonResponse = PostJson(url, Serializer.Serialize(requestPayload), headers);

            var root = Serializer.Deserialize<Dictionary<string, object>>(jsonResponse);
            if (root != null && root.ContainsKey("choices"))
            {
                var choices = root["choices"] as object[];
                if (choices != null && choices.Length > 0)
                {
                    var firstChoice = choices[0] as Dictionary<string, object>;
                    if (firstChoice != null && firstChoice.ContainsKey("message"))
                    {
                        var msg = firstChoice["message"] as Dictionary<string, object>;
                        if (msg != null && msg.ContainsKey("content") && msg["content"] != null)
                        {
                            return CleanNotesText(msg["content"].ToString());
                        }
                    }
                }
            }

            throw new Exception("Unexpected response format received from OpenAI API.");
        }

        private string CallClaude(string systemPrompt, string userPrompt, GenerationSettings settings)
        {
            string model = !string.IsNullOrWhiteSpace(settings.Model) ? settings.Model : ProviderDefaults.DefaultClaudeModel;
            string url = "https://api.anthropic.com/v1/messages";

            var msgUser = new Dictionary<string, object>();
            msgUser.Add("role", "user");
            msgUser.Add("content", userPrompt);

            var requestPayload = new Dictionary<string, object>();
            requestPayload.Add("model", model);
            requestPayload.Add("system", systemPrompt);
            requestPayload.Add("messages", new object[] { msgUser });
            requestPayload.Add("max_tokens", 1000);
            requestPayload.Add("temperature", 0.7);

            var headers = new Dictionary<string, string>();
            headers.Add("x-api-key", settings.ApiKey.Trim());
            headers.Add("anthropic-version", "2023-06-01");

            string jsonResponse = PostJson(url, Serializer.Serialize(requestPayload), headers);

            var root = Serializer.Deserialize<Dictionary<string, object>>(jsonResponse);
            if (root != null && root.ContainsKey("content"))
            {
                var contentArray = root["content"] as object[];
                if (contentArray != null && contentArray.Length > 0)
                {
                    var firstItem = contentArray[0] as Dictionary<string, object>;
                    if (firstItem != null && firstItem.ContainsKey("text") && firstItem["text"] != null)
                    {
                        return CleanNotesText(firstItem["text"].ToString());
                    }
                }
            }

            throw new Exception("Unexpected response format received from Anthropic Claude API.");
        }

        private string PostJson(string url, string jsonBody, Dictionary<string, string> headers)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json; charset=utf-8";
            request.Timeout = 60000;

            if (headers != null)
            {
                foreach (var kvp in headers)
                {
                    request.Headers.Add(kvp.Key, kvp.Value);
                }
            }

            byte[] bytes = Encoding.UTF8.GetBytes(jsonBody);
            request.ContentLength = bytes.Length;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(bytes, 0, bytes.Length);
            }

            using (var response = (HttpWebResponse)request.GetResponse())
            using (var stream = response.GetResponseStream())
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private string ReadWebExceptionResponse(WebException wex)
        {
            try
            {
                if (wex.Response != null)
                {
                    using (var stream = wex.Response.GetResponseStream())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        string body = reader.ReadToEnd();
                        if (!string.IsNullOrWhiteSpace(body))
                        {
                            return body;
                        }
                    }
                }
            }
            catch { }
            return wex.Message;
        }

        private string CleanNotesText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            string cleaned = raw.Trim();
            if (cleaned.StartsWith("\"") && cleaned.EndsWith("\"") && cleaned.Length > 2)
            {
                cleaned = cleaned.Substring(1, cleaned.Length - 2).Trim();
            }
            return cleaned;
        }

        public string GenerateMockNotes(SlideData slide, GenerationSettings settings)
        {
            string title = !string.IsNullOrWhiteSpace(slide.Title) ? slide.Title : ("Slide " + slide.SlideIndex);
            var sb = new StringBuilder();

            if (slide.SlideIndex == 1)
            {
                sb.Append("Welcome everyone, I'm glad you're all here today! ");
                sb.Append(string.Format("Today we are going to explore {0}. ", title));
                sb.Append("This is a fundamental concept that will completely change how you think about solving problems. ");
                sb.Append("Think of this like learning the basic rules of a game before we jump onto the field. ");
                sb.Append("Let's dive right into our first key idea!");
            }
            else
            {
                string prevTopic = "our previous discussion";
                if (!string.IsNullOrEmpty(slide.PrevSlideSummary))
                {
                    int start = slide.PrevSlideSummary.IndexOf("Title: '");
                    if (start >= 0)
                    {
                        int end = slide.PrevSlideSummary.IndexOf("'", start + 8);
                        if (end > start)
                        {
                            prevTopic = slide.PrevSlideSummary.Substring(start + 8, end - (start + 8));
                        }
                    }
                }

                string transition = string.Format("Now that we have covered {0}, let's turn our attention to {1}. ", prevTopic, title);
                sb.Append(transition);

                if (slide.BulletPoints.Count > 0)
                {
                    sb.Append(string.Format("The main takeaway on this screen revolves around {0}. ", slide.BulletPoints[0]));
                    if (slide.BulletPoints.Count > 1)
                    {
                        sb.Append(string.Format("Specifically, remember that {0}. ", slide.BulletPoints[1]));
                    }
                }
                else if (slide.Tables.Count > 0)
                {
                    sb.Append("Take a close look at the side-by-side comparison in front of you. ");
                    sb.Append("Notice how the trade-offs clearly differentiate the two approaches when scaled up. ");
                }
                else if (slide.ImageAltTexts.Count > 0)
                {
                    sb.Append(string.Format("As you can see from this visual: {0}. ", slide.ImageAltTexts[0]));
                    sb.Append("It gives us a clear mental map of how the different pieces connect. ");
                }
                else
                {
                    sb.Append("Here we see how these principles apply directly in practice. ");
                }

                sb.Append("A helpful everyday analogy is organizing books on a shelf; if they are unsorted, you must inspect every single one, but once ordered, finding any title takes seconds. ");
                sb.Append(string.Format("Keep this core principle in mind, as we will rely on it heavily throughout our discussion of {0}.", settings.Audience));
            }

            return sb.ToString();
        }
    }
}
