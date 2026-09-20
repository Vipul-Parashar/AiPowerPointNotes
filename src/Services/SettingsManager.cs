using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using AiSpeakerNotes.Models;

namespace AiSpeakerNotes.Services
{
    public class SettingsDataTransfer
    {
        public int Provider { get; set; }
        public string EncryptedApiKey { get; set; }
        public string Model { get; set; }
        public string LanguageLevel { get; set; }
        public string Length { get; set; }
        public string Audience { get; set; }
        public bool OverwriteExisting { get; set; }
    }

    public static class SettingsManager
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AiSpeakerNotes");

        private static readonly string SettingsFilePath = Path.Combine(SettingsDir, "settings.json");

        public static GenerationSettings Load()
        {
            var settings = new GenerationSettings();
            try
            {
                if (!File.Exists(SettingsFilePath))
                    return settings;

                string json = File.ReadAllText(SettingsFilePath, Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                var dto = serializer.Deserialize<SettingsDataTransfer>(json);
                if (dto == null) return settings;

                settings.Provider = (LlmProvider)dto.Provider;
                settings.Model = !string.IsNullOrEmpty(dto.Model) ? dto.Model : ProviderDefaults.GetDefaultModel(settings.Provider);
                settings.LanguageLevel = !string.IsNullOrEmpty(dto.LanguageLevel) ? dto.LanguageLevel : "Normal";
                settings.Length = !string.IsNullOrEmpty(dto.Length) ? dto.Length : "Medium (5-6 sentences)";
                settings.Audience = !string.IsNullOrEmpty(dto.Audience) ? dto.Audience : "college students, computer science";
                settings.OverwriteExisting = dto.OverwriteExisting;

                if (!string.IsNullOrEmpty(dto.EncryptedApiKey))
                {
                    try
                    {
                        byte[] encryptedBytes = Convert.FromBase64String(dto.EncryptedApiKey);
                        byte[] decryptedBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
                        settings.ApiKey = Encoding.UTF8.GetString(decryptedBytes);
                    }
                    catch
                    {
                        settings.ApiKey = string.Empty;
                    }
                }
            }
            catch
            {
                // Fallback to defaults on read errors
            }
            return settings;
        }

        public static void Save(GenerationSettings settings)
        {
            try
            {
                if (!Directory.Exists(SettingsDir))
                {
                    Directory.CreateDirectory(SettingsDir);
                }

                string encryptedApiKey = string.Empty;
                if (!string.IsNullOrEmpty(settings.ApiKey))
                {
                    try
                    {
                        byte[] plainBytes = Encoding.UTF8.GetBytes(settings.ApiKey);
                        byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                        encryptedApiKey = Convert.ToBase64String(encryptedBytes);
                    }
                    catch
                    {
                        encryptedApiKey = string.Empty;
                    }
                }

                var dto = new SettingsDataTransfer
                {
                    Provider = (int)settings.Provider,
                    EncryptedApiKey = encryptedApiKey,
                    Model = settings.Model,
                    LanguageLevel = settings.LanguageLevel,
                    Length = settings.Length,
                    Audience = settings.Audience,
                    OverwriteExisting = settings.OverwriteExisting
                };

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(dto);
                File.WriteAllText(SettingsFilePath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error saving settings: " + ex.Message);
            }
        }
    }
}
