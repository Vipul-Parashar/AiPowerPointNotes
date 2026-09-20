using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using AiSpeakerNotes.Models;
using AiSpeakerNotes.Services;

namespace AiSpeakerNotes
{
    [ComVisible(true)]
    [Guid("D8E5F2A1-9B3C-4E7F-8A12-34567890ABCD")]
    [ProgId("AiSpeakerNotes.TaskPaneControl")]
    public class TaskPaneControl : UserControl
    {
        private dynamic _pptApp;
        private GenerationSettings _settings;
        private readonly LlmService _llmService = new LlmService();
        private volatile bool _cancelRequested = false;
        private Thread _workerThread;

        // UI Controls
        private Label lblHeader;
        private Label lblSubHeader;

        private GroupBox grpStyle;
        private Label lblLevel;
        private ComboBox cboLevel;
        private Label lblLength;
        private ComboBox cboLength;
        private Label lblAudience;
        private TextBox txtAudience;
        private CheckBox chkOverwrite;

        private GroupBox grpProvider;
        private Label lblProvider;
        private ComboBox cboProvider;
        private Label lblApiKey;
        private TextBox txtApiKey;
        private Button btnToggleKey;
        private Label lblModel;
        private TextBox txtModel;
        private Button btnSaveSettings;

        private GroupBox grpActions;
        private Button btnGenerateAll;
        private Button btnGenerateCurrent;
        private Button btnCancel;
        private ProgressBar progressBar;
        private Label lblStatus;

        private GroupBox grpPreview;
        private TextBox txtLog;

        public TaskPaneControl()
        {
            InitializeComponent();
            _settings = SettingsManager.Load();
            LoadSettingsToUI();
        }

        public void SetPowerPointApp(object app)
        {
            _pptApp = app;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.AutoScroll = true;
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.Size = new Size(340, 750);
            this.Padding = new Padding(10);

            int y = 10;

            // Header
            lblHeader = new Label
            {
                Text = "AI Speaker Notes",
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = Color.FromArgb(196, 62, 28), // PowerPoint themed accent
                Location = new Point(10, y),
                AutoSize = true
            };
            y += 28;

            lblSubHeader = new Label
            {
                Text = "Natural spoken notes for teachers in Presenter View",
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(10, y),
                Size = new Size(320, 20)
            };
            y += 25;

            // Group 1: Style Settings
            grpStyle = new GroupBox
            {
                Text = " Teacher Voice & Note Style ",
                Location = new Point(10, y),
                Size = new Size(310, 195),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            lblLevel = new Label { Text = "Language Level:", Location = new Point(10, 24), Size = new Size(100, 20) };
            cboLevel = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(115, 20),
                Size = new Size(185, 23)
            };
            cboLevel.Items.AddRange(new object[] { "Normal", "Very Simple" });
            cboLevel.SelectedIndex = 0;

            lblLength = new Label { Text = "Length:", Location = new Point(10, 56), Size = new Size(100, 20) };
            cboLength = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(115, 52),
                Size = new Size(185, 23)
            };
            cboLength.Items.AddRange(new object[] {
                "Short (3-4 sentences)",
                "Medium (5-6 sentences)",
                "Detailed (7-8 sentences)"
            });
            cboLength.SelectedIndex = 1;

            lblAudience = new Label { Text = "Audience:", Location = new Point(10, 88), Size = new Size(100, 20) };
            txtAudience = new TextBox
            {
                Location = new Point(115, 84),
                Size = new Size(185, 23),
                Text = "college students, computer science"
            };

            chkOverwrite = new CheckBox
            {
                Text = "Overwrite existing notes (skip if unchecked)",
                Location = new Point(10, 120),
                Size = new Size(290, 24),
                Checked = false
            };

            Label lblTransitionInfo = new Label
            {
                Text = "✓ Smooth slide-to-slide transitions included\n✓ Real-life analogies for abstract ideas",
                Font = new Font("Segoe UI", 8F, FontStyle.Regular),
                ForeColor = Color.FromArgb(70, 130, 80),
                Location = new Point(10, 150),
                Size = new Size(290, 36)
            };

            grpStyle.Controls.AddRange(new Control[] {
                lblLevel, cboLevel, lblLength, cboLength,
                lblAudience, txtAudience, chkOverwrite, lblTransitionInfo
            });
            y += 205;

            // Group 2: Provider Settings
            grpProvider = new GroupBox
            {
                Text = " AI Model & API Key ",
                Location = new Point(10, y),
                Size = new Size(310, 160),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            lblProvider = new Label { Text = "Provider:", Location = new Point(10, 24), Size = new Size(70, 20) };
            cboProvider = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(85, 20),
                Size = new Size(215, 23)
            };
            cboProvider.Items.AddRange(new object[] {
                "Google Gemini",
                "OpenAI",
                "Anthropic Claude",
                "Offline Mock (Test Mode)"
            });
            cboProvider.SelectedIndex = 0;
            cboProvider.SelectedIndexChanged += CboProvider_SelectedIndexChanged;

            lblApiKey = new Label { Text = "API Key:", Location = new Point(10, 56), Size = new Size(70, 20) };
            txtApiKey = new TextBox
            {
                Location = new Point(85, 52),
                Size = new Size(165, 23),
                UseSystemPasswordChar = true
            };
            btnToggleKey = new Button
            {
                Text = "👁",
                Location = new Point(255, 51),
                Size = new Size(45, 25)
            };
            btnToggleKey.Click += (s, e) => {
                txtApiKey.UseSystemPasswordChar = !txtApiKey.UseSystemPasswordChar;
                btnToggleKey.Text = txtApiKey.UseSystemPasswordChar ? "👁" : "🔒";
            };

            lblModel = new Label { Text = "Model:", Location = new Point(10, 88), Size = new Size(70, 20) };
            txtModel = new TextBox
            {
                Location = new Point(85, 84),
                Size = new Size(215, 23),
                Text = ProviderDefaults.DefaultGeminiModel
            };

            btnSaveSettings = new Button
            {
                Text = "Save Key & Settings",
                Location = new Point(85, 118),
                Size = new Size(215, 28),
                BackColor = Color.FromArgb(235, 238, 242)
            };
            btnSaveSettings.Click += (s, e) => SaveSettingsFromUI(true);

            grpProvider.Controls.AddRange(new Control[] {
                lblProvider, cboProvider, lblApiKey, txtApiKey, btnToggleKey,
                lblModel, txtModel, btnSaveSettings
            });
            y += 170;

            // Group 3: Actions
            grpActions = new GroupBox
            {
                Text = " Actions ",
                Location = new Point(10, y),
                Size = new Size(310, 155),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            btnGenerateAll = new Button
            {
                Text = "▶  Generate Notes (All Slides)",
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(196, 62, 28),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(10, 22),
                Size = new Size(290, 34)
            };
            btnGenerateAll.FlatAppearance.BorderSize = 0;
            btnGenerateAll.Click += (s, e) => StartGeneration(allSlides: true);

            btnGenerateCurrent = new Button
            {
                Text = "🔄  Regenerate Current Slide",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Location = new Point(10, 62),
                Size = new Size(185, 28)
            };
            btnGenerateCurrent.Click += (s, e) => StartGeneration(allSlides: false);

            btnCancel = new Button
            {
                Text = "⏹ Cancel",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Location = new Point(200, 62),
                Size = new Size(100, 28),
                Enabled = false
            };
            btnCancel.Click += (s, e) => CancelGeneration();

            progressBar = new ProgressBar
            {
                Location = new Point(10, 98),
                Size = new Size(290, 16),
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            lblStatus = new Label
            {
                Text = "Ready. Click 'Generate Notes' to begin.",
                Location = new Point(10, 120),
                Size = new Size(290, 26),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            grpActions.Controls.AddRange(new Control[] {
                btnGenerateAll, btnGenerateCurrent, btnCancel, progressBar, lblStatus
            });
            y += 165;

            // Group 4: Log / Preview
            grpPreview = new GroupBox
            {
                Text = " Notes Log & Live Preview ",
                Location = new Point(10, y),
                Size = new Size(310, 170),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            txtLog = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 8.5F),
                Location = new Point(10, 20),
                Size = new Size(290, 140)
            };
            grpPreview.Controls.Add(txtLog);

            this.Controls.AddRange(new Control[] {
                lblHeader, lblSubHeader, grpStyle, grpProvider, grpActions, grpPreview
            });

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void CboProvider_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cboProvider.SelectedIndex)
            {
                case 0: // Gemini
                    txtModel.Text = ProviderDefaults.DefaultGeminiModel;
                    txtApiKey.Enabled = true;
                    break;
                case 1: // OpenAI
                    txtModel.Text = ProviderDefaults.DefaultOpenAiModel;
                    txtApiKey.Enabled = true;
                    break;
                case 2: // Claude
                    txtModel.Text = ProviderDefaults.DefaultClaudeModel;
                    txtApiKey.Enabled = true;
                    break;
                case 3: // Mock
                    txtModel.Text = ProviderDefaults.DefaultMockModel;
                    txtApiKey.Enabled = false;
                    break;
            }
        }

        private void LoadSettingsToUI()
        {
            if (_settings == null) return;

            cboProvider.SelectedIndex = (int)_settings.Provider;
            txtApiKey.Text = _settings.ApiKey;

            string m = _settings.Model;
            if (string.IsNullOrEmpty(m) || 
                (_settings.Provider == LlmProvider.Gemini && (m.IndexOf("1.5", StringComparison.OrdinalIgnoreCase) >= 0 || m.IndexOf("2.5-flash", StringComparison.OrdinalIgnoreCase) >= 0 || m.IndexOf("3.6-flash", StringComparison.OrdinalIgnoreCase) >= 0)))
            {
                m = ProviderDefaults.GetDefaultModel(_settings.Provider);
            }
            txtModel.Text = m;

            if (cboLevel.Items.Contains(_settings.LanguageLevel))
                cboLevel.SelectedItem = _settings.LanguageLevel;

            if (cboLength.Items.Contains(_settings.Length))
                cboLength.SelectedItem = _settings.Length;

            txtAudience.Text = _settings.Audience;
            chkOverwrite.Checked = _settings.OverwriteExisting;
        }

        private void SaveSettingsFromUI(bool showPrompt = false)
        {
            _settings.Provider = (LlmProvider)cboProvider.SelectedIndex;
            _settings.ApiKey = txtApiKey.Text.Trim();
            _settings.Model = txtModel.Text.Trim();
            _settings.LanguageLevel = cboLevel.SelectedItem != null ? cboLevel.SelectedItem.ToString() : "Normal";
            _settings.Length = cboLength.SelectedItem != null ? cboLength.SelectedItem.ToString() : "Medium (5-6 sentences)";
            _settings.Audience = txtAudience.Text.Trim();
            _settings.OverwriteExisting = chkOverwrite.Checked;

            SettingsManager.Save(_settings);

            if (showPrompt)
            {
                MessageBox.Show("Settings and API key saved securely!", "AI Speaker Notes", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void StartGeneration(bool allSlides)
        {
            SaveSettingsFromUI(false);

            if (_pptApp == null)
            {
                try
                {
                    _pptApp = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
                }
                catch
                {
                    MessageBox.Show("Could not attach to PowerPoint. Please ensure PowerPoint is running.", "AI Speaker Notes", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            if (_settings.Provider != LlmProvider.Mock && string.IsNullOrWhiteSpace(_settings.ApiKey))
            {
                MessageBox.Show("Please enter an API Key for " + _settings.Provider + " or switch to 'Offline Mock' mode to test.", "API Key Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtApiKey.Focus();
                return;
            }

            btnGenerateAll.Enabled = false;
            btnGenerateCurrent.Enabled = false;
            btnCancel.Enabled = true;
            _cancelRequested = false;

            progressBar.Value = 0;
            txtLog.Clear();
            AppendLog("Starting notes generation...");

            _workerThread = new Thread(() => ProcessSlidesWorker(allSlides))
            {
                IsBackground = true
            };
            _workerThread.Start();
        }

        private void CancelGeneration()
        {
            _cancelRequested = true;
            AppendLog("Cancellation requested. Stopping after current slide...");
            btnCancel.Enabled = false;
        }

        private void ProcessSlidesWorker(bool allSlides)
        {
            try
            {
                dynamic pres = null;
                try
                {
                    pres = _pptApp.ActivePresentation;
                }
                catch
                {
                    pres = null;
                }

                if (pres == null || (int)pres.Slides.Count == 0)
                {
                    UpdateUI(() =>
                    {
                        AppendLog("No active presentation or slides found.");
                        lblStatus.Text = "No active presentation found.";
                        ResetButtons();
                    });
                    return;
                }

                int totalSlides = (int)pres.Slides.Count;
                var slideIndices = new List<int>();

                if (allSlides)
                {
                    for (int i = 1; i <= totalSlides; i++) slideIndices.Add(i);
                }
                else
                {
                    int currentSlide = 1;
                    try
                    {
                        dynamic curSlide = _pptApp.ActiveWindow.View.Slide;
                        if (curSlide != null)
                        {
                            currentSlide = (int)curSlide.SlideIndex;
                        }
                    }
                    catch
                    {
                        currentSlide = 1;
                    }
                    slideIndices.Add(currentSlide);
                }

                // Pre-extract slide summaries for neighbor context
                var slideDataMap = new Dictionary<int, SlideData>();
                foreach (int idx in slideIndices)
                {
                    var s = pres.Slides[idx];
                    slideDataMap[idx] = SlideContentExtractor.ExtractSlide(s);
                }

                int processedCount = 0;
                int skippedCount = 0;

                for (int i = 0; i < slideIndices.Count; i++)
                {
                    if (_cancelRequested)
                    {
                        UpdateUI(() => {
                            AppendLog("Generation stopped by user.");
                            lblStatus.Text = "Generation stopped by user.";
                        });
                        break;
                    }

                    int slideIdx = slideIndices[i];
                    var slideObj = pres.Slides[slideIdx];
                    var slideData = slideDataMap[slideIdx];

                    // Check overwrite setting
                    if (slideData.HasNotes && !_settings.OverwriteExisting)
                    {
                        UpdateUI(() => {
                            AppendLog(string.Format("Slide {0} ('{1}'): Skipped (notes exist, overwrite disabled).", slideIdx, slideData.Title));
                        });
                        skippedCount++;
                        int pct = (int)(((i + 1) / (float)slideIndices.Count) * 100);
                        UpdateUI(() => {
                            progressBar.Value = Math.Min(100, pct);
                            lblStatus.Text = string.Format("Slide {0} of {1} (Skipped)", slideIdx, totalSlides);
                        });
                        continue;
                    }

                    if (slideData.IsEmpty)
                    {
                        UpdateUI(() => {
                            AppendLog(string.Format("Slide {0}: Skipped (slide has no content).", slideIdx));
                        });
                        skippedCount++;
                        continue;
                    }

                    // Populate neighbor summaries
                    if (slideIdx > 1)
                    {
                        try
                        {
                            var prevSlide = pres.Slides[slideIdx - 1];
                            var prevData = SlideContentExtractor.ExtractSlide(prevSlide);
                            slideData.PrevSlideSummary = SlideContentExtractor.SummarizeSlide(prevData);
                        }
                        catch { }
                    }
                    if (slideIdx < totalSlides)
                    {
                        try
                        {
                            var nextSlide = pres.Slides[slideIdx + 1];
                            var nextData = SlideContentExtractor.ExtractSlide(nextSlide);
                            slideData.NextSlideSummary = SlideContentExtractor.SummarizeSlide(nextData);
                        }
                        catch { }
                    }

                    UpdateUI(() => {
                        lblStatus.Text = string.Format("Processing Slide {0} of {1}: \"{2}\"...", slideIdx, totalSlides, slideData.Title);
                        AppendLog(string.Format("\n[Slide {0}] Requesting notes from {1} ({2})...", slideIdx, _settings.Provider, _settings.Model));
                    });

                    try
                    {
                        string notes = _llmService.GenerateNotesWithRetry(slideData, _settings, maxRetries: 3);

                        // Write notes into PowerPoint slide NotesPage
                        NotesWriter.SetNotes(slideObj, notes);

                        processedCount++;
                        UpdateUI(() => {
                            AppendLog(string.Format("✓ Slide {0} Notes saved:\n\"{1}\"", slideIdx, notes));
                        });
                    }
                    catch (Exception ex)
                    {
                        UpdateUI(() => {
                            AppendLog(string.Format("✗ Slide {0} Error: {1}", slideIdx, ex.Message));
                        });
                    }

                    int progressPct = (int)(((i + 1) / (float)slideIndices.Count) * 100);
                    UpdateUI(() => {
                        progressBar.Value = Math.Min(100, progressPct);
                    });

                    // Modest polite delay between slides to respect API rate limits
                    Thread.Sleep(300);
                }

                UpdateUI(() => {
                    lblStatus.Text = string.Format("Finished: {0} generated, {1} skipped.", processedCount, skippedCount);
                    AppendLog(string.Format("\n=== Done! Generated: {0}, Skipped: {1} ===", processedCount, skippedCount));
                    AppendLog("Tip: Go to 'Slide Show > Presenter View' (or Alt+F5) to see your notes live!");
                    ResetButtons();
                });
            }
            catch (Exception ex)
            {
                UpdateUI(() => {
                    AppendLog("Critical error during processing: " + ex.Message);
                    lblStatus.Text = "Error during generation.";
                    ResetButtons();
                });
            }
        }

        private void ResetButtons()
        {
            btnGenerateAll.Enabled = true;
            btnGenerateCurrent.Enabled = true;
            btnCancel.Enabled = false;
        }

        private void AppendLog(string message)
        {
            txtLog.AppendText(message + Environment.NewLine);
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private void UpdateUI(Action action)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }
    }
}
