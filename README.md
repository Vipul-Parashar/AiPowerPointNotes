# AI Speaker Notes for PowerPoint

An intelligent PowerPoint Add-in designed specifically for teachers and educators. With one click, it analyzes every slide in your presentation—titles, text boxes, bullet points, data tables, and image/diagram alternative text—and generates clear, conversational speaker notes saved directly into PowerPoint's slide **Notes** field for **Presenter View**.

---

## Technical Approach & Architecture

### Why a C# COM Add-in instead of an Office.js Web Add-in?
- **Office.js Limitation**: As of 2026, Microsoft's JavaScript API for PowerPoint (`Office.js`) **does not provide any API to read or write slide speaker notes** (presenter notes). The `PowerPoint.Slide` API in Office.js only exposes slide shapes, layouts, and tags; notes are completely inaccessible programmatically.
- **The Solution**: We built a native **C# COM Add-in** (.NET Framework). Through the PowerPoint Object Model (`Slide.NotesPage`), it provides 100% reliable read and write access to slide notes. Notes appear immediately in the PowerPoint editing pane and in **Presenter View**.
- **No Admin Rights Required**: The add-in is registered per-user under `HKEY_CURRENT_USER\Software\Classes` and `HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins`, allowing instant installation without Administrator permissions.

---

## Project Structure

```
AiSpeakerNotes/
├── bin/
│   └── AiSpeakerNotes.dll            # Compiled COM Add-in assembly
├── src/
│   ├── AssemblyInfo.cs               # COM attributes and GUID declarations
│   ├── Connect.cs                    # IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
│   ├── TaskPaneControl.cs            # Windows Forms Task Pane UI (controls, progress, live log)
│   ├── Models/
│   │   ├── GenerationSettings.cs     # Settings model (provider, model, tone, length, audience)
│   │   ├── LlmProvider.cs            # Provider enum (Gemini, OpenAI, Claude, Mock)
│   │   └── SlideData.cs              # Extracted slide data and neighbor context
│   └── Services/
│       ├── LlmService.cs             # Multi-LLM client (Gemini, OpenAI, Claude, Mock) with retry logic
│       ├── NotesWriter.cs            # Dynamic reader & writer for Slide.NotesPage
│       ├── PromptBuilder.cs          # Pedagogical teacher prompt generator
│       ├── SettingsManager.cs        # DPAPI-encrypted credential storage
│       └── SlideContentExtractor.cs  # Reads titles, shapes, bullets, tables, and alt text
├── scripts/
│   ├── build.ps1                     # Compiles the assembly using 64-bit csc.exe
│   ├── install.ps1                   # Registers the COM Add-in in HKCU
│   ├── uninstall.ps1                 # Cleanly removes registry keys
│   ├── create_sample_presentation.ps1# Creates a realistic 5-slide educational presentation
│   ├── test_generator.ps1            # End-to-end automated test runner
│   └── open_presentation.ps1         # Launches PowerPoint to inspect notes
├── build.bat                         # 1-click build launcher
├── install.bat                       # 1-click installer / PowerPoint upload
├── uninstall.bat                     # 1-click uninstaller
├── test_demo.bat                     # 1-click demo test and inspection
└── README.md                         # Documentation
```

---

## How to Upload / Install the Add-in to PowerPoint

You can load the add-in into PowerPoint in less than 10 seconds:

### Step 1: Run the 1-Click Installer
Double-click **`install.bat`** (or open PowerShell and run):
```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\install.ps1"
```
This registers the add-in under `HKCU\Software\Microsoft\Office\PowerPoint\Addins\AiSpeakerNotes.Connect` and sets `LoadBehavior = 3` (load automatically on startup).

### Step 2: Launch PowerPoint
Open **PowerPoint**. You will see:
1. A new tab on the top ribbon: **"AI Speaker Notes"**.
2. A large button: **"Generate Notes"**.
3. A button: **"Notes Task Pane"**.

### (Optional) Verifying Add-in Status in PowerPoint Options
If you ever want to check or re-enable the add-in inside PowerPoint:
1. In PowerPoint, click **File** > **Options**.
2. Select **Add-ins** on the left menu.
3. At the bottom, in the **Manage** drop-down, select **COM Add-ins** and click **Go...**.
4. You will see **"AI Speaker Notes Generator"** checked in the list.

---

## How to Use the Add-in

### 1. Open Your Presentation
Open any presentation in PowerPoint.

### 2. Open the Task Pane
Click the **"Notes Task Pane"** button on the **AI Speaker Notes** ribbon tab. The task pane will open on the right side of the window.

### 3. Customize Note Style & Settings
- **Teacher Voice & Note Style**:
  - **Language Level**: Choose *Normal* or *Very Simple* (for middle school or complete beginners).
  - **Length**: *Short (3-4 sentences)*, *Medium (5-6 sentences)*, or *Detailed (7-8 sentences)*.
  - **Audience / Subject**: Enter your target audience (e.g., `"college students, computer science"`, `"high school chemistry"`).
  - **Overwrite existing notes**: When unchecked, slides that already have speaker notes will be skipped. Check this box if you want to rewrite all notes.
- **AI Model & Key Configuration**:
  - **Provider**: Select your preferred provider:
    - **Google Gemini**: Uses Google's `generativelanguage.googleapis.com` (default: `gemini-1.5-flash`).
    - **OpenAI**: Uses OpenAI's `api.openai.com` (default: `gpt-4o-mini`).
    - **Anthropic Claude**: Uses Anthropic's `api.anthropic.com` (default: `claude-3-5-sonnet-20241022`).
    - **Offline Mock (Test Mode)**: Generates realistic teacher notes offline without requiring an API key.
  - **API Key**: Enter your API key. Click the eye icon to toggle visibility.
  - **Save Key & Settings**: Click to store your settings. Keys are encrypted with Windows DPAPI (`ProtectedData`) and stored in `%LOCALAPPDATA%\AiSpeakerNotes\settings.json`.

### 4. Generate Notes
- Click **"▶ Generate Notes (All Slides)"** to process the entire deck.
- Or click **"🔄 Regenerate Current Slide"** to generate notes for only the slide currently visible on screen.
- You can track real-time progress via the progress bar, status message, and the live log preview box.

---

## Verifying Notes in Presenter View

Once notes are generated, you can present with confidence:

1. In PowerPoint, look at the **Notes** pane below any slide. The notes are already written there!
2. Press **Alt + F5** (or go to **Slide Show** > check **"Use Presenter View"** > click **"From Beginning"**).
3. In **Presenter View**:
   - The current slide is on the left.
   - The next slide is on the top right.
   - **Your AI-generated speaker notes appear prominently in the large Notes pane on the bottom right!**
   - You can click the **A+ / A-** font size buttons in Presenter View to make the text larger and easier to read while teaching.

---

## Testing Against the Sample 5-Slide Presentation

The repository includes an end-to-end automated testing suite and a 5-slide sample presentation covering:
- Slide 1: Title & Course Introduction
- Slide 2: Bullet points on "What is an Algorithm?"
- Slide 3: Comparison table ("Linear Search vs. Binary Search")
- Slide 4: Diagram shapes with Alt Text ("Divide and Conquer Strategy")
- Slide 5: Abstract concept ("Why Time Complexity Matters in the Real World")

To run the test demo:
Double-click **`test_demo.bat`** (or in PowerShell run):
```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\test_generator.ps1"
powershell -ExecutionPolicy Bypass -File ".\scripts\open_presentation.ps1"
```
This processes all 5 slides, writes notes to each slide's `NotesPage`, verifies read-back integrity, and launches PowerPoint so you can test Presenter View immediately!

---

## How to Customize the Note-Writing Prompt

All prompt instructions and teacher tone rules are centralized in [`src/Services/PromptBuilder.cs`](./src/Services/PromptBuilder.cs):

### System Prompt Rules
Located in `PromptBuilder.BuildSystemPrompt()`:
- **Plain, Simple English**: Instructs the model to speak like a warm, engaging educator.
- **Short Sentences**: Enforces punchy, spoken sentence structure.
- **Real-Life Analogies**: Enforces including an everyday analogy (like books on a shelf) for abstract ideas.
- **Transitions**: Instructs the LLM to begin with a natural bridge connecting from the previous slide.
- **Explain, Don't Recite**: Instructs the LLM never to simply read off bullets or table rows.

To modify the prompt:
1. Open [`src/Services/PromptBuilder.cs`](./src/Services/PromptBuilder.cs) in any text editor.
2. Adjust the rules or formatting in `BuildSystemPrompt` or `BuildUserPrompt`.
3. Double-click `build.bat` to recompile `bin\AiSpeakerNotes.dll`.
4. Restart PowerPoint to see your changes in action!

---

## Uninstallation

To cleanly unregister the add-in from PowerPoint:
Double-click **`uninstall.bat`** (or in PowerShell run):
```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\uninstall.ps1"
```
All registry keys in `HKCU` will be deleted.
