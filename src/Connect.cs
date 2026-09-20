using System;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AiSpeakerNotes
{
    [ComVisible(true)]
    [Guid("E4B5C6D7-1A2B-3C4D-5E6F-7A8B9C0D1E2F")]
    [ProgId("AiSpeakerNotes.Connect")]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        private dynamic _pptApp;
        private Office.CustomTaskPane _customTaskPane;
        private TaskPaneControl _taskPaneControl;

        #region IDTExtensibility2 Members

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _pptApp = application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (_customTaskPane != null)
            {
                try
                {
                    _customTaskPane.Visible = false;
                    _customTaskPane.Delete();
                }
                catch { }
                _customTaskPane = null;
            }
            _pptApp = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            EnsureTaskPaneCreated();
        }

        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region ICustomTaskPaneConsumer Members

        private Office.ICTPFactory _ctpFactory;

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _ctpFactory = CTPFactoryInst;
            EnsureTaskPaneCreated();
        }

        private void EnsureTaskPaneCreated()
        {
            if (_customTaskPane == null && _ctpFactory != null)
            {
                try
                {
                    _customTaskPane = _ctpFactory.CreateCTP("AiSpeakerNotes.TaskPaneControl", "AI Speaker Notes", Type.Missing);
                    _customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    _customTaskPane.Width = 340;

                    _taskPaneControl = _customTaskPane.ContentControl as TaskPaneControl;
                    if (_taskPaneControl != null && _pptApp != null)
                    {
                        _taskPaneControl.SetPowerPointApp(_pptApp);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error initializing Task Pane: " + ex.Message);
                }
            }
        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabAiSpeakerNotes"" label=""AI Speaker Notes"">
        <group id=""grpNotes"" label=""Speaker Notes Generator"">
          <button id=""btnGenerateAll"" 
                  label=""Generate Notes"" 
                  size=""large"" 
                  onAction=""OnGenerateAllClicked"" 
                  imageMso=""SlideNotesPage"" 
                  screentip=""Generate Notes (All Slides)"" 
                  supertip=""Reads each slide and writes natural, spoken teacher notes into the Notes field for Presenter View."" />
          <button id=""btnGenerateCurrent"" 
                  label=""Current Slide"" 
                  size=""large"" 
                  onAction=""OnGenerateCurrentClicked"" 
                  imageMso=""DirectRepliesTo"" 
                  screentip=""Regenerate Current Slide"" 
                  supertip=""Regenerates notes for the active slide only."" />
          <separator id=""sep1"" />
          <button id=""btnTogglePane"" 
                  label=""Notes Task Pane"" 
                  size=""large"" 
                  onAction=""OnTogglePaneClicked"" 
                  imageMso=""TaskPaneToggle"" 
                  screentip=""Show/Hide AI Notes Task Pane"" 
                  supertip=""Opens the settings, style preferences, and live generation log pane."" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        #endregion

        #region Ribbon Callbacks

        public void OnGenerateAllClicked(Office.IRibbonControl control)
        {
            ShowTaskPane();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.SetPowerPointApp(_pptApp);
                _taskPaneControl.StartGeneration(allSlides: true);
            }
        }

        public void OnGenerateCurrentClicked(Office.IRibbonControl control)
        {
            ShowTaskPane();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.SetPowerPointApp(_pptApp);
                _taskPaneControl.StartGeneration(allSlides: false);
            }
        }

        public void OnTogglePaneClicked(Office.IRibbonControl control)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = !_customTaskPane.Visible;
                if (_customTaskPane.Visible && _taskPaneControl != null && _pptApp != null)
                {
                    _taskPaneControl.SetPowerPointApp(_pptApp);
                }
            }
        }

        private void ShowTaskPane()
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = true;
                if (_taskPaneControl != null && _pptApp != null)
                {
                    _taskPaneControl.SetPowerPointApp(_pptApp);
                }
            }
        }

        #endregion
    }
}
