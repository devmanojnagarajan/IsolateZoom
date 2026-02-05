using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using wf = System.Windows.Forms;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;


namespace SectionPlaneZoom
{
    [PluginAttribute("IsolateZoom", "devmanojnagarajan", DisplayName = "SectionPlaneZoom",
        ToolTip = "Create Section Plane and Zoom", ExtendedToolTip = "Create a Section Plane at the clash point and zoom in to that plane")]
    public class SectionPlaneZoomView : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            if (doc == null)
            {
                return 0;
            }

            Autodesk.Navisworks.Api.Clash.DocumentClash clashDoc = doc.GetClash();

            if (clashDoc == null)
            {
                MessageBox.Show($"No Clash Test Detected");
                return 0;
            }

            var selectedTest = ClashHelper.ClashTestSelect(clashDoc.TestsData.Tests);
            if (selectedTest == null)
            {
                MessageBox.Show($"No Clash Test Detected");
            }
            if (selectedTest != null)
            {
                MessageBox.Show($"Selected Test: {selectedTest.DisplayName}");
            }

            return 0;
        }

        /*
        progress bar creation for status bar display using wpf
        */

        public class ProgressForm : wf.Form
        {

            private wf.ProgressBar progressBar;
            private wf.Label lblStatus;
            private wf.Label lblProgress;
            private wf.Button btnCancel;
            public bool CancelRequested { get; private set; } = false;

            public ProgressForm(int total)

            {
                this.Text = "Processing...";
                this.Width = 450; this.Height = 180;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false; this.MinimizeBox = false;
                this.ControlBox = false; this.TopMost = true;

                lblStatus = new Label { Text = "Initializing...", Left = 20, Top = 20, Width = 400, AutoEllipsis = true };
                lblProgress = new Label { Text = "0 / " + total, Left = 20, Top = 45, Width = 400 };
                progressBar = new ProgressBar { Left = 20, Top = 70, Width = 395, Height = 25, Minimum = 0, Maximum = total, Value = 0 };
                btnCancel = new Button { Text = "Cancel", Left = 170, Top = 105, Width = 100, Height = 30 };

                btnCancel.Click += (s, e) =>
                {
                    CancelRequested = true;
                    btnCancel.Enabled = false;
                    btnCancel.Text = "Cancelling...";
                };

                this.Controls.Add(lblStatus);
                this.Controls.Add(lblProgress);
                this.Controls.Add(progressBar);
                this.Controls.Add(btnCancel);



            }
            public void UpdateProgress(int current, int total, string message)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => UpdateProgress(current, total, message)));
                    return;
                }
                progressBar.Value = Math.Min(current, progressBar.Maximum);
                lblProgress.Text = $"{current} / {total}";
                lblStatus.Text = message;
                this.Refresh();
            }
        }


        /*
        ClashHelper for all clash related needs
        1 identifying all clash reports.
        2 getting the clash instance given clash test
        3 getting the elements of the clash instance provided clash instance
        4 getting the clash point of an clash instance
        5 creating section plane on a given clash point

        */
        public static class ClashHelper
        {
            public static Autodesk.Navisworks.Api.Clash.ClashTest ClashTestSelect(SavedItemCollection tests)
            {
                List<Autodesk.Navisworks.Api.Clash.ClashTest> clashTests = new List<Autodesk.Navisworks.Api.Clash.ClashTest>();

                foreach (SavedItem item in tests)
                {
                    if (item is Autodesk.Navisworks.Api.Clash.ClashTest ct)
                    {
                        clashTests.Add(ct);
                    }
                }

                if (clashTests.Count == 0)
                {
                    return null;
                }
                if (clashTests.Count == 1)
                {
                    return clashTests[0];
                }

                using (Form form = new Form())
                {
                    form.Text = "Select Clash Test";
                    form.Width = 350;
                    form.Height = 150;

                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    ComboBox combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Left = 20, Top = 20, Width = 290 };
                    foreach (var ct in clashTests) combo.Items.Add(ct.DisplayName);
                    combo.SelectedIndex = 0;

                    Button btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 150, Top = 60, Width = 75 };
                    Button btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 235, Top = 60, Width = 75 };

                    form.Controls.Add(combo);
                    form.Controls.Add(btnOk);
                    form.Controls.Add(btnCancel);
                    form.AcceptButton = btnOk;

                    if (form.ShowDialog() == DialogResult.OK)
                        return clashTests[combo.SelectedIndex];
                }

                return null;

            }
        }

        /*
            Provided a clash test, filter out only those that are in new , active or reviewed status and return the list of clash results        
        */

        public static List<ClashResult> GetFilteredClashResults(ClashTest test)
        {
            List<ClashResult> results = new List<ClashResult>();

            foreach (SavedItem item in test.Children)
            {
                if (item is ClashResult cr)
                {
                    // Filter: Only include New, Active, or Reviewed clashes
                    if (cr.Status == ClashResultStatus.New ||
                        cr.Status == ClashResultStatus.Active ||
                        cr.Status == ClashResultStatus.Reviewed)
                    {
                        results.Add(cr);
                    }
                }
            }

            return results;
        }

        /*
            given a clash instance of an clash test,
            i want to conduct a series of actions:
            there is a focus on clash button on the clash detective panel. i want to use that functionality
            then once focused, i want to select 1st item of the clash and, then section plane and turn on the section plane.
            this should provide a provide a proper zoomed in view of the clash point with both elements visible with a section plane.
        */

        public static void ViewPointCreation(Document doc, ClashResult clash, InwOpState10 comState, GroupItem folder)
        {
            if (clash == null || clash.Status != ClashResultStatus.New || clash.Status != ClashResultStatus.Active || clash.Status != ClashResultStatus.Reviewed)
            {
                return;
            }

            ModelItemCollection clashElements = new ModelItemCollection();
            
            if (clash.Item1 != null)
            {
                clashElements.Add(clash.Item1);
            }
            if (clash.Item2 != null)
            {
                clashElements.Add(clash.Item2);
            }

            Point3D clashPoint = clash.Center;

            // now zoom into the clashElements emulating focus on clash command
            doc.CurrentSelection.CopyFrom(clashElements);
            comState.ZoomInCurViewOnCurSel();


            // create a section plane on the clash point 
            dynamic view = comState.CurrentView;
            dynamic clipPlane = view.ClippingPlanes();

            while (clipPlane.Count > 0)
            { 
                clipPlane.RemovePlane(1);
            }

            dynamic plane = comState.ObjectFactory(nwEObjectType.eObjectType_nwOaClipPlane, null, null);

            plane.Plane.SetValue(0, 0, 1, -clashPoint.Z);

            // enable clipping plane 
            plane.Enabled = true;
            clipPlane.AddPlane(plane);
            clipPlane.Enabled = true;


            // save it to the specific folder with the clash name 
            Viewpoint currentViewPoint = doc.CurrentViewpoint.ToViewpoint();
            SavedViewpoint savedVP = new SavedViewpoint(currentViewPoint);
            savedVP.DisplayName = clash.DisplayName;

            doc.SavedViewpoints.InsertCopy(folder, folder.Children.Count, savedVP);        

            

        }
    }


}