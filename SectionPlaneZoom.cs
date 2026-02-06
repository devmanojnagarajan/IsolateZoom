using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using wf = System.Windows.Forms;
using NavisColor = Autodesk.Navisworks.Api.Color;


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
            if (clashDoc.TestsData?.Tests == null || clashDoc.TestsData.Tests.Count == 0)
            {
                MessageBox.Show($"No Clash Test Detected");
                return 0;
            }


            var selectedTest = ClashHelper.ClashTestSelect(clashDoc.TestsData.Tests);
            if (selectedTest == null)
            {
                return 0;
            }

            List<ClashResult> results = ClashHelper.GetFilteredClashResults(selectedTest);
            if (results.Count == 0) { MessageBox.Show("No valid clashes to process."); return 0; }

            // COM and Folder
            InwOpState10 comState = ComBridge.State;
            GroupItem folder = ClashHelper.CreateViewPointFolder(doc, "Clash Section Views");

            using (ProgressForm progress = new ProgressForm(results.Count))
            {
                progress.Show();

                List<string> failedClashes = new List<string>();

                for (int i = 0; i < results.Count; i++)
                {
                    if (progress.CancelRequested) break;

                    ClashResult clash = results[i];
                    progress.UpdateProgress(i + 1, results.Count, $"Processing: {clash.DisplayName}");

                    try
                    {
                        ClashHelper.CreateViewPointForClash(doc, clash, comState, folder);
                    }
                    catch (Exception ex)
                    {
                        failedClashes.Add($"{clash.DisplayName}: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"Error processing {clash.DisplayName}: {ex}");
                    }
                }

                // After loop completes
                if (failedClashes.Count > 0)
                {
                    string summary = failedClashes.Count <= 5
                        ? string.Join("\n", failedClashes)
                        : $"{failedClashes.Count} clashes failed. Check debug output for details.";

                    MessageBox.Show($"Completed with errors:\n{summary}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Viewpoint creation complete.");
                }
            }
            doc.CurrentSelection.Clear();
            doc.Models.ResetAllPermanentMaterials();
            ClashHelper.DisableSectioning(comState);
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
            public static bool ViewpointExists(GroupItem folder, string name)
            {
                if (folder?.Children == null) return false;

                foreach (SavedItem item in folder.Children)
                {
                    if (item.DisplayName == name)
                    {
                        return true;
                    }
                }
                return false;
            }
            public static GroupItem CreateViewPointFolder(Document doc, string folderName)
            {

                foreach (SavedItem item in doc.SavedViewpoints.Value)
                {
                    if (item is GroupItem existing && existing.DisplayName == folderName)
                    {
                        return existing;
                    }
                }

                FolderItem folder = new FolderItem { DisplayName = folderName };
                doc.SavedViewpoints.AddCopy(folder);

                SavedItem added = doc.SavedViewpoints.Value[doc.SavedViewpoints.Value.Count - 1];

                if (added is GroupItem groupItem)
                {
                    return groupItem;
                }

                throw new InvalidOperationException("Failed to create viewpoint folder.");
            }
            public static void DisableSectioning(InwOpState10 comState)
            {
                if (comState == null) return;
                try
                {
                    dynamic view = comState.CurrentView;
                    dynamic clipPlanes = view.ClippingPlanes();
                    clipPlanes.Enabled = false;
                    while (clipPlanes.Count > 0) clipPlanes.RemovePlane(1);
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine($"DisableSectioning failed: {e.Message}");
                }
            }

            /*
            given a clash instance of an clash test

            
            there is a focus on clash button on the clash detective panel. emulate that functionality
            
            zoomed in view of the clash point with both elements visible with a section plane.
            */

            public static void CreateViewPointForClash(Document doc, ClashResult clash, InwOpState10 comState, GroupItem folder)
            {
                if (clash == null) return;

                if (clash.Center == null ||
                    double.IsNaN(clash.Center.X) ||
                    double.IsNaN(clash.Center.Y) ||
                    double.IsNaN(clash.Center.Z))
                {
                    System.Diagnostics.Debug.WriteLine($"Skipping clash '{clash.DisplayName}': Invalid center point.");
                    return;
                }

                dynamic view = null;
                dynamic clipPlanes = null;
                dynamic plane = null;

                try
                {
                    ModelItemCollection clashElements = new ModelItemCollection();
                    if (clash.Item1 != null) clashElements.Add(clash.Item1);
                    if (clash.Item2 != null) clashElements.Add(clash.Item2);

                    if (clashElements.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping clash '{clash.DisplayName}': No elements found.");
                        return;

                    }
                    Point3D clashPoint = clash.Center;

                    doc.Models.OverridePermanentColor(clashElements, new NavisColor(1.0, 0.0, 0.0));


                    doc.CurrentSelection.CopyFrom(clashElements);
                    comState.ZoomInCurViewOnCurSel();

                    if (doc.CurrentViewpoint == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping clash '{clash.DisplayName}': Failed to set viewpoint.");
                        return;
                    }

                    // COM operations
                    view = comState.CurrentView;
                    clipPlanes = view.ClippingPlanes();

                    // Clear existing planes
                    while (clipPlanes.Count > 0)
                    {
                        clipPlanes.RemovePlane(1);
                    }

                    // Create new plane
                    plane = comState.ObjectFactory(nwEObjectType.eObjectType_nwOaClipPlane, null, null);
                    plane.Plane.SetValue(0, 0, 1, -clashPoint.Z);
                    plane.Enabled = true;
                    clipPlanes.AddPlane(plane);
                    clipPlanes.Enabled = true;

                    // Save viewpoint
                    Viewpoint currentViewpoint = doc.CurrentViewpoint.ToViewpoint();
                    SavedViewpoint savedVP = new SavedViewpoint(currentViewpoint);
                    savedVP.DisplayName = clash.DisplayName;
                    if (folder == null || folder.Children == null)
                    {
                        throw new InvalidOperationException("Viewpoint folder is invalid.");
                    }

                    int insertIndex = folder.Children.Count;
                    doc.SavedViewpoints.InsertCopy(folder, insertIndex, savedVP);
                    doc.Models.ResetPermanentMaterials(clashElements);
                    doc.CurrentSelection.Clear();
                }
                finally
                {
                    // Release COM objects in reverse order of creation
                    if (plane != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(plane);
                    if (clipPlanes != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(clipPlanes);
                    if (view != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(view);
                }
            }
        }



    }


}