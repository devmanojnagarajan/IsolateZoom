using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Interop;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BF = System.Reflection.BindingFlags;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using NavisColor = Autodesk.Navisworks.Api.Color;
using wf = System.Windows.Forms;



namespace SectionPlaneZoom
{
    [PluginAttribute("IsolateZoom", "devmanojnagarajan", DisplayName = "SectionPlaneZoom",
        ToolTip = "Create Section Plane and Zoom",
        ExtendedToolTip = "Create a Section Plane at the clash point and zoom in to that plane")]
    public class SectionPlaneZoomView : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            if (doc == null) return 0;

            //DebugLog.Reset();
            //DebugLog.Log("=== Clash VP - Section Started ===");

            DocumentClash clashDoc = doc.GetClash();
            if (clashDoc == null || clashDoc.TestsData?.Tests == null || clashDoc.TestsData.Tests.Count == 0)
            {
                MessageBox.Show("No Clash Test Detected");
                return 0;
            }

            var selectedTest = ClashHelper.ClashTestSelect(clashDoc.TestsData.Tests);
            if (selectedTest == null) return 0;

            //DebugLog.Log($"Selected test: {selectedTest.DisplayName}");

            List<ClashResult> results = ClashHelper.GetFilteredClashResults(selectedTest);
            //DebugLog.Log($"Found {results.Count} clashes with status New/Active/Reviewed");

            if (results.Count == 0)
            {
                MessageBox.Show("No valid clashes to process.");
                return 0;
            }

            InwOpState10 comState = ComBridge.State;
            // checking view state name
            //ComHelper.DiscoverViewMethods(comState);
            //ComHelper.DiscoverViewpointClipping(doc);
            //ComHelper.DiscoverClippingPlanes(comState);
            //ComHelper.DiscoverClipPlaneSetMethods(doc);


            /*
            // === COM API Discovery ===
            //DebugLog.Log("--- COM API Discovery ---");
            try
            {
                object view = comState.CurrentView;
                //DebugLog.Log($"CurrentView type: {view.GetType().FullName}");

                foreach (Type iface in view.GetType().GetInterfaces())
                    DebugLog.Log($"  Interface: {iface.FullName}");

                object cp = view.GetType().InvokeMember("ClippingPlanes",
                    BF.InvokeMethod, null, view, null);
                //DebugLog.Log($"ClippingPlanes type: {cp.GetType().FullName}");

                foreach (Type iface in cp.GetType().GetInterfaces())
                    DebugLog.Log($"  CP Interface: {iface.FullName}");
            }
            catch (Exception dEx)
            {
                DebugLog.Log($"Discovery error: {dEx.Message}");
            }
            DebugLog.Log("--- End Discovery ---");
            DebugLog.Save();
            */

            // Create/find the viewpoint folder
            GroupItem folder = ClashHelper.CreateViewPointFolder(doc, "Clash Section Views");
            //DebugLog.Log($"Folder ready: '{folder.DisplayName}', existing children: {folder.Children.Count}");
            //DebugLog.Save();


            // -- Check redlines -- //
            string redlineJson = Autodesk.Navisworks.Api.Application.ActiveDocument.ActiveView.GetRedlines();
            //DebugLog.Log($"MANUAL REDLINE JSON: {redlineJson}");

            using (ProgressForm progress = new ProgressForm(results.Count))
            {
                progress.Show();
                List<string> failedClashes = new List<string>();
                int successCount = 0;

                for (int i = 0; i < results.Count; i++)
                {
                    if (progress.CancelRequested) break;

                    ClashResult clash = results[i];
                    progress.UpdateProgress(i + 1, results.Count, $"Processing: {clash.DisplayName}");

                    try
                    {
                        ClashHelper.CreateViewPointForClash(doc, clash, comState);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        string detail = $"{clash.DisplayName} | {ex.GetType().Name}: {ex.Message}";
                        failedClashes.Add(detail);
                        //DebugLog.Log($"EXCEPTION: {detail}");
                        //DebugLog.Save();
                    }
                }


                //DebugLog.Log($"Loop complete: {successCount} succeeded, {failedClashes.Count} failed");

                if (failedClashes.Count > 0)
                {
                    string summary = failedClashes.Count <= 5
                        ? string.Join("\n", failedClashes)
                        : $"{failedClashes.Count} clashes failed. Check ClashVP_Debug.log on Desktop.";
                    MessageBox.Show($"Completed with errors:\n{summary}", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show($"Viewpoint creation complete. {successCount} viewpoints created.");
                }
            }

            doc.CurrentSelection.Clear();
            doc.Models.ResetAllPermanentMaterials();
            ComHelper.DisableSectioningNet(doc);

            //not working enable sectioning
            //ComHelper.EnableSectioningForAllViewpoints(doc, "Clash Section Views");

            //DebugLog.Log("=== Finished ===");
            //DebugLog.Save();
            return 0;
        }


        // ──────────────────────────────────────────────
        // COM Helper — all COM access via reflection
        // ──────────────────────────────────────────────
        public static class ComHelper
        {
            public static void DisableSectioningNet(Document doc)
            {
                try
                {
                    Viewpoint vp = doc.CurrentViewpoint.ToViewpoint().CreateCopy();
                    Type vpType = vp.GetType();
                    PropertyInfo internalProp = vpType.GetProperty("InternalClipPlanes",
                        BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

                    if (internalProp != null)
                    {
                        dynamic clipPlaneSet = internalProp.GetValue(vp);
                        clipPlaneSet.SetEnabled(false);
                        doc.CurrentViewpoint.CopyFrom(vp);
                    }
                }
                catch (Exception ex)
                {
                    //DebugLog.Log($"  DisableSectioningNet failed: {ex.Message}");
                }
            }           

                        

        }

        // ──────────────────────────────────────────────
        // Debug Logger
        // ──────────────────────────────────────────────
        public static class DebugLog
        {
            private static List<string> _lines = new List<string>();
            private static DateTime _startTime = DateTime.Now;

            public static void Reset()
            {
                _lines = new List<string>();
                _startTime = DateTime.Now;
            }

            public static void Log(string message)
            {
                string elapsed = $"{(DateTime.Now - _startTime).TotalMilliseconds:F0}ms";
                string line = $"{DateTime.Now:HH:mm:ss.fff} [{elapsed}] {message}";
                _lines.Add(line);
                System.Diagnostics.Debug.WriteLine(line);
            }

            public static void Save()
            {
                try
                {
                    string logPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "ClashVP_Debug.log");
                    File.WriteAllLines(logPath, _lines);
                }
                catch { }
            }
        }


        // ──────────────────────────────────────────────
        // Progress Form
        // ──────────────────────────────────────────────
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
                this.Width = 450;
                this.Height = 180;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ControlBox = false;
                this.TopMost = true;

                lblStatus = new Label
                {
                    Text = "Initializing...",
                    Left = 20,
                    Top = 20,
                    Width = 400,
                    AutoEllipsis = true
                };
                lblProgress = new Label
                {
                    Text = "0 / " + total,
                    Left = 20,
                    Top = 45,
                    Width = 400
                };
                progressBar = new ProgressBar
                {
                    Left = 20,
                    Top = 70,
                    Width = 395,
                    Height = 25,
                    Minimum = 0,
                    Maximum = total,
                    Value = 0
                };
                btnCancel = new Button
                {
                    Text = "Cancel",
                    Left = 170,
                    Top = 105,
                    Width = 100,
                    Height = 30
                };

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


        // ──────────────────────────────────────────────
        // ClashHelper
        // ──────────────────────────────────────────────

        public static class ClashHelper
        {
            public static ClashTest ClashTestSelect(SavedItemCollection tests)
            {
                List<ClashTest> clashTests = new List<ClashTest>();
                foreach (SavedItem item in tests)
                {
                    if (item is ClashTest ct) clashTests.Add(ct);
                }

                if (clashTests.Count == 0) return null;
                if (clashTests.Count == 1) return clashTests[0];

                using (Form form = new Form())
                {
                    form.Text = "Select Clash Test";
                    form.Width = 350;
                    form.Height = 150;
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    ComboBox combo = new ComboBox
                    {
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Left = 20,
                        Top = 20,
                        Width = 290
                    };
                    foreach (var ct in clashTests) combo.Items.Add(ct.DisplayName);
                    combo.SelectedIndex = 0;

                    Button btnOk = new Button
                    {
                        Text = "OK",
                        DialogResult = DialogResult.OK,
                        Left = 150,
                        Top = 60,
                        Width = 75
                    };
                    Button btnCancel = new Button
                    {
                        Text = "Cancel",
                        DialogResult = DialogResult.Cancel,
                        Left = 235,
                        Top = 60,
                        Width = 75
                    };

                    form.Controls.Add(combo);
                    form.Controls.Add(btnOk);
                    form.Controls.Add(btnCancel);
                    form.AcceptButton = btnOk;

                    if (form.ShowDialog() == DialogResult.OK)
                        return clashTests[combo.SelectedIndex];
                }
                return null;
            }

            public static void CreateSelectionSetForClash(Document doc, ClashResult clash, ModelItemCollection clashElements)
            {
                try
                {
                    // Create a SelectionSet from the clash elements
                    SelectionSet selSet = new SelectionSet(clashElements);
                    selSet.DisplayName = clash.DisplayName; // Same name as viewpoint

                    // Find or create the folder
                    GroupItem folder = null;
                    foreach (SavedItem item in doc.SelectionSets.Value)
                    {
                        if (item is GroupItem gi && gi.DisplayName == "Clash Section Sets")
                        {
                            folder = gi;
                            break;
                        }
                    }

                    if (folder == null)
                    {
                        FolderItem newFolder = new FolderItem { DisplayName = "Clash Section Sets" };
                        doc.SelectionSets.AddCopy(newFolder);
                        folder = doc.SelectionSets.Value.Last() as GroupItem;
                    }

                    int idx = folder.Children.Count;
                    doc.SelectionSets.InsertCopy(folder, idx, selSet);
                    //DebugLog.Log($"  SelectionSet created: '{clash.DisplayName}' ({clashElements.Count} items)");
                }
                catch (Exception ex)
                {
                    //DebugLog.Log($"  SelectionSet FAILED: {ex.Message}");
                }
            }

            private static void AddRedlineCircleToActiveView()
            {
                try
                {
                    var activeView = Autodesk.Navisworks.Api.Application.ActiveDocument.ActiveView;

                    // Draw a circle centered on screen (0,0) with ~10% screen radius
                    double size = 0.10;
                    string s = size.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                    string ns = (-size).ToString("R", System.Globalization.CultureInfo.InvariantCulture);

                    string json = "{\"Type\":\"RedlineCollection\",\"Version\":1," +
                        "\"Values\":[{\"Type\":\"RedlineEllipse\",\"Version\":1," +
                        "\"Thickness\":3," +
                        "\"Color\":[1,0,0]," +
                        "\"MinPoint\":[" + ns + "," + ns + "]," +
                        "\"MaxPoint\":[" + s + "," + s + "]}]}";

                    activeView.SetRedlines(json);
                    //DebugLog.Log($"  Redline circle set on ActiveView");
                }
                catch (Exception ex)
                {
                    //DebugLog.Log($"  Redline FAILED: {ex.Message}");
                }
            }

            public static GroupItem FindNetFolder(Document doc, string folderName)
            {
                foreach (SavedItem item in doc.SavedViewpoints.Value)
                {
                    if (item is GroupItem gi && gi.DisplayName == folderName)
                        return gi;
                }
                return null;
            }

            public static List<ClashResult> GetFilteredClashResults(ClashTest test)
            {
                List<ClashResult> results = new List<ClashResult>();
                CollectClashResults(test.Children, results);
                return results;
            }

            private static void SetSectionPlaneAtClash(double clashZ, bool enable)
            {
                try
                {
                    var activeView = Autodesk.Navisworks.Api.Application.ActiveDocument.ActiveView;

                    if (enable)
                    {
                        // Distance must be NEGATIVE of the Z coordinate
                        double distance = -clashZ;
                        string distStr = distance.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

                        string json = "{\"Type\":\"ClipPlaneSet\",\"Version\":1," +
                            "\"Planes\":[{\"Type\":\"ClipPlane\",\"Version\":1," +
                            "\"Normal\":[0,0,-1]," +
                            "\"Distance\":" + distStr + "," +
                            "\"Enabled\":true}]," +
                            "\"Linked\":false,\"Enabled\":true}";

                        activeView.SetClippingPlanes(json);
                        //DebugLog.Log($"  SetClippingPlanes: Plane1 at Z={clashZ:F3}, Distance={distance:F3}, Enabled=true");
                    }
                    else
                    {
                        string json = "{\"Type\":\"ClipPlaneSet\",\"Version\":1," +
                            "\"Planes\":[]," +
                            "\"Linked\":false,\"Enabled\":false}";

                        activeView.SetClippingPlanes(json);
                        //DebugLog.Log($"  SetClippingPlanes: Cleared");
                    }
                }
                catch (Exception ex)
                {
                    //DebugLog.Log($"  SetSectionPlaneAtClash FAILED: {ex.Message}");
                }
            }

            private static void CollectClashResults(SavedItemCollection items, List<ClashResult> results)
            {
                foreach (SavedItem item in items)
                {
                    if (item is ClashResult cr)
                    {
                        if (cr.Status == ClashResultStatus.New ||
                            cr.Status == ClashResultStatus.Active ||
                            cr.Status == ClashResultStatus.Reviewed)
                        {
                            results.Add(cr);
                        }
                    }
                    else if (item is ClashResultGroup group)
                    {
                        CollectClashResults(group.Children, results);
                    }
                }
            }

            public static GroupItem CreateViewPointFolder(Document doc, string folderName)
            {
                foreach (SavedItem item in doc.SavedViewpoints.Value)
                {
                    if (item is GroupItem existing && existing.DisplayName == folderName)
                        return existing;
                }

                FolderItem folder = new FolderItem { DisplayName = folderName };
                doc.SavedViewpoints.AddCopy(folder);

                SavedItem added = doc.SavedViewpoints.Value[doc.SavedViewpoints.Value.Count - 1];
                if (added is GroupItem groupItem) return groupItem;

                throw new InvalidOperationException("Failed to create viewpoint folder.");
            }            

            public static void CreateViewPointForClash(
                Document doc, ClashResult clash, InwOpState10 comState)
            {
                if (clash == null) return;

                if (clash.Center == null ||
                    double.IsNaN(clash.Center.X) ||
                    double.IsNaN(clash.Center.Y) ||
                    double.IsNaN(clash.Center.Z))
                {
                    //DebugLog.Log($"Skipping '{clash.DisplayName}': Invalid center.");
                    return;
                }

                //DebugLog.Log($"Processing '{clash.DisplayName}' [{clash.Status}]");

                // 1. Collect elements
                ModelItemCollection clashElements = new ModelItemCollection();
                if (clash.Item1 != null) clashElements.Add(clash.Item1);
                if (clash.Item2 != null) clashElements.Add(clash.Item2);

                if (clashElements.Count == 0)
                {
                    //DebugLog.Log($"  Skipping: No elements.");
                    return;
                }

                Point3D clashPoint = clash.Center;
                var activeView = Autodesk.Navisworks.Api.Application.ActiveDocument.ActiveView;


                // 2. Highlight red
                doc.Models.OverridePermanentColor(clashElements, new NavisColor(1.0, 0.0, 0.0));

                // 3. Select and zoom
                doc.CurrentSelection.CopyFrom(clashElements);
                comState.ZoomInCurViewOnCurSel();
                //DebugLog.Log($"  Zoomed.");

                // 4. Section plane using json
                SetSectionPlaneAtClash(clashPoint.Z, true);

                // 5. Add redline circle
                //AddRedlineCircleToActiveView();
                activeView.SetRedlines("{\"Type\":\"RedlineCollection\",\"Version\":1,\"Values\":[]}");

                //Create matching selection set
                CreateSelectionSetForClash(doc, clash, clashElements);


                // 6. Save viewpoint via .NET
                Viewpoint currentViewpoint = doc.CurrentViewpoint.ToViewpoint();
                SavedViewpoint savedVP = new SavedViewpoint(currentViewpoint);
                savedVP.DisplayName = clash.DisplayName;

                

                GroupItem netFolder = FindNetFolder(doc, "Clash Section Views");
                if (netFolder != null)
                {
                    int idx = netFolder.Children.Count;
                    doc.SavedViewpoints.InsertCopy(netFolder, idx, savedVP);


                    // ReplaceFromCurrentView captures redlines + clip planes into the saved VP
                    SavedItem lastAdded = netFolder.Children.ElementAt(idx);
                    if (lastAdded is SavedViewpoint svp)
                        doc.SavedViewpoints.ReplaceFromCurrentView(svp);

                    //DebugLog.Log($"  Saved (index {idx})");
                }
                else
                {
                    doc.SavedViewpoints.AddCopy(savedVP);
                    //DebugLog.Log($"  Saved to root");
                }


                // 7. Reset

                
                SetSectionPlaneAtClash(0, false);

                doc.Models.ResetPermanentMaterials(clashElements);
                doc.CurrentSelection.Clear();

                //DebugLog.Save();
            }
        }
    }
}