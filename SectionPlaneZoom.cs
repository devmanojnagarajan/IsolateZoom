using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
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
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using wf = System.Windows.Forms;
using NavisColor = Autodesk.Navisworks.Api.Color;
using BF = System.Reflection.BindingFlags;


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

            DebugLog.Reset();
            DebugLog.Log("=== Clash VP - Section Started ===");

            DocumentClash clashDoc = doc.GetClash();
            if (clashDoc == null || clashDoc.TestsData?.Tests == null || clashDoc.TestsData.Tests.Count == 0)
            {
                MessageBox.Show("No Clash Test Detected");
                return 0;
            }

            var selectedTest = ClashHelper.ClashTestSelect(clashDoc.TestsData.Tests);
            if (selectedTest == null) return 0;

            DebugLog.Log($"Selected test: {selectedTest.DisplayName}");

            List<ClashResult> results = ClashHelper.GetFilteredClashResults(selectedTest);
            DebugLog.Log($"Found {results.Count} clashes with status New/Active/Reviewed");

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
            ComHelper.DiscoverClipPlaneSetMethods(doc);



            // === COM API Discovery ===
            DebugLog.Log("--- COM API Discovery ---");
            try
            {
                object view = comState.CurrentView;
                DebugLog.Log($"CurrentView type: {view.GetType().FullName}");

                foreach (Type iface in view.GetType().GetInterfaces())
                    DebugLog.Log($"  Interface: {iface.FullName}");

                object cp = view.GetType().InvokeMember("ClippingPlanes",
                    BF.InvokeMethod, null, view, null);
                DebugLog.Log($"ClippingPlanes type: {cp.GetType().FullName}");

                foreach (Type iface in cp.GetType().GetInterfaces())
                    DebugLog.Log($"  CP Interface: {iface.FullName}");
            }
            catch (Exception dEx)
            {
                DebugLog.Log($"Discovery error: {dEx.Message}");
            }
            DebugLog.Log("--- End Discovery ---");
            DebugLog.Save();

            // Create/find the viewpoint folder
            GroupItem folder = ClashHelper.CreateViewPointFolder(doc, "Clash Section Views");
            DebugLog.Log($"Folder ready: '{folder.DisplayName}', existing children: {folder.Children.Count}");
            DebugLog.Save();

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
                        DebugLog.Log($"EXCEPTION: {detail}");
                        DebugLog.Save();
                    }
                }


                DebugLog.Log($"Loop complete: {successCount} succeeded, {failedClashes.Count} failed");

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

            DebugLog.Log("=== Finished ===");
            DebugLog.Save();
            return 0;
        }


        // ──────────────────────────────────────────────
        // COM Helper — all COM access via reflection
        // ──────────────────────────────────────────────
        public static class ComHelper
        {
            public static object Invoke(object obj, string method, params object[] args)
            {
                return obj.GetType().InvokeMember(method, BF.InvokeMethod, null, obj, args);
            }

            public static object Get(object obj, string prop)
            {
                return obj.GetType().InvokeMember(prop, BF.GetProperty, null, obj, null);
            }

            

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
                    DebugLog.Log($"  DisableSectioningNet failed: {ex.Message}");
                }
            }

            public static Viewpoint SetSectionPlaneOnViewpoint(Document doc, Viewpoint sourceVp, double z)
            {
                try
                {
                    // Create a copy of the viewpoint that we'll modify
                    Viewpoint vpCopy = sourceVp.CreateCopy();

                    // Get InternalClipPlanes via reflection
                    Type vpType = vpCopy.GetType();
                    PropertyInfo internalProp = vpType.GetProperty("InternalClipPlanes",
                        BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

                    if (internalProp == null)
                    {
                        DebugLog.Log("  InternalClipPlanes not found");
                        return vpCopy;
                    }

                    dynamic clipPlaneSet = internalProp.GetValue(vpCopy);

                    // Enable sectioning and link to viewpoint
                    clipPlaneSet.SetEnabled(true);
                    clipPlaneSet.SetLinked(true);

                    DebugLog.Log($"  Section ENABLED at Z={z:F3}");

                    return vpCopy;  // Return the modified viewpoint
                }
                catch (Exception ex)
                {
                    DebugLog.Log($"  SetSectionPlaneOnViewpoint error: {ex.Message}");
                    return sourceVp;
                }
            }

            public static void DiscoverClipPlaneSetMethods(Document doc)
            {
                try
                {
                    DebugLog.Log("=== LcOaClipPlaneSet DISCOVERY ===");

                    Viewpoint vp = doc.CurrentViewpoint.ToViewpoint();

                    // Get the internal clip planes via reflection
                    Type vpType = vp.GetType();
                    PropertyInfo internalProp = vpType.GetProperty("InternalClipPlanes",
                        BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

                    if (internalProp != null)
                    {
                        object clipPlaneSet = internalProp.GetValue(vp);
                        DebugLog.Log($"  ClipPlaneSet type: {clipPlaneSet?.GetType().FullName}");

                        if (clipPlaneSet != null)
                        {
                            Type setType = clipPlaneSet.GetType();

                            // List all properties
                            DebugLog.Log("  --- Properties ---");
                            foreach (var prop in setType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                            {
                                try
                                {
                                    object val = prop.GetValue(clipPlaneSet);
                                    DebugLog.Log($"    {prop.Name} = {val}");
                                }
                                catch { DebugLog.Log($"    {prop.Name} = (error)"); }
                            }

                            // List all methods
                            DebugLog.Log("  --- Methods ---");
                            foreach (var method in setType.GetMethods(BindingFlags.Public | BindingFlags.Instance))
                            {
                                if (!method.Name.StartsWith("get_") && !method.Name.StartsWith("set_"))
                                {
                                    var parms = method.GetParameters();
                                    string sig = string.Join(", ", parms.Select(p => p.ParameterType.Name));
                                    DebugLog.Log($"    {method.Name}({sig})");
                                }
                            }
                        }
                    }
                    else
                    {
                        DebugLog.Log("  InternalClipPlanes property not found");
                    }

                    // Also check ActiveView for sectioning toggle
                    DebugLog.Log("=== ActiveView DISCOVERY ===");
                    var activeView = doc.ActiveView;
                    Type avType = activeView.GetType();
                    foreach (var prop in avType.GetProperties())
                    {
                        if (prop.Name.ToLower().Contains("clip") ||
                            prop.Name.ToLower().Contains("section") ||
                            prop.Name.ToLower().Contains("plane"))
                        {
                            try
                            {
                                object val = prop.GetValue(activeView);
                                DebugLog.Log($"  ActiveView.{prop.Name} = {val}");
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception e)
                {
                    DebugLog.Log($"DiscoverClipPlaneSetMethods failed: {e.Message}");
                }
            }

            public static void DiscoverViewpointClipping(Document doc)
            {
                try
                {
                    DebugLog.Log("=== VIEWPOINT CLIPPING PROPERTIES ===");

                    Viewpoint vp = doc.CurrentViewpoint.ToViewpoint();
                    Type vpType = vp.GetType();

                    foreach (var prop in vpType.GetProperties())
                    {
                        if (prop.Name.ToLower().Contains("clip") ||
                            prop.Name.ToLower().Contains("section") ||
                            prop.Name.ToLower().Contains("plane"))
                        {
                            try
                            {
                                object val = prop.GetValue(vp);
                                DebugLog.Log($"  {prop.Name} = {val}");
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception e)
                {
                    DebugLog.Log($"DiscoverViewpointClipping failed: {e.Message}");
                }
            }

            public static void DiscoverClippingPlanes(InwOpState10 comState)
            {
                try
                {
                    object view = comState.CurrentView;
                    object clipPlanes = Invoke(view, "ClippingPlanes");

                    DebugLog.Log("=== CLIPPING PLANES MEMBERS ===");

                    // Try common property/method names
                    string[] names = { "Count", "Enabled", "Linked", "Current", "Planes",
                          "AddPlane", "RemovePlane", "Add", "Remove", "Clear",
                          "Item", "GetPlane", "CreatePlane", "Mode" };

                    foreach (string name in names)
                    {
                        try
                        {
                            object result = Get(clipPlanes, name);
                            DebugLog.Log($"  Property '{name}' = {result}");
                        }
                        catch
                        {
                            try
                            {
                                object result = Invoke(clipPlanes, name);
                                DebugLog.Log($"  Method '{name}()' = {result}");
                            }
                            catch { }
                        }
                    }

                    // Also try to create a clip plane object and discover its members
                    try
                    {
                        object plane = comState.ObjectFactory(nwEObjectType.eObjectType_nwOaClipPlane, null, null);
                        DebugLog.Log($"ClipPlane created: {plane.GetType().FullName}");

                        string[] planeNames = { "Plane", "Enabled", "Distance", "Normal", "Equation", "SetValue", "GetValue" };
                        foreach (string name in planeNames)
                        {
                            try
                            {
                                object result = Get(plane, name);
                                DebugLog.Log($"  Plane.{name} = {result}");
                            }
                            catch { }
                        }
                        Marshal.ReleaseComObject(plane);
                    }
                    catch (Exception e)
                    {
                        DebugLog.Log($"  ClipPlane creation failed: {e.Message}");
                    }
                }
                catch (Exception e)
                {
                    DebugLog.Log($"DiscoverClippingPlanes failed: {e.Message}");
                }
            }

            public static void EnableSectioningForAllViewpoints(Document doc, string folderName)
            {
                try
                {
                    DebugLog.Log("=== Enabling Sectioning for All Saved Viewpoints ===");

                    GroupItem folder = ClashHelper.FindNetFolder(doc, folderName);
                    if (folder == null) return;

                    foreach (SavedItem item in folder.Children)
                    {
                        if (item is SavedViewpoint savedVp)
                        {
                            // Load the viewpoint
                            doc.CurrentViewpoint.CopyFrom(savedVp.Viewpoint);

                            // Get current viewpoint and enable sectioning on it
                            Viewpoint vp = doc.CurrentViewpoint.ToViewpoint().CreateCopy();
                            Type vpType = vp.GetType();
                            PropertyInfo internalProp = vpType.GetProperty("InternalClipPlanes",
                                BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

                            if (internalProp != null)
                            {
                                dynamic clipPlaneSet = internalProp.GetValue(vp);
                                clipPlaneSet.SetEnabled(true);
                                clipPlaneSet.SetLinked(true);

                                // Apply back and save
                                doc.CurrentViewpoint.CopyFrom(vp);
                            }

                            DebugLog.Log($"  Enabled: {savedVp.DisplayName}");
                        }
                    }

                    DebugLog.Log("=== Sectioning Enabled ===");
                }
                catch (Exception ex)
                {
                    DebugLog.Log($"EnableSectioningForAllViewpoints error: {ex.Message}");
                }
            }
            public static void SaveViewpointWithSection(InwOpState10 comState, string name)
            {
                try
                {
                    object savedViewsColl = comState.SavedViews();
                    object currentView = Get(comState, "CurrentView");

                    // Create a copy of current view state (includes clipping)
                    object newViewpoint = Invoke(savedViewsColl, "CreateCopy", currentView);
                    Set(newViewpoint, "Name", name);

                    Invoke(savedViewsColl, "Add", newViewpoint);

                    DebugLog.Log($"  COM Viewpoint saved: {name}");
                }
                catch (Exception e)
                {
                    DebugLog.Log($"  COM SaveViewpoint failed: {e.Message}");
                }
            }
            public static void Set(object obj, string prop, object value)
            {
                obj.GetType().InvokeMember(prop, BF.SetProperty, null, obj, new[] { value });
            }



            public static void SetSectionPlaneNet(Document doc, double z)
            {
                try
                {
                    // Get the current viewpoint and make a copy we can modify
                    Viewpoint vp = doc.CurrentViewpoint.ToViewpoint().CreateCopy();

                    // Access InternalClipPlanes via reflection
                    Type vpType = vp.GetType();
                    PropertyInfo internalProp = vpType.GetProperty("InternalClipPlanes",
                        BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

                    if (internalProp == null)
                    {
                        DebugLog.Log("  InternalClipPlanes not found");
                        return;
                    }

                    dynamic clipPlaneSet = internalProp.GetValue(vp);

                    // Enable sectioning and link to viewpoint
                    clipPlaneSet.SetEnabled(true);
                    clipPlaneSet.SetLinked(true);

                    // Apply the modified viewpoint back
                    doc.CurrentViewpoint.CopyFrom(vp);

                    DebugLog.Log($"  Section ENABLED at Z={z:F3}");
                }
                catch (Exception ex)
                {
                    DebugLog.Log($"  SetSectionPlaneNet error: {ex.Message}");
                }
            }

            public static void DisableSectioning(InwOpState10 comState)
            {
                try
                {
                    object view = comState.CurrentView;
                    object clipPlanes = Invoke(view, "ClippingPlanes");
                    Set(clipPlanes, "Enabled", false);

                    int count = (int)Get(clipPlanes, "Count");
                    while (count > 0)
                    {
                        Invoke(clipPlanes, "RemovePlane", 1);
                        count = (int)Get(clipPlanes, "Count");
                    }
                }
                catch (Exception e)
                {
                    DebugLog.Log($"DisableSectioning failed: {e.Message}");
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
                    DebugLog.Log($"Skipping '{clash.DisplayName}': Invalid center.");
                    return;
                }

                DebugLog.Log($"Processing '{clash.DisplayName}' [{clash.Status}]");

                // 1. Collect elements
                ModelItemCollection clashElements = new ModelItemCollection();
                if (clash.Item1 != null) clashElements.Add(clash.Item1);
                if (clash.Item2 != null) clashElements.Add(clash.Item2);

                if (clashElements.Count == 0)
                {
                    DebugLog.Log($"  Skipping: No elements.");
                    return;
                }

                Point3D clashPoint = clash.Center;

                // 2. Highlight red
                doc.Models.OverridePermanentColor(clashElements, new NavisColor(1.0, 0.0, 0.0));

                // 3. Select and zoom
                doc.CurrentSelection.CopyFrom(clashElements);
                comState.ZoomInCurViewOnCurSel();
                DebugLog.Log($"  Zoomed.");

                // 4. Section plane - apply to viewpoint we're about to save
                Viewpoint currentViewpoint = doc.CurrentViewpoint.ToViewpoint();
                Viewpoint vpWithSection = null;
                try
                {
                    vpWithSection = ComHelper.SetSectionPlaneOnViewpoint(doc, currentViewpoint, clashPoint.Z);
                }
                catch (Exception secEx)
                {
                    DebugLog.Log($"  Section plane FAILED: {secEx.Message}");
                    vpWithSection = currentViewpoint; // fallback to original
                }

                // 5. Save the MODIFIED viewpoint
                SavedViewpoint savedVP = new SavedViewpoint(vpWithSection);  // <-- use vpWithSection, not currentViewpoint
                savedVP.DisplayName = clash.DisplayName;

                GroupItem netFolder = FindNetFolder(doc, "Clash Section Views");
                if (netFolder != null)
                {
                    int idx = netFolder.Children.Count;
                    doc.SavedViewpoints.InsertCopy(netFolder, idx, savedVP);
                    DebugLog.Log($"  Saved (index {idx})");
                }
                else
                {
                    doc.SavedViewpoints.AddCopy(savedVP);
                    DebugLog.Log($"  Saved to root");
                }

                // 6. Reset
                //doc.Models.ResetPermanentMaterials(clashElements);
                //doc.CurrentSelection.Clear();

                DebugLog.Save();
            }
        }
    }
}