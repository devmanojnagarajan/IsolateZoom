# SectionPlaneZoom for Autodesk Navisworks

SectionPlaneZoom is a Navisworks add-in that automates clash review visualization.  
It iterates through clash test results, creates a horizontal section plane at the exact clash elevation, zooms into the conflicting elements, and saves dedicated viewpoints and selection sets for faster coordination.

---

## 📸 Overview

<img width="1918" height="118" alt="Overview" src="https://github.com/user-attachments/assets/397e5c06-2df7-4410-a7ea-49a7e7473261" />

<img width="341" height="148" alt="UI" src="https://github.com/user-attachments/assets/a1081c8c-a1ef-42ba-a0eb-792dc75b4b7f" />

---

## 🚀 Key Features

### 📁 Automatic Folder Management
- Creates structured Saved Viewpoints folder: **Clash Section Views**
- Creates structured Selection Sets folder: **Clash Section Sets**

### ✂️ Smart Sectioning
- Generates horizontal section planes at precise clash **Z elevation**
- Uses Navisworks JSON API for accurate clipping

### 🎯 Visual Highlights
- Temporarily overrides clash elements to **red**
- Improves visual clarity during generation

### 🔗 Selection Set Sync
- Creates a matching Selection Set for every viewpoint
- Allows instant re-selection of clash elements

### 📊 Progress Tracking
- Real-time processing progress bar
- Displays current clash name
- Includes cancel option for large clash tests

<img width="180" height="450" alt="Progress UI" src="https://github.com/user-attachments/assets/294c84eb-6093-4be3-9933-5a0b9ccd63d2" />

---

## 🛠️ Technical Overview

The plugin uses a hybrid approach combining the **Navisworks .NET API** and the **COM API Bridge** to achieve advanced functionality.

### 🔍 Selection & Zoom
Identifies clashing `ModelItems` and frames them precisely using:

```
comState.ZoomInCurViewOnCurSel()
```

### ✂️ JSON Clip Plane Injection
Creates clipping planes using JSON:

```json
{
  "Type": "ClipPlaneSet",
  "Planes": [
    { "Normal": [0,0,-1], "Distance": -clashZ }
  ]
}
```

### 🧹 Cleanup
After saving viewpoints:
- model materials reset
- selections cleared
- workspace restored

<img width="1913" height="298" alt="Technical Workflow" src="https://github.com/user-attachments/assets/16271086-a1c6-4463-a5fa-4eebf52ebaa7" />

---

## 📖 How to Use

### 1️⃣ Run the Plugin
In Navisworks:

```
Tool Add-ins → SectionPlaneZoom
```

### 2️⃣ Select Clash Test
If multiple clash tests exist:
- A selection dialog appears
- Choose the desired test
- Click OK

### 3️⃣ Processing
The plugin processes:
- New clashes
- Active clashes
- Reviewed clashes

A progress window shows:
- current clash
- processing status

### 4️⃣ Review Results

#### Saved Viewpoints
```
Clash Section Views
```
Each viewpoint is sectioned at clash elevation.

#### Selection Sets
```
Clash Section Sets
```
Each set contains clash elements for quick selection.

<img width="1338" height="976" alt="Results" src="https://github.com/user-attachments/assets/b61d2b9f-d209-423b-90b3-3c348323203c" />

---

## 📋 Requirements
- Autodesk Navisworks Manage (2018 or newer recommended)
- Default view with:
  - all elements active
  - section planes disabled
- Clash Detective data present in the active document

<img width="1342" height="970" alt="Requirements" src="https://github.com/user-attachments/assets/e85c135c-c94a-4f8d-9982-2db4c57f4a92" />

---

## 📂 Installation

### Step 1 — Locate Plugins Directory
Paste into Windows File Explorer:

```
%AppData%\Autodesk\Navisworks Manage 2024\Plugins\
```

Replace **2024** with your Navisworks version.

### Step 2 — Create Plugin Folder
Create a folder named exactly:

```
SectionPlaneZoom
```

⚠️ Folder name must match the DLL filename.

### Step 3 — Copy Plugin Files
Place inside the folder:

```
SectionPlaneZoom.dll
+ any dependency files
```

Restart Navisworks.

The plugin will appear under:

```
Tool Add-ins
```
