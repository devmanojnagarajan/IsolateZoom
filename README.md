SectionPlaneZoom for Autodesk Navisworks
<img width="1918" height="118" alt="image" src="https://github.com/user-attachments/assets/397e5c06-2df7-4410-a7ea-49a7e7473261" />

SectionPlaneZoom is a Navisworks Add-in designed to streamline the clash review process. It automatically iterates through clash test results, creates a section plane at the exact clash elevation, zooms into the conflicting elements, and saves a dedicated viewpoint for easier coordination.

<img width="341" height="148" alt="image" src="https://github.com/user-attachments/assets/a1081c8c-a1ef-42ba-a0eb-792dc75b4b7f" />

🚀 Key Features
* Automatic Folder Management: Creates organized viewpoint folders ("Clash Section Views") and Selection Sets ("Clash Section Sets").
* Smart Sectioning: Uses the Navisworks JSON API to generate horizontal section planes at the precise $Z$ coordinate of the clash.
* Visual Highlights: Temporarily overrides element colors to red during the generation process for clear identification.
* Selection Set Sync: Generates a matching Selection Set for every viewpoint so you can instantly select the clashing items later.
* Progress Tracking: Includes a real-time progress bar with a cancel option for large clash tests.

<img width="180" height="450" alt="image" src="https://github.com/user-attachments/assets/294c84eb-6093-4be3-9933-5a0b9ccd63d2" />

🛠️ How It Works (Technical Overview)
The plugin leverages both the .NET API and the COM API Bridge to achieve results that standard .NET methods cannot easily handle:

Selection & Zoom: It identifies the clashing ModelItems and uses comState.ZoomInCurViewOnCurSel() for a perfect frame.

JSON Clipping: It bypasses traditional COM limitations by injecting a ClipPlaneSet via JSON:

JSON
{ "Type": "ClipPlaneSet", "Planes": [{ "Normal": [0,0,-1], "Distance": -clashZ }] }
Viewpoint Persistence: It uses ReplaceFromCurrentView to ensure that the active clipping planes and redlines are baked into the saved viewpoint.

Clean Up: After saving, it resets the model materials and clears the selection to leave the workspace ready for the next task.

<img width="1913" height="298" alt="image" src="https://github.com/user-attachments/assets/16271086-a1c6-4463-a5fa-4eebf52ebaa7" />

📖 How to Use
1. Run the Plugin
In Navisworks, navigate to the Tool Add-ins tab and click on SectionPlaneZoom.

2. Select Your Clash Test
If your file contains multiple clash tests, a dialog box will appear. Select the specific test you wish to process and click OK.

3. Processing
The plugin will begin iterating through all New, Active, and Reviewed clashes. A progress bar will show you the current status and the name of the clash being processed.

4. Review Results
Once complete, check your Saved Viewpoints window and Selection Sets window.

Viewpoints: You will find a new folder called Clash Section Views. Every viewpoint is sectioned horizontally at the clash point.

Selection Sets: A corresponding folder Clash Section Sets will contain the items involved in each clash for quick selection.

<img width="1338" height="976" alt="image" src="https://github.com/user-attachments/assets/b61d2b9f-d209-423b-90b3-3c348323203c" />

📋 Requirements
Autodesk Navisworks Manage (2018 or newer recommended).
A default view will all elements active and section plane turned off.
Clash Detective data must be present in the active document.

<img width="1342" height="970" alt="image" src="https://github.com/user-attachments/assets/e85c135c-c94a-4f8d-9982-2db4c57f4a92" />
📂 Installation
To install the plugin manually, you need to place the compiled .dll (and any dependencies) into the Navisworks Plugins directory.

1. Locate the Plugins Folder
Navisworks looks for plugins in the %AppData% directory. Paste the following path into your File Explorer address bar: %AppData%\Autodesk\Navisworks Manage 2024\Plugins\

Note: Replace 2024 with your specific version (e.g., 2023, 2025).

2. Create a Subfolder
Inside the Plugins folder, create a new folder named exactly SectionPlaneZoom.

Important: The folder name must match the name of your .dll file for Navisworks to load it correctly.

3. Copy the DLL
Place your SectionPlaneZoom.dll into that new folder.
