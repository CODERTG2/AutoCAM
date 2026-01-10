# AutoCAM
AutoCAM is a Fusion 360 script designed to automate Computer-Aided Manufacturing (CAM) operations for milling. It simplifies the setup and toolpath generation process for specific materials like Polycarbonate and Aluminum.

## Features

- **Automated Setup Creation**: Automatically creates a new milling setup with predefined stock and WCS settings.
- **Material-Based Configuration**: Prompts the user to select between "Polycarb" and "Aluminum" to automatically configure:
  - Spindle Speed
  - Plunge Feed Rate
  - Depth Passes
  - Coolant settings (Disabled by default)
- **Hole Detection & Boring**: Automatically detects cylindrical holes in the model and creates **Bore** operations for them.
- **2D Contour Generation**: Identifies the bottom face of the model and creates a **2D Contour** operation for the outer profile.
- **Manual NC Stop**: Inserts a manual NC stop command between operations.
- **Tool Selection**: Attempts to use the first tool available in the document's tool library, or falls back to a default 3.175mm flat end mill.

## Installation & Setup

To use this script in Autodesk Fusion 360, follow these steps:

### 1. Locate the Scripts Directory
Fusion 360 looks for scripts in specific directories on your computer.

*   **Mac OS**: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/Scripts/`
*   **Windows**: `%appdata%\Autodesk\Autodesk Fusion 360\API\Scripts\`

### 2. Install the Script
1.  Navigate to the `Scripts` directory mentioned above.
2.  Create a new folder named `AutoCAM` (if it doesn't already exist).
3.  Copy the following files into this folder:
    *   `AutoCAM.py`
    *   `AutoCAM.manifest`
    *   `ScriptIcon.svg`

### 3. Run the Script in Fusion 360
1.  Open **Autodesk Fusion 360**.
2.  Open a design file that has a 3D model (bodies) ready for manufacturing.
3.  Switch to the **Design** or **Manufacture** workspace.
4.  Go to the **Utilities** tab (in Design) or **Add-ins** tab (in Manufacture).
5.  Click on **Scripts and Add-Ins** (or press `Shift + S`).
6.  Click on the **Scripts** tab.
7.  Scroll down to find **AutoCAM-MVP** under "My Scripts".
8.  Select it and click **Run**.

## Usage

1.  Upon running the script, a prompt will appear asking: "Is this polycarb or aluminum?".
2.  Type either `polycarb` or `aluminum` and click OK.
3.  The script will:
    *   Create a Setup.
    *   Detect holes and creating a Boring operation.
    *   Add a Manual NC Stop.
    *   Create a 2D Contour operation for the part outline.
    *   Generate toolpaths.
4.  Check the **Browser** in the Manufacture workspace to see the generated operations.

## Requirements

*   Autodesk Fusion 360
*   A design with valid BRep bodies (3D geometry).
