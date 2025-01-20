# Batch Word Table Shaded Color Change Macro Guide 📄
A comprehensive guide for using the VBA macro to automatically change non-white cell colors in Microsoft Word tables across multiple documents.

## 📋 Table of Contents
1. [Step-by-Step Guide](#step-by-step-guide)
2. [Prerequisites](#prerequisites)
3. [Technical Requirements](#technical-requirements)
4. [Troubleshooting](#troubleshooting)
5. [Security](#security)
6. [Support](#support)

## 🎯 Step-by-Step Guide

### 1. Organize Your Documents
* Create a dedicated folder for your Word documents
* Move all documents you want to process into this folder
* Ensure all documents are closed before proceeding

### 2. Enable the Developer Tab
* Open Microsoft Word
* Go to **File** > **Options** > **Customize Ribbon**
* Check the box next to **Developer** in the right column
* Click **OK** to save changes

### 3. Open the VBA Editor
* Press **Alt + F11** on your keyboard
* This will open the Microsoft Visual Basic for Applications editor

### 4. Insert a New Module
* In the VBA editor, click **Insert** > **Module**
* A new module window will appear

### 5. Paste the VBA Code
* Check the code in the script.bas

### 6. Update the Folder Path
* Locate the `folderPath = "C:\Path\To\Your\Folder\"` line in the code
* Replace it with your actual folder path
* Example: `folderPath = "C:\Users\YourName\Documents\WordFiles\"`
* Ensure the path ends with a backslash `\`

### 7. Change the color
* Change the color : `cell.Shading.BackgroundPatternColor = wdColorBlue`
* Color Reference : https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor

### 8. Run the Script
* Click **Run** the script from the VBA Editor

### 9. Check the Results
* Wait for the completion message
* Open a processed document to verify the changes
* All non-white shaded cells should now be blue

## 🛠️ Technical Requirements
- Microsoft Word (any modern version)
- Developer tab enabled
- Macro security settings configured
- Write permissions for documents
- All documents must be .docx format

## ⚠️ Important Notes
- **Always backup your documents before running the macro**
- Ensure all target documents are closed
- The macro processes all .docx files in the specified folder
- Changes are saved automatically
- The process cannot be undone after saving

## 🔍 Troubleshooting
If you encounter errors:
- Verify the folder path is correct
- Check file permissions
- Ensure documents aren't open
- Verify macro security settings
- Check document format (.docx only)

## 🔒 Security
- Only run macros from trusted sources
- Keep document backups
- Check organization's macro policies
- Review code before running

## 📞 Support
For assistance:
- Please write the problem in issue

---
Happy document processing! 🎉 Remember to always backup your files before running batch operations!