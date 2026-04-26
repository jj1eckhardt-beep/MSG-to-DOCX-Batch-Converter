# THE ARCHIVIST 
### Universal Outlook Email Archival Suite:

Is a high-integrity archival tool designed to transform Outlook `.msg` files into permanent, searchable, and sortable digital libraries.

A lightweight, free automation tool to batch convert Microsoft Outlook `.msg` files into universal `.html`, modern Word `.docx`, or plain text documents while handling any attached files as selected. 

This project was born out of frustration with finding only paid, third-party software to extract and convert old archived `.msg` files. This solution is 100% free and runs locally on your machine.

## 🚀 Key Features 
 **Triple-Output Engine:** Convert to Portable HTML, Microsoft Word (with OLE PDF embedding), or Lite Plain Text.
 
 **Smart Handshake:** Automatically connects to active Outlook/Word sessions to prevent "Server Execution" errors.
 
 **The Auditor:** A standalone indexing engine that crawls your archive to build a dynamic `INDEX.html`.
 
 **Interactive Dashboard:** Real-time stats including % Done, Items/sec, and Time Remaining.
 
 **Universal Attachment Handling:** Automatically extracts or embeds PDFs, XLSX, and DOCX files.
 
 **Deep Audit Mode:** Peeks into original source files to recover true "Sent Date" and "Sender" metadata.

   
## 📊 The Master Index (Audit Report)
  The crown jewel of The Archivist is the automated Index.html. This isn't just a list of files; it's a dynamic, dark-mode dashboard for your digital library.
 
 **Click-to-Sort Headers:** Instantly organize your entire archive by Date, Subject, or Sender with a single click—no page reload required.
 
 **Deep Audit Metadata:** Unlike standard file explorers, the Index peeks into the original email headers to display the True Sent Date and Sender Name, even if the files were moved or renamed.
 
 **Visual Status Bar:** High-visibility Cyan stats bar at the top displays the original source path, the archive destination, and the total item count.
 
 **Paperclip Quick-Links:** For HTML archives, paperclip icons (📎) provide one-click access to Email parent folder containing .html and attachments.  Optimized for high-integrity text and attachment archival (perfect for students, educators, and project managers).
 
 **Universal Compatibility:** Self-contained and ultra-portable. The Index uses pure HTML/JS, meaning it works in any browser, on any device, forever—no internet connection or 

## 📋 Prerequisites
 Because this script currently uses COM object automation to tap into native desktop applications, you must have the following installed on your Windows machine:
 **Microsoft Outlook** (Desktop application)
 
 **Microsoft Word** (Desktop application)
 
 **PowerShell** (Built into Windows)

*Note: Outlook and Word are only required on the host machine performing the conversion. (For Now) The finished documents can be opened on any computer using free software like LibreOffice!*

## 🚀 How to Use

## 1. Repository Setup
1. Download files from this repository, (the `Latest.ps1` script and the `Launcher.bat` file) and place them into a folder.  
Or, download the .zip file and extract to a location of your choice.  
2. I encourage you to open the files in Notepad and inspect the code first before executing.

## 2. Open the UI
Open the Launcher.bat file (It handles permissions and launching the UI). 

## 3. Execute
1.	Use MASTER and ARCHIVE buttons to set the desired folder paths.
2.	Use Radio buttons and Checkboxes to configure the desired output format and selections.
3.	Start Archival Process, ABORT will pause, Start will resume and continue until finished.
4.	Reset clears selections.
   
## 🖥️ System Screenshots
UI Paths Set and Ready.
<img width="1777" height="888" alt="image" src="https://github.com/user-attachments/assets/00a9d78f-36be-499d-9c46-e9aa9e8a4dfb" />
Index.html
<img width="1717" height="1247" alt="image" src="https://github.com/user-attachments/assets/96f3a260-6c53-4570-8a34-fd6acb99a594" />
HTML and Docx Output with Headers side by side.
<img width="1824" height="950" alt="image" src="https://github.com/user-attachments/assets/3581ccc2-c913-4197-9087-6ef84e129b22" />

## 🖼️ Dashboard in Action
Processing...
<img width="1708" height="866" alt="image" src="https://github.com/user-attachments/assets/fbbfa6ac-a877-4c17-9dca-642efcf95f51" />
<img width="1000" height="849" alt="image" src="https://github.com/user-attachments/assets/5dd0a938-d978-4b87-ab69-7702ddbae85a" />
Done.
<img width="992" height="882" alt="image" src="https://github.com/user-attachments/assets/9c26bcd3-ddaa-43d6-acef-d63dafca69fc" />

## ⚖️ License & Disclaimer
This project is licensed under the MIT License.

CAUTION: This software is still a work in progress and under development. Always backup your Emails. The author is not responsible for data loss or hardware issues.

## ☕ Support the Project
If "THE ARCHIVIST" saved you some time or made your conversion easier, feel free to help keep the gears turning!

* [**Support via Ko-fi**](https://ko-fi.com/kofisupporter19535)
* [**Support via Buy Me a Coffee**](https://www.buymeacoffee.com/jj1eckhardt)

