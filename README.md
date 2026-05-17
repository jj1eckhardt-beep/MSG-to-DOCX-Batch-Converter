# THE ARCHIVIST 
### Universal Outlook Email Archival Suite:

Is a high-integrity archival tool designed to transform Outlook `.msg` or `.pst` files into permanent, searchable, and sortable digital libraries.

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
 
 **Universal Compatibility:** Self-contained and ultra-portable. The Index uses pure HTML/JS, meaning it works in any browser, on any device, forever—no internet connection required.

## 📋 Prerequisites
 Because this script currently uses COM object automation to tap into native desktop applications, you must have the following installed on your Windows machine:

 **Windows 10/11**

 **Microsoft Outlook** (Desktop application)
 
 **Microsoft Word** (Desktop application)
 
 **PowerShell 5.1** (Built into Windows)

*Note: Outlook and Word are only required on the host machine performing the conversion. (For Now) The finished documents can be opened on any computer using any browser or free software like LibreOffice!*

## 🚀 How to Use

## 1. Repository Setup
1.	 Download the latest release from the [Releases Page](ttps://github.com/jj1eckhardt-beep/THE-ARCHIVIST/releases). Files needed are The-Archivist.ps1, and Launcher.bat in the same folder.
2. 	I encourage you to open the files in Notepad and inspect the code for yourself first before executing.
3.	 Run `Launcher.bat` to bypass execution policies and clear ghost Office processes.

## 2. Open the UI and make selections
1.	Execute the Launcher.bat file (It handles Admin permissions and launching the UI). 
2. 	Use MASTER and ARCHIVE buttons to set the desired folder paths.	
3.	 Use Radio buttons and Checkboxes to configure the desired output format and selections.

## 3. Execute
1. 	Click Start Archival Process, ABORT will pause, Start Archival Process again will resume and continue until finished.
2. 	Reset clears selections.
3.	After finished, your converted files can now be opened, moved, and viewed on any computer.
   
## 🖥️ System Screenshots
UI Paths Set and Ready.

<img width="1209" height="596" alt="image" src="https://github.com/user-attachments/assets/bee393dc-5ac6-457f-a829-d48ab4b11e9c" />

Index.html

<img width="1168" height="837" alt="image" src="https://github.com/user-attachments/assets/e622468c-42a0-4779-872a-fe9051c9e40b" />

HTML and Docx Output with Headers side by side.

<img width="969" height="557" alt="image" src="https://github.com/user-attachments/assets/ce2bdea3-845d-442e-bb7f-2652ddc5a3ce" />


## 🖼️ Dashboard in Action
Processing...

<img width="1202" height="537" alt="image" src="https://github.com/user-attachments/assets/06772ff8-b0d0-4866-9f64-b8b3fac5d72d" />
<img width="648" height="550" alt="image" src="https://github.com/user-attachments/assets/e6c66ef4-12f4-4781-9a2e-ae8d5760dcb0" />

Done.

<img width="625" height="590" alt="image" src="https://github.com/user-attachments/assets/19cd14f7-df54-4f55-92aa-e59942172d8f" />

## ⚖️ License & Disclaimer
This project is licensed under the MIT License.

CAUTION: This software is still a work in progress and under development. Always backup your Emails. The author is not responsible for data loss or hardware issues.

## ☕ Support the Project
If "THE ARCHIVIST" saved you some time or made your conversion easier, feel free to help keep the gears turning!

* [**Support via Buy Me a Coffee**](https://www.buymeacoffee.com/jj1eckhardt)
