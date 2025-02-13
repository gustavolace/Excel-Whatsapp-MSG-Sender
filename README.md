# WhatsApp Automation from Excel

## Summary
1. [Features](#features)
2. [Usage](#usage)

## Features
The WhatsApp automation code from Excel aims to bypass the requirement imposed by WhatsApp for using its API (which is paid and complex). I used the Selenium library, along with other techniques, to access the Excel spreadsheet.

- **List numbers with names**:
  
  ![List of Numbers with Names](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/0cbb1d48-8fe8-495c-9967-53696ffade7a)

- **Send perfectly formatted messages, including line breaks**:
  
  I used the reserved keyword `$cl` to replace the customer's name in the text.

  ![Formatted Message](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/e751e543-82c6-4923-99d1-20d0472b570b)

- **Send images along with text**:
  
  Just place the image over the cells.

  ![Image Sending](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/d647be5c-1c46-4a04-b6c1-e17296b3a55e)

- **Complete Spreadsheet**:
  
  ![Complete Spreadsheet](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/77730f35-fcbf-46f5-bd0b-1510af6a1210)
  ![image](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/a62553a2-017d-441a-a715-1fd38dc71471)

## Usage
First of all, pay attention to the compressed files; it would be impossible to send the entire file to GitHub, so extract both main files.

Select all and click to extract here <br>
Move the extracted file to the root folder

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/f292a377-95f5-40fd-adb3-4222a9ecca27)
![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/eacd20ec-0127-43c0-afb4-8bf78bb6f245)

The code includes a series of actions to acquire information from Excel and then send it through the browser. To start your setup, open the `settings.exe` file. This will open the browser; scan the code within 30 seconds and then close it normally. Everything is ready to use.

Note: The executables must be in the same folder as the spreadsheet for it to work. Do not make changes to the overall structure of the spreadsheet, except for adding or removing images, text, or contacts.

The code contains two executables: `headless` runs the browser offscreen, while `debug` shows the messages. By default, `headless` is enabled.

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/d3e52f50-cd80-4957-8e7e-2bbfa389c93f)

To make the change, open Excel in administrator mode, right-click on the "send" button, assign macro, `ExecutarExe`, edit.

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/0f99dbd7-de31-4768-83b7-cb16fac9884a)

Make the necessary change by replacing `_headless` with `_debug`.

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/1fb89815-814e-4f20-ab11-292e5bafb8e2)

With that in mind, just fill in the contacts and names in the spreadsheet. Make sure to put a name, even if it is not relevant when sending the message. For possible errors, restart the computer or contact me. If you are afraid that messages will be sent improperly or to the wrong person, the worst that can happen is that the message is not sent due to memory or internet issues. Note that the maximum time the code waits before it stops sending the message is 40 seconds.

A new simple setup and use version has been made.
