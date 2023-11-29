# <div align="center">Sampling Point Transfer Macro | SampleManager LIMS</div>

## Description
### Overview
This macro loads environmental monitoring sampling points from an Excel .csv file into SampleManager LIMS (a ***L***aboratory ***I***nformation ***M***anagement ***S***ystem). I built this project because I had approximately 4,000 sampling points, each containing a several pieces of data, to load into the information management system and I didn't want to do it manually. The data had to be loaded from a .csv file because the global administrators of the system, who superceded myself as a lcoal admin, did not allow the bulk loading of data from a .csv file. This program completed in just a few hours a job that would've taken my colleague and I several weeks to complete.

### What I Learned
I was learning about object oriented programming (OOP) at the time I wrote this program. I took it as an opportunity to define classes and instantiate objects even though a far simpler list of instructions would have sufficed. You'll notice that some classes only have one method, that no class has attributes, and likewise that no class has getter, setter, or deleter methods. I understand the use of classes for encapsulation of data - I really just wanted to define classes since I hadn't encountered the concept before this point.

### Other Notes and Relation to Past, Undocumented Works
Rather than try to change anything in this script to relfect what I've learned since its completion, I'm preserving it as it was when it last functioned as I no longer have access to the software or to the (proprietary) data that it used. This is purely a historical documentation of work that I've completed.

Since this is the beginning of my portfolio, I will note that this script is reminiscent of a handful of Runescape bots that I wrote in high school and following my college graduation, so it also serves as a representative for sevaral of my lost works. The key differences in the Runescape bots I wrote were: 1) I did not know about the commands for activating windows by name, so I had to place windows in precisely the correct position which made my past scripts less flexible, and 2) to avoid detection by the bot monitoring systems, I included a randomization function for the pixel coordinates on every click the Runescape bots made.

### Why I Published this Work
I wrote this script on company time. However, it is in no way related to any product that my former employer sells. It is essentially a single-use tool written by me to solve a problem they will likely only encounter 2-3 more times worldwide; a problem which they anticipate paying people to solve manually. Furthermore, I would not expect the average employee in my former role to be able to pick this script up with no oversight and use it to complete the task themselves, so I don't realistically believe that it is of value to the company any longer. Nothing in this script reveals anything related to company intellectual property (aside from the script itself). The only remaining value of the script is to document that I applied programming knowledge to solve a real problem, that I saved my employer thousands of dollars, and that I saved myself weeks of time by executing this project.


## Features
The sole feature of this script is to open SampleManager LIMS, navigate to the Sample Point entry page, copy the pertinent data from a pre-created Excel file, and use those data to complete the Sample Point creation form.

The sleep times were manually optimized where applicable to complete the data entry as fast as possible while accounting for occasional system latency. In many cases, I used a pre-built method to wait for windows to become available in order to prevent the program from running too fast.

An improvement upon past works of mine which were similar is the use of the built-in window activation method. With this improvement, windows need not be placed in precisely the correct location on the screen as all coordinates descripbe points relative to the origin of the active window as opposed to being an absolute position on the screen. The windows need only be large enough to display all required items without scrolling.


## How to Use
### Software Requirements:
* SampleManager LIMS version 12
* Microsoft Excel
* AutoHotKeys

### Instructions
1. Log into the "Prod" environment of SampleManager LIMS.
2. Open the Excel .csv file titled "Sampling Point List Edited" which contains the exported table of sampling points from LabWare LIMS (the legacy information management system from which the data are being migrated).
3. Ensure that columns 1-3 contain the information for sampling point ID, Name, and Desc., respectively. These columns should be fully visible in Excel and on the left-most side of the screen. Column widths should be: 18.71, 55, and 96.29, respectively.
4. This script enters samples from the bottom of the column going up. Scroll down until none of the data points to be entered are visible. Click on a cell in column A (the ID column) and use the "up" arrow key to navigate up until the first data point to be entered is visible at the top of the window. Of all the data points to be entered, this should be the only one visible.
5. When the above tasks are complete, click "OK" to proceed or click "Cancel" to exit.


## Technologies
* **AutoHotKeys (AHK):** AutoHotKeys was chosen because I did not have access to any sort of API for the information management system software. As such, I wasn't able to communicate directly with the software and had to resort to a macro. I have used AHK in the past to great effect and I was already fairly familiar with the capabilities. I did not need anything more complicated than AHK to solve this problem.

## Collaborators:
I developed this simple script by myself.