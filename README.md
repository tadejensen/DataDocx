**DataDocx** is an open source tool for generating structured, professional wind turbine inspection reports as `.docx` files.
It is designed for **ease of use**, **high customizability**, and **full data ownership** - your data remains private and fully under your control.

## Key Features
* Fully customizable checklist and templates
* Report generated in `.docx` (Microsoft Word) format
* Efficient remark and data entry and reuse via prefills and shortcuts
* Visual helpers: maintenance timeline, temperature comparison, chapter status bar, project overview block
* Offline-capable, Git-compatible data syncing
* Data stored in CSV files - ideal for backups, analysis, and versioning

## How it works
DataDocx is centered around **remarks** drawn from a customizable checklist. These remarks, combined with inspection metadata, turbine components, and optional temperature data, are compiled into a complete `.docx` report.
The checklist holds information about the chapter structure and a remark's title, summarizing its content and grouping similar remarks.

## Installation:
1. Configure your setup in config.py
2. put your signature into /databases/report as png file
3. (optional) create git in databases folder, link to GitHub Repository

## Workflow:
1. Create a project folder with subfolders for each turbine
    *(Folder name must exactly match turbine ID)*
2. Launch DataDocx and:
    * Configure the project
    * Set up at least one turbine
3. Add remarks and images for each chapter
    *(images go into the auto-created /0-Fertig folder)*
4. Mark chapters as done to preview content
5. Enter turbine components and temperature data (if applicable)
6. Click **"Prüfbericht erstellen"** to generate the final report

## Important hints/bugs:
* Turbine folder names must match turbine ID exactly and cannot be changed after project creation.
* Data like turbine ID, year of inspection, inspection type and OEM cannot be changed after project creation.
* DataDocx is in **Alpha**, so please report bugs.

## Q & A
### What do I see on a typical DataDocx page?
You see the turbines at the top with their respective progress of the report. Underneath, you see the remark selection frame. It holds the titles and remarks for a certain chapter of the repotr. Here you see all the remarks that have already been added to the report and all the titles that are in the checklist for the current chapter. From here, you can add, delete and modify remarks. The checklist is not shown directly, because one title can hold several pre-existing remarks present in the checklist. The title basically summarizes a bunch of similar remarks in the checklist. This reduces cluttering compared to showing all remarks in the checklist as full text.
### What is a remark?
A remark is the text (and flag and images) that is displayed in the report docx. The remark also has a title, which is like a small handle summarizing the remark's content. One title can fit for several remarks.
### What is a "Flag"?
Basically, it's the symbol the precedes a remark, so the well known P, E, I, V flags. These have colors red, orange, blue and green.
There are special Flags:
- *: (hidden) indented remark, but without a number or a visible flag. Use e.g. in "Fazit".
- -: (bullet) indented remark, but with a bullet. Use e.g. in "Zusätzliche Ausrüstung der Anlage". With an "-" flag, a new bullet will be given to every new paragraph of a remark.
- S: (sentence) unindented, no symbol. A regular sentence. Use e.g. in "Pflichten des Auftraggebers bzw. des Betreibers".
- 0, 2, 3, 4: (healine colors) make a square before an underlined line of text. Use in "Fazit". Colors: 0: grey, 2: green, 3: yellow, 4: red.
- RAW: The remark's text will be interpreted as python code. Use only the pre-existing RAW remarks, except if you know what you're doing. 
### How do I manipulate the checklist?
The checklist is the source of the pre-existing remarks. You manipulate it by first going into the remark editor for the remark whose checklist entry you want to manipulate. There, you click "Zu Vorlage hinzufügen" or "Vorlage anpassen". This opens the checklist editor with / without the text that is currently in the remark editor's textbox. You can also set the order of the remarks in the checklist, as well as recommended or default flags. Remarks with default flags will be added to the report upon setup of the turbine. You can also blacklist or whitelist the remark for certain oems/turbine types/tower types.
### How do I add a new title?
Click the bottommost "+" button. Enter the new title in the top right Textbox. Enter the remarktext as usual.
### How do I skip a chapter?
Leave the chapter empty and mark it as done. If you don't want the Chapter to appear for the turbine type/tower type, put the turbine/tower type into the blacklist for every remark of the chapter.
### How do I insert a chapter?
Press the botttommost "+" Button (the one that has no title next to it). In the top left Textbox, enter the new chapter's path, which consists of (optional) the new chapter's superchapter and the new chapters name like so: Prüfbemerkungen|{superchapter}|{subchapter}
Note: only use in "Prüfbemerkungen". Separate superchapter and subchapter with the "|" character. Leave no blanks before or after the "|". Your new chapter should appear after pressing the button with the turbine's ID. To permanently add the newly added chapter, put any remark for this chapter into the checklist.
### I encountered a Bug. What do I do?
Try to decrypt the error message, maybe some input you made caused the bug. In that case, adjust the input. 
Else: Contact me. See my contact data in the 8.2 Forum (I do not want to make them public, to avoid spam) or contact me on GitHub (tade-jensen).
### I want to change the order of my remarks. How do I do it?
In the Remark Manipulation Window, put in the desired position in the "Pos." Textbox. After checking the "Abschnitt fertig" checkbutton, you will see the order of the remarks as it will appear in the report docx.
### What does the colorful grid on the right mean?
It's a helper to let you see if the title is present in other turbines. Each square represents one turbine, the color represents the "worst" flag of this title in each turbine. If you click the squares you will go to the corresponding turbine's page for the current section.
### What does the colorful bar at the bottom mean?
It's the "Chapterbar". It's build from a series of rectangles, representing each section of the report. Click on any rectangle to get to the corresponding chapter. It highlights the chapter you are currently in and shows you which chapters are marked as done (they turn green, rest is white). They also show you the worst flag in each chapter (small rectangle below). The vertical lines represent changes in the current subchapter (eg. change from "Fundament und Turm" to "Maschinenhaus"). Use it and you will become good at eyeballing the chapter you are looking for.
### How do I insert a linebreak in the timeline?
Use the "|" symbol (Alt Gr + <).


