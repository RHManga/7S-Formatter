# Seven Seas Script Formatter

Bumbling through rudimentary coding until 7S fixes their script formatting.

Version 0.1.1
2024/04/22

Current release here: https://github.com/RHManga/7S-Formatter/releases/tag/main

# What this application is: 
    
This is an automated script formatter for Seven Seas .DOCX scripts, designed to take the headache out of formatting manga scripts for lettering.

# Setup (Python)
	
Currently this is only usable if run through Python from the command line, because Windows Defender keeps marking the EXE version as malware, so you will need Python installed from the link below:
https://www.python.org/downloads/
	
Once you've installed python, go to your Command line (Windows+R, type cmd.exe in the field, press enter; or you can just right click a space in your file explorer window and select "Open in Terminal") and type in the following commands:
`pip install python-docx`
`pip install doc2txt`

These commands will install some extensions that will be necessary to run the formatter scripts. With that, you're good to go. Read the next section for a detailed breakdown on each of the Formatter's functions, or skip the the How to Use section below that to get started.

# What each button does:

**Open File** — Selects a .DOCX format Microsoft Word document. This is the script of whatever book you're working with. This program is **not** compatible with .DOC format.

**Format Script for Lettering** — This performs several functions:
        
 - Extracts the Word document into a .TXT script
 - Adds text tags ([b][/b], [i][/i], [k][/k]) around bold, italic, and bold-italic words in the text version of the script.
 - Checks to make sure that tags are only attached to words they're modifying, so that there are no tags touching [/each][other] or touching words before/after the opening/closing tags[b] like [/b]this(this would turn into [b]like[/b] this)
 - Runs the standard RegEx deletions on the script, removing speaker and panel labels.
 - Changes "<heart>" to the heart glyph
 - Runs a function that removes all sound effects from the script so you can just focus on dialogue.
 - Removes all empty lines to condense the line count.
 - Separates "Page XXX" lines with a new line so the script is nice and readable.

**Extract Sound Effects** — This generates a .TXT file containing ONLY page labels followed by the sound effects that appear on that page. Useful for books with lots of sound effects, where they might clutter up the script (or vice versa)

# How to use:

 1. Open the folder containing the SevenSeasScriptFormatter.py file and right click an empty space in the folder. 
 2. Open the application by typing 'python SevenSeasScriptFormatter.py' and pressing Enter. Read the instructions at the top of the window.
 3. Click the Open File button and navigate to your DOCX Script.
 4. Open the script. You will see your file name appear above the buttons now.
 5. Click "Format Script for Lettering." This generates a .TXT file of your now clean script in the .DOCX's folder.
 6. Click "Extract Sound Effects" to generate a .TXT file with only sound effects.
 7. Click "Close" to close the application.

# Additional Instructions for InDesign:

- Make sure you install the included InDesign script (Style Tag Converter_Dialog.jsx) in your InDesign script folder.
- When typesetting, ensure your text does not overflow outside of the text frames.
- You may need to widen your text frames a little more than normal.
- The reason for this is that the find/replace functions of the InDesign script will not work on any text that is not inside a text frame. Overflowed text counts as outside a text frame.
- Run this InDesign script after typesetting.
- After running the InDesign script, you'll need to Ctrl+F to delete the tags. 
- Just type `[b][/b][i][/i][k]`, and `[/k]` into the Find field, leave the Replace fields blank (both content and style fields), and make sure it's applied to "Document" and click the "Replace All" button.
- You'll have to do this one at a time until I find a fix for this surprisingly hard-to-implement function.

# Issues and Compatibility problems:

If an error occurs, it will generate an error.log file in the application directory.
	
This is currently only compatible with .docx format files. It doesn't work on .doc files. Do not try to use a .doc file. The file selector is limited to .DOCX format, so this shouldn't be an issue, but don't even try it. I'm currently trying to add compatibility.

This application has no configuration options and will always format scripts in the way described above. I'm working on a way to make the functions modular so you can turn individual formatting functions on or off, but as this was designed for a very Seven-Seas-specific purpose, I haven't implemented them yet.

The InDesign script uses Character Styles to format text, so make sure you have Character Styles defined for Bold, Italic, and BoldItalic before running it.
	
For some reason, this does not work with .docx files saved in LibreOffice. I don't know about its compatibility with Google Docs .docx files (though as of now, it doesn't seem like there are any), and there's honestly not much I can do about that, so make sure that you're not saving documents in LibreOffice (or OpenOffice, to be safe) and then running them through the formatter.
	
# Features to implement:
	
- Fix InDesign script to automatically remove tags without fucking up the formatting.
- Make the script formatter configurable so you can select and turn off the features
- Make the "replace hearts" function of the script formatter work with music notes (impossible without a lot of 7S scripts for reference since every translator has their own way of denoting them)
	- To this end, I also want to make it configurable so that you can choose to use the glyph or format it into text or a designated symbol or however you want to deal with it in your own typesetting.
- Find a more efficient way of doing the RegEx stuff (it currently looks like YandereDev code)
- Make an Extract Asides function that keeps the panel labels
