# Word-Count
VBA script to calculate the number of words in a Microsoft Word document, excluding captions, headings and tables.

## Installation
In Microsoft Word, open the VBA editor using Alt + F11. Click on module, then insert module. This will open a VBA editor.
Paste the script in script.vba into this window and save.

To add a button to run the script in the quick-access toolbar, go to file, options, then quick access toolbar. Select the
ExcludeTableCaptionAndHeadingWordsFromWordCount macro and press add. Through the modify button you can then choose a 
custom name and icon for the script.

This can be setup to work off a keyboard shortcut. Go to file, options, customize ribbon, keyboard shortcuts, macros,
ExcludeTableCpationAndHeadingWordsFromWordCount, and add the new shortcut there. Ctrl + Shift + C is a good example,
as it is inline with Google Docs' advanced word count, and is currently unused in Microsoft Word.
