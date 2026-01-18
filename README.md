# PowerPoint2PDF exporter
Hi there! This is a very simple vbs script that takes ppt, pptx, ppsx inputs and outputs a vertical 2 slide page of the PowerPoint.
I've been using a Remarkable 2 tablet for revision and found this is a more usable/manageable format for lectures with lots of slides.

## Files
Either drag and drop the file/folder of files you're wanting to convert onto the vbs script, or in the command line (using cscript) ``cscript Converter.vbs <file>``
cscript may not be natively installed.

## Output
Bit of a misnomer when considering this as a "converter", it doesn't convert anything, it just allows for a single or a batch of files to be exported. They are exported with the same name into the same directory, it is worth noting that this means there will be *duplicates* of whatever filename you have. 

## How does it work?
- Is this a file or a folder?
- If it's a file, is it a file I can export with?
- If it is, open up PowerPoint and open the file, export it with our settings, and close PowerPoint.
- If not, dont do that.
- If it's a folder, look through it's contents and determine if it has folders or files.

This should work with almost any amount of PowerPoints, and you'll get a PDF of each one right next to it.
