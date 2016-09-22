# Basic Export to Word for Jama

This is a replacement for the default export to Word included with Jama that offers the following features:
* Only Components, Sets and Folders appear in the Document Map / Table of Contents
* Items are exported in a clean, easy to read format
* Images are automatically resized to never extend beyond the width and height of the page
* Tables are resized to never extend beyond the width of the page
* Text that is not a heading is forced to have an outline level of Body Text to ensure it does not appear in the Document Map / Table of Contents
* Description contents are forced to always use the font Arial
* Extra white space is removed

## Installation

To install this export in Jama:

1. Click on a Component, Set or Folder in your Jama project
2. Select Export -> Office Templates
3. Select Upload a Templates
4. Select the default_word_template.doc in this repository
5. Give the report a name
6. Click Save Report

Alternatively, to replace the include default export to Word:

3. Select Export to Word Default
4. Click Edit
5. Select the default_word_template.doc in this repository
6. Click Save Report

## Using the export

To try out the export:

1. Click on a Component, Set or Folder in your Jama project
2. Select Export -> Office Templates
3. Select the template you uploaded above
4. Click Run
5. Open the downloaded Word document
6. Enable macros
7. Wait for the macro to finish running
8. Save the file as a docx formatting Word document to strip out the macro as it is no longer needed

## Macro source code

Much of the formatting is accomplished using a macro embedded in the Word document. Main.bas and Functions.bas include the source code for the macro.