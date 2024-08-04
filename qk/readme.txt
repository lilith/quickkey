This project contains the actual Quick Key application files and deployment project. All custom controls, logging classes, APIS, and some common utilities are stored in the JCL package.


Files:
Quick Key.sln - Visual Studio 2005 Solution File
Folders:
Quick Key - Source code and project file directory.
Quick Key Setup - Installer package files
Help - Documentation directory.
Final Build - These are the files the installer uses to create the msi package.
media.xcf - Gimp files used to create application graphics.

NOTICE! To run quick key, you will need to copy to contents of Final Build to the runtime directory!

Email me at nathanaeljones@users.sourceforge.net


Remeber to update the following things with each new release:
Final Build\Readme.rtf (If needed)
Final Build\Changelog.txt (Always)
bugstofix.txt (For personal tracking)
newfeaturestodo.txt (For personal tracking)
Final Build\Translators.txt (IMPORTANT! For translators to use!)
Update the copy in the Help\ directory, and recompile the .chm file

Update the web site cache in the Help\ directory and recompile if changes are made.

Ensure that the Upgrade Code is not changed, so that installations will be handled correctly
{917AC10B-38EE-4EF7-AD32-E8C9C3F95848}

Change the Product code and the Setup Package version

The version of JCL in use

Keep the zip file names in sync with the version of QuickKey.exe
