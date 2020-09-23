Richsoft ZipIt 1.0 beta
-------- ----- --- ----

The Richsoft ZipIt control is a freeware ActiveX control that allows developers to use
standard Zip compression in their applications.

Installation
------------

The following files are required when using the control

* RichsoftZipit.ocx - The actuall ActiveX control
* Zipit.dll } Dll required to access InfoZip dlls (Working on how to by-pass this in later versions)
* Zipdll.dll } Info-Zip dlls
* Unzdll.dll }
* VB6 Runtime files

Control Details
------- -------

The following methods are implemented:

* About - shows information about the control
* Add(Filenames as Collection) - Adds the specified files to the open archive		}
* Delete(Filenames As Collection) - Deletes the specified files from the archive	} Returns the number of files succussfully operated on
* Extract(Filenames As Collection) - Extracts the specified files from the archive	}
* Read - Rereads the archives contents, updating the contents list

The following properties are implemented:

* AddAction as ZipAction - Specifies the action that is performed while adding files
* CompressionLevel As ZipLevel - Specifies the compression level that is used
* ExtractDir As String - Specifies where files are extracted to
* ExtrAction As ZipAction - Specifies the action to be taken while extracting files
* Filename As String - The archive filename.  When this changes the contents is automatically reread
* Overwrite As Boolean - Specifies whether files should be overwritten when extracting from the archive
* UseDirectoryInfo As Boolean - Specifies whether path information is stored within the archive adn whether it is used when extracting files
* UseDOS83Format As Boolean - Specifies whether files should be stored in the 8.3 format for compatibity reasons
* vFiles As Collection - This is a collection of ZipFileEntry objects and contains information about the files within the archive

The following Events are implemented:

* OnArchiveUpdate - Fires whenever the contents of the vFiles property change ie. when the contents of the archive change
* OnDeleteComplete(Successful As Long) - Occurs whenever a file deletion from the archive has completed, returns the number of files succussfully deleted
* OnDeleteProgress(Percentage As Integer, Filename As String) - Gives feedback on the current deletion job, returns the progress as a percentage and also give the filename currently being worked on
* OnUnzipComplete(Successful As Long) - Occurs when an extraction has completed
* OnUnzipProgress(Percentage As Integer, Filename As String) - Gives feed back on the progress of a extraction
* OnZipComplete(Successful As Long) - Occurs when a zip process has completed
* OnZipProgess(Percentage As Integer, Filename As String) - Gives feedback on the current zip process

The following objects are exposed by the control:

* ZipFileEntry - Used to convey information about files contained in the archive
	Version As Integer
	Flag As Integer
	CompressionMethod As Integer
	FileDateTime As String
	CRC32 As Long
	CompressedSize As Long
	UncompressedSize As Long
	FileNameLength As Integer
	ExtraFieldLength As Integer
	Filename As String

The following Enums are exposed by the control:

* ZipLevel - Used to set the compression level
    zipStore = 0
    zipLevel1 = 1
    zipSuperFast = 2
    zipFast = 3
    zipLevel4 = 4
    zipNormal = 5
    zipLevel6 = 6
    zipLevel7 = 7
    zipLevel8 = 8
    zipMax = 9

* ZipAction - Used to set the action that take place when adding or extraction files
    zipDefault = 1 - Adds (and replaces) files 
    zipFreshen = 2 - Freshens (existing) files
    zipUpdate = 3 - Updates (and adds) files


PLEASE NOTE
------ ----

* The RecurseSubFolders property has no effect as I have not implemented this feature yet
* The sample application is not complete - work is continuing and it will be uploaded as soon as it is finished

Legal Stuff
----- -----
Info-Zip Dlls (Zipdll.dll & Unzdll.dll)
Copyright (c) 1990-1999 Info-ZIP.  All rights reserved.
Please read licence.txt for more information

Although I am releasing this software as freeware, if you use any part of it please
give credit.   A phrase such as 'Uses code from Richsoft Computing www.richsoftcomputing.btinternet.co.uk'
is sufficent.

Good luck - More detailed documentation is on the way

Richard Southey
Richsoft Computing (c)2000
