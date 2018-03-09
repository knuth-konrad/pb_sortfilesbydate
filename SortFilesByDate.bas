'------------------------------------------------------------------------------
'Purpose  : Sorts files by their date time (LastWriteTime) and creates a folder
'           chronological folder structure
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 2017
'   Source: -
'  Changed: 15.05.2017
'           - #Break On to prevent console window property's menu issue
'           - Application manifest added
'           - Replace source code include with SLL
'           09.03.2018
'           - Put UNC path derived from mapped drive on its own line to prevent
'           screen clutter
'------------------------------------------------------------------------------
#Compile Exe ".\SortFilesByDate.exe"
#Option Version5
#Dim All

#Link "baCmdLine.sll"

#Break On
#Debug Error Off
#Tools Off

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 3

' Version Resource information
#Include ".\SortFilesByDateRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
   ' Valid CLI parameters are:
   ' /dp or /destinationpath
   ' /sp= or /sourcepath
   ' /f= or /filepattern
   ' /so= or /sortorder
   ' /s= or /subfolders
   ' /ta or /timeattribute
   ' /v= or /verbose
Type ParamsTYPE
   SortOrder As String * 3
   Subfolders As Byte
   Verbose As Byte
   TimeAttribute As Byte   ' 0 = LastWriteTime, 1 - CreationTime, 2 - LastAccessTime
End Type

Type FileSizeTYPE
   Lo As Dword
   Hi As Dword
End Type

Union FileSizeUNION
   Full As Quad
   Part As FileSizeTYPE
End Union
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include Once "win32api.inc"
#Include "ImageHlp.inc"
#Include "sautilcc.inc"       ' General console helpers
'------------------------------------------------------------------------------
'*** Variabels ***
'------------------------------------------------------------------------------
'==============================================================================

Function PBMain () As Long
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sSourcePath, sDestPath, sFilePattern, sCmd, sTemp As String
   Local i, j As Dword
   Local lResult, lTemp As Long
   Local vntResult As Variant
   Local udtCfg As ParamsTYPE

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   ' Application intro
   ConHeadline "SortFilesByDate", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2017-2018", $COMPANY_NAME
   Print ""

   Trace New ".\SortFilesByDate.tra"

   ' *** Parse the parameters
   ' Initialization and basic checks
   sCmd = Command$

   Local o As IBACmdLine
   Local vnt As Variant

   Let o = Class "cBACmdLine"

   If IsFalse(o.Init(sCmd)) Then
      Print "Couldn't parse parameters: " & sCmd
      Print "Type SortFilesByDate /? for help"
      Let o = Nothing
      Exit Function
   End If

   If Len(Trim$(Command$)) < 1 Or InStr(Command$, "/?") > 0 Then
      ShowHelp
      Exit Function
   End If

   ' Parse the passed parameters
   ' Valid CLI parameters are:
   ' /dp or /destinationpath
   ' /sp= or /sourcepath=
   ' /f= or /filepattern=
   ' /so= or /sortorder=
   ' /ta or /timeattribute
   ' /s= or /subfolders=
   ' /v= or /verbose
   i = o.ValuesCount

   If (i < 3) Or (i > 7) Then
      Print "Invalid number of parameters."
      Print ""
      ShowHelp
      Exit Function
   End If

   ' Parse CLI parameters

   ' ** SourcePath
   If IsTrue(o.HasParam("sp", "sourcepath")) Then
      sTemp = Variant$(o.GetValueByName("sp", "sourcepath"))
      sSourcePath = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** DestinationPath
   If IsTrue(o.HasParam("dp", "destinationpath")) Then
      sTemp = Variant$(o.GetValueByName("dp", "destinationpath"))
      sDestPath = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** File pattern
   If IsTrue(o.HasParam("f", "filepattern")) Then
      sTemp = Variant$(o.GetValueByName("f", "filepattern"))
      sFilePattern = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** SortOrder
   If IsTrue(o.HasParam("so", "sortorder")) Then
      sTemp = LCase$(Variant$(o.GetValueByName("so", "sortorder")))
      udtCfg.SortOrder = Trim$(Remove$(sTemp, $Dq))
   End If

   ' ** Recurse subfolders
   If IsTrue(o.HasParam("s", "subfolders")) Then
      vntResult = o.GetValueByName("s", "subfolders")
      udtCfg.Subfolders = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' ** File time attribute
   If IsTrue(o.HasParam("ta", "timeattribute")) Then
      vntResult = o.GetValueByName("ta", "timeattribute")
      udtCfg.TimeAttribute = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' ** Verbose output
   If IsTrue(o.HasParam("v", "verbose")) Then
      vntResult = o.GetValueByName("v", "verbose")
      udtCfg.Verbose = Sgn(Abs(VariantVT(Variant$(vntResult))))
   End If

   ' ** Defaults
   If Len(Trim$(sFilePattern)) < 2 Then
      sFilePattern = "*.*"
   End If

   If Len(Trim$(udtCfg.SortOrder)) < 1 Then
      udtCfg.SortOrder = "ymd"
   End If


   ' *** Echo the CLI parameters

   ' Parse the passed parameters
   ' Valid CLI parameters are:
   ' /dp or /destinationpath
   ' /sp= or /sourcepath=
   ' /f= or /filepattern=
   ' /so= or /sortorder=
   ' /s= or /subfolders=
   ' /ta or /timeattribute
   ' /v= or /verbose

   ' Determine if it's a relative or absolute path, i.e. .\MyFolder or C:\MyFolder and/or a UNC share
   Local sPathFull As String

   sPathFull = sSourcePath
   sPathFull = FullPathAndUNC(sSourcePath)
   Con.StdOut "Source folder     : " & sSourcePath
   ' If path is a relative path, display the full path also
   If LCase$(NormalizePath(sSourcePath)) <> LCase$(NormalizePath(sPathFull)) Then
      Con.StdOut "                    (" & sPathFull & ")"
   End If
   sPathFull = sDestPath
   sPathFull = FullPathAndUNC(sDestPath)
   Con.StdOut "Destination folder: " & sDestPath;
   ' If path is a relative path, display the full path also
   If LCase$(NormalizePath(sDestPath)) <> LCase$(NormalizePath(sPathFull)) Then
      Con.StdOut "                    (" & sPathFull & ")"
   End If

   Con.StdOut "File pattern      : " & sFilePattern
   Con.StdOut "Sort by           : " & Choose$(udtCfg.TimeAttribute, "Last write time", "Creation time", "Last access time")
   Con.StdOut "Sort order        : " & UCase$(udtCfg.SortOrder)
   Con.StdOut "Recurse subfolders: " & IIf$(IsTrue(udtCfg.Subfolders), "True", "False")
   Con.StdOut "Verbose           : " & IIf$(IsTrue(udtCfg.Verbose), "True", "False")

   If IsTrue(udtCfg.Verbose) Then
      Call oPTNow.Now()
      Con.StdOut "Current date/time : " & oPTNow.DateString & ", " & oPTNow.TimeStringFull
   End If

   StdOut ""

   ' *** Sanity checks of CLI parameters
   ' Source folder
   If Len(Trim$(sSourcePath)) < 2 Then
      Print "Missing source folder."
      Print ""
      ShowHelp
      Exit Function
   End If

   If Not IsFolder(sSourcePath) Then
      Print "Source folder doesn't exist: " & sSourcePath
      Print ""
      ShowHelp
      Exit Function
   End If

   ' Destination folder
   If Len(Trim$(sDestPath)) < 2 Then
      Print "Missing destination folder."
      Print ""
      ShowHelp
      Exit Function
   End If

   If Not IsFolder(sDestPath) Then
      Print "Destination folder doesn't exist: " & sDestPath
      Print ""
      ShowHelp
      Exit Function
   End If

   ' Sort order
   If Len(Trim$(udtCfg.SortOrder)) > 1 Then
      sTemp = LCase$(Trim$(udtCfg.SortOrder))
      If Verify(sTemp, "dhmnsy") > 0 Then
         Print "Invalid sort order qualifier: " & UCase$(udtCfg.SortOrder)
         Print ""
         ShowHelp
         Exit Function
      End If
   End If

   Trace On

   lResult = AnalyzeFiles(sSourcePath, sDestPath, sFilePattern, udtCfg)
   StdOut ""
   StdOut "Done. " & Format$(lResult) & " file(s) sorted and moved"

   Trace Off
   Trace Close

   StdOut ""

   PBMain = lResult

End Function
'---------------------------------------------------------------------------

Function AnalyzeFiles(ByVal sSourcePath As String, ByVal sDestPath As String, ByVal sFilePattern As String, ByVal udtCfg As ParamsTYPE) As Long
'------------------------------------------------------------------------------
'Purpose  : Recursivly scan folders for the file patterns passed and move the
'           files over to their "new home", according to their time stamps.
'
'Prereq.  : -
'Parameter: sSourcePath    - Analyze/Move files from this folder (and below).
'           sDestPath      - (Root) folder to move files to.
'           sFilePattern   - File pattern(s) to look for in sSourcePath
'           udtCfg         - Configuration baed on passed CLI parameters (include subfolders, file sort order)
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local sSourceFile, sPattern, sFile, sFileTime, sNewPath As String
   Local sMsg, sTemp As String
   Local i, lCount As Long
   Local udtDirData As DirData
   Local szSourceFile As WStringZ * %Max_Path

   Local oPTNow As IPowerTime
   Let oPTNow = Class "PowerTime"

   Local hSearch As Dword                ' Search handle
   Local udtWFD As WIN32_FIND_DATAW      ' FindFirstFile structure

   Trace On
   Trace Print FuncName$

   For i = 1 To ParseCount(sFilePattern, ";")

      Trace Print " -- AnalyzeFiles sFilePattern: " & Format$(i)

      sMsg = "-- Scanning folder "
      Con.StdOut  sMsg & ShortenPathText(sSourcePath, Con.Screen.Col-(1+Len(sMsg)))

      Trace Print " -- AnalyzeFiles sSourcePath: " & sSourcePath & " (" & Format$(Len(sSourcePath)) & ")"

      sPattern = Parse$(sFilePattern, ";", i)
      Con.StdOut " - File pattern: " & sPattern

      Trace Print " -- AnalyzeFiles sPattern: " & sPattern & " (" & Format$(Len(sPattern)) & ")"

      sSourceFile = NormalizePath(sSourcePath) & sPattern
      Trace Print " -- AnalyzeFiles sSourceFile: " & sSourceFile & " (" & Format$(Len(sSourceFile)) & ")"

      szSourceFile = sSourceFile

      hSearch = FindFirstFileW(szSourceFile, udtWFD)  ' Get search handle, if success
      If hSearch <> %INVALID_HANDLE_VALUE Then        ' Loop through directory for files

         lCount = 0

         Do

            If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY Then ' If not directory bit is set (files only here...)

               sFile = Remove$(udtWFD.cFileName, Any Chr$(0))

               If IsTrue(udtCfg.Verbose) Then
                  sFileTime = GetFileTimeStr(udtWFD, udtCfg)
               End If

               sNewPath = CreateDestinationPath(sDestPath, udtWFD, udtCfg)

               sMsg = "  - Moving "
               Con.StdOut sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)));
               If IsTrue(udtCfg.Verbose) Then
                  Con.StdOut ", Time stamp: " & sFileTime
               Else
                  Con.StdOut ""
               End If
               Con.StdOut "    -> " & sNewPath

               Incr lCount

               ' Create the destination directory, if it doesn't exist and move the file
               Call MakeSureDirectoryPathExists(ByCopy sNewPath)
               Try
                  Call BackupFile(NormalizePath(sSourcePath) & sFile, sNewPath & sFile, 0)

               Catch
                  Con.Color 12, -1
                  sMsg = "  - ERROR: can't move "
                  Con.StdOut sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                  Con.StdOut "Error: " & Format$(Err) & ", " & Error$(Err)
                  ErrClear

               End Try



               Try
'                  If IsFalse(udtCfg.RecycleBin) Then
'                     Kill NormalizePath(sPath) & sFile
'                  Else
'                     Call Delete2RecycleBin(NormalizePath(sPath) & sFile)
'                  End If

               Catch
                  Con.Color 12, -1
                  sMsg = "  - ERROR: can't delete "
                  Con.StdOut  sMsg & ShortenPathText(sFile, Con.Screen.Col-(1+Len(sMsg)))
                  If IsTrue(udtCfg.Verbose) Then
                     Con.StdOut ""
                  End If
                  Con.Color 7, -1
                  Decr lCount

               End Try

            End If   '// If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY

         Loop While FindNextFileW(hSearch, udtWFD)

         FindClose hSearch

      End If   '// If hSearch <> %INVALID_HANDLE_VALUE

      Con.StdOut ""


      If IsTrue(udtCfg.Subfolders) Then  'if to search in subdirectories.

         szSourceFile = NormalizePath(sSourcePath) & "*"
         hSearch = FindFirstFileW(szSourceFile, udtWFD)

         If hSearch <> %INVALID_HANDLE_VALUE Then

            Do

               If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) = %FILE_ATTRIBUTE_DIRECTORY _
                  And (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_HIDDEN) = 0 Then  ' If dirs, but not hidden..

                  If udtWFD.cFileName <> "." And udtWFD.cFileName <> ".." Then          ' Not these..
                     lCount = lCount + AnalyzeFiles(NormalizePath(sSourcePath) & RTrim$(udtWFD.cFileName, $Nul), sDestPath, sFilePattern, udtCfg)
                  End If

               End If

            Loop While FindNextFileW(hSearch, udtWFD)

            FindClose hSearch

         End If

      End If

   Next i

   AnalyzeFiles = lCount

End Function
'---------------------------------------------------------------------------

Function GetFileTimeStr(ByVal udt As DirData, ByVal udtCfg As ParamsTYPE) As String
'------------------------------------------------------------------------------
'Purpose  : Turn a file timestamp into a readable date string
'           time stamps.
'
'Prereq.  : -
'Parameter: udt         - File information about the current file (Win32_Find_Data)
'           udtCfg      - CLI parameters
'Returns  : Readable date & time as string
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local oPTFile As IPowerTime

   Let oPTFile = Class "PowerTime"

   ' 0 = LastWriteTime, 1 - CreationTime, 2 - LastAccessTime
   Select Case udtCfg.TimeAttribute
   Case 0
      oPTFile.FileTime = udt.LastWriteTime
   Case 1
      oPTFile.FileTime = udt.CreationTime
   Case 2
      oPTFile.FileTime = udt.LastAccessTime
   Case Else
      oPTFile.FileTime = udt.LastWriteTime
   End Select

   GetFileTimeStr = oPTFile.DateString & ", " & oPTFile.TimeStringFull

End Function
'---------------------------------------------------------------------------

Function CreateDestinationPath(ByVal sDestPath As String, ByVal udt As DirData, _
   ByVal udtCfg As ParamsTYPE) As String
'------------------------------------------------------------------------------
'Purpose  : Creates the resulting destination string by analyzing the file's
'           time stamps.
'
'Prereq.  : -
'Parameter: sDestPath   - (Root) destination folder
'           udt         - File information about the current file (Win32_Find_Data)
'           udtCfg      - CLI parameters
'Returns  : Constructed full destination path
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 04.04.2017
'           - Allow to sort by different types of file timestamps.
'------------------------------------------------------------------------------
   Local oPTFile As IPowerTime
   Local sResult, sSort As String
   Local i As Long

   Let oPTFile = Class "PowerTime"

   ' 0 = LastWriteTime, 1 - CreationTime, 2 - LastAccessTime
   Select Case udtCfg.TimeAttribute
   Case 0
      oPTFile.FileTime = udt.LastWriteTime
   Case 1
      oPTFile.FileTime = udt.CreationTime
   Case 2
      oPTFile.FileTime = udt.LastAccessTime
   Case Else
      oPTFile.FileTime = udt.LastWriteTime
   End Select

   sSort = LCase$(Trim$(udtCfg.SortOrder))

   sResult &= NormalizePath(sDestPath)

   For i = 1 To Len(sSort)

      ' Figure out the sort qualifier
      Select Case Mid$(sSort, i, 1)
      Case "d"
         ' Day
         sResult &= NormalizePath(Format$(oPTFile.Day, "00"))
      Case "m"
         ' Month
         sResult &= NormalizePath(Format$(oPTFile.Month, "00"))
      Case "y"
         ' Year
         sResult &= NormalizePath(Format$(oPTFile.Year, "0000"))
      Case "h"
         ' Hour as 24 HH
         sResult &= NormalizePath(Format$(oPTFile.Hour, "00"))
      Case "n"
         ' Minute
         sResult &= NormalizePath(Format$(oPTFile.Minute, "00"))
      Case "s"
         ' Second
         sResult &= NormalizePath(Format$(oPTFile.Second, "0000"))
      End Select

   Next i

   Let oPTFile = Nothing
   CreateDestinationPath = sResult

End Function
'---------------------------------------------------------------------------

Sub ShowHelp

   ' Valid CLI parameters are:
   ' /dp or /destinationpath
   ' /sp= or /sourcepath=
   ' /f= or /filepattern=
   ' /so= or /sortorder=
   ' /s= or /subfolders=
   ' /ta or /timeattribute
   ' /v= or /verbose

   Con.StdOut ""
   Con.StdOut "SortFilesByDate"
   Con.StdOut "---------------"
   Con.StdOut "SortFilesByDate searches files matching the passed file pattern in the source folder. It analyses the files' time stamps and sorts them"
   Con.StdOut "accordingly in the destination path by creating the necessary folder structure, with the passed destination path acting as the root folder."
   Con.StdOut ""
   Con.StdOut "Usage:   SortFilesByDate _"
   Con.StdOut "            /sp=<source folder> /dp=<destination (root) folder> /so=ymdhns [/f=<files to sort/move>[;<files to sort/move]] _"
   Con.StdOut "            [/s=0|1]"
   Con.StdOut "i.e.     SortFilesByDate /sp=C:\Data\Incoming /dp=C:\Backup\Invoices /f=*.txt /s=1"
   Con.StdOut ""
   Con.StdOut "Parameters"
   Con.StdOut "----------"
   Con.StdOut "/sp or /sourcepath        = Folder from where to start the sorting process."
   Con.StdOut "/dp or /destinationpath   = (Root) destination folder."
   Con.StdOut "/so or /sortorder         = Sorting pattern."
   Con.StdOut "       Valid parameters are:"
   Con.StdOut "       y - Year"
   Con.StdOut "       m - Month"
   Con.StdOut "       d - Day"
   Con.StdOut "       h - Hour (format 24 HH)"
   Con.StdOut "       n - Minute"
   Con.StdOut "       s - Second"
   Con.StdOut "/ta or /timeattribute     = the file's timestamp attribute by which to sort."
   Con.StdOut "       Valid parameters are:"
   Con.StdOut "       0 - LastWriteTime (default, if the parameter is omitted)"
   Con.StdOut "       1 - CreationTime"
   Con.StdOut "       2 - LastAccessTime"
   Con.StdOut "/f or /filepattern        = file pattern"
   Con.StdOut "       If omitted, all files are scanned (equals /f=*.*)."
   Con.StdOut "/s or /subfolders         = recurse subfolders yes(1) or no (0)"
   Con.StdOut "       If omitted, only the folder passed via /sp is scanned for matching files (equals /s=0)."
   Con.StdOut ""
   Con.StdOut "You may specify more than one file pattern for the parameter /f by using ; (semicolon) as a separator, i.e."
   Con.StdOut "       /f=*.doc;*.rtf -> sorts all *.doc and all *.rtf files from the specified folder."
   Con.StdOut "       /f=Backup*.bak;Log*.trn -> sorts all Backup*.bak and all Log*.trn files from the specified folder."

End Sub
'---------------------------------------------------------------------------

Function FullPathAndUNC(ByVal sPath As String) As String
'------------------------------------------------------------------------------
'Purpose  : Resolves/expands a path from a relative path to an absolute path
'           and UNC path, if the drive is mapped
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 30.01.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   ' Determine if it's a relative or absolute path, i.e. .\MyFolder or C:\MyFolder
   Local szPathFull As AsciiZ * %Max_Path, sPathFull As String, lResult As Long
   sPathFull = sPath
   lResult = GetFullPathName(ByCopy sPath, %Max_Path, szPathFull, ByVal 0)
   If lResult <> 0 Then
      sPathFull = Left$(szPathFull, lResult)
   End If

   ' Now that we've got that sorted, resolve the UNC path, if any
   Local dwError As Dword
   FullPathAndUNC = UNCPathFromDriveLetter(sPathFull, dwError, 0)

End Function
'------------------------------------------------------------------------------

Function UNCPathFromDriveLetter(ByVal sPath As String, ByRef dwError As Dword, _
   Optional ByVal lDriveOnly As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a fully qualified UNC path location from a (mapped network)
'           drive letter/share
'
'Prereq.  : -
'Parameter: sPath       - Path to resolve
'           dwError     - ByRef(!), Returns the error code from the Win32 API, if any
'           lDriveOnly  - If True, return only the drive letter
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 17.07.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   ' 32-bit declarations:
   Local sTemp As String
   Local szDrive As AsciiZ * 3, szRemoteName As AsciiZ * 1024
   Local lSize, lStatus As Long

   ' The size used for the string buffer. Adjust this if you
   ' need a larger buffer.
   Local lBUFFER_SIZE As Long
   lBUFFER_SIZE = 1024

   If Len(sPath) > 2 Then
      sTemp = Mid$(sPath, 3)
      szDrive = Left$(sPath, 2)
   Else
      szDrive = sPath
   End If

   ' Return the UNC path (\\Server\Share).
   lStatus = WNetGetConnectionA(szDrive, szRemoteName, lBUFFER_SIZE)

   ' Verify that the WNetGetConnection() succeeded. WNetGetConnection()
   ' returns 0 (NO_ERROR) if it successfully retrieves the UNC path.
   If lStatus = %NO_ERROR Then

      If IsTrue(lDriveOnly) Then

         ' Display the UNC path.
         UNCPathFromDriveLetter = Trim$(szRemoteName, Any $Nul & $WhiteSpace)

      Else

         UNCPathFromDriveLetter = Trim$(szRemoteName, Any $Nul & $WhiteSpace) & sTemp

      End If

   Else

      ' Return the original filename/path unaltered
      UNCPathFromDriveLetter = sPath

   End If

   dwError = lStatus

End Function
'------------------------------------------------------------------------------
