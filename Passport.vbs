''''''''''''''''''''''''''''''
'''' SCRIPT CONFIGURATION ''''
''''''''''''''''''''''''''''''

' TrueCypt Drive
Const TCD = "W"

' TrueCrypt Password
Const TCPW = "RHAMbackup2008"

' Note:  The /C parameter makes the cmd.exe window close after completing the command sequence.
'        The /K parameter, alternatively, keeps the command window open.
Const CMDEXE = "CMD /C "
Const CHGDIR = "CD\ & C:"
Const LINK = " & "
Const NormalFocus = 1

''''''''''''''''''''''''''''''
''' TRUECYPT COMMAND LINE ''''
''''''''''''''''''''''''''''''

' Command line to instruct TrueCrypt to open any devices,
' with the given password, and in quiet (background) mode.
strLaunchTC = Chr(34) & "C:\Program Files\TrueCrypt\TRUECRYPT.EXE" & Chr(34) & " /L " & TCD & " /A DEVICES /P " & TCPW & " /Q"

strRun = CMDEXE & CHGDIR & LINK & strLaunchTC

' Mount TrueCrypt volume.
Set objShell = CreateObject("WScript.Shell")

objShell.Run strRun, NormalFocus
wScript.Sleep(5000)


''''''''''''''''''''''''''''''
'''''' IDENTIFY SOURCE '''''''
''''''''''''''''''''''''''''''

' Make the remote backup from the most recent local backup
' GetFolder does not accept trailing slashes on folder paths
strBKUPDIR = "X:\Backups_Local"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strBKUPDIR)
Set colSubfolders = objFolder.Subfolders

' Add trailing slash for use in building path names.
strBKUPDIR = strBKUPDIR & "\"

' Find THE folder whose name is of the form -0_BKUP_*
' This is the folder to mirror.
blnAbort = False
strSourceDir = ""
For Each objSubfolder in colSubfolders
  strSubFolderName = objSubfolder.Name
  If Left(strSubFolderName, 8) = "-0_BKUP_" Then
    If strSourceDir = "" Then
      strSourceDir = strSubFolderName
    Else
      blnAbort = True
    End If
  End If
Next

' Do not proceed if multiple suitable source folders exist.
If blnAbort Then MsgBox "Multiple source folders (-0_BKUP_*) exist."

' Do not proceed if no suitable source folder found.
If strSourceDir = "" Then
  blnAbort = True
  MsgBox "No suitable source folder (-0_BKUP_*) exists."
End If


''''''''''''''''''''''''''''''
''''''' IDENTIFY DEST. '''''''
''''''''''''''''''''''''''''''

' If source folder found, then look for destination folder.
If Not blnAbort Then

  Set objFolder = objFSO.GetFolder(TCD & ":")
  Set colSubfolders = objFolder.Subfolders

' Find THE folder whose name is of the form Backup*
' This is the destination folder.
  blnAbort = False
  strDestDir = ""
  For Each objSubfolder in colSubfolders
    strSubFolderName = objSubfolder.Name
    If Left(strSubFolderName, 6) = "Backup" Then
      If strDestDir = "" Then
        strDestDir = strSubFolderName
      Else
        blnAbort = True
      End If
    End If
  Next

' Do not proceed if multiple suitable destination folders exist.
  If blnAbort Then MsgBox "Multiple destination folders (Backup*) exist."

' Do not proceed if no suitable destination folder found.
  If strDestDir = "" Then
    blnAbort = True
    MsgBox "No suitable destination folder (Backup*) exists."
  End If

End If


''''''''''''''''''''''''''''''
'''''''' MAKE MIRROR '''''''''
''''''''''''''''''''''''''''''

' Proceed only if single source and single destination folders exist.
If Not blnAbort Then

' Update name of destination folder.
  strDestDirOld = TCD & ":\" & strDestDir
  strDestDirNew = TCD & ":\Backup_" & Right(strSourceDir, Len(strSourceDir) - 8)
  objFSO.MoveFolder strDestDirOld , strDestDirNew
  wScript.Sleep(5000)

  strSourceDir = strBKUPDIR & strSourceDir

' Robocopy with command line parameters to mirror the source folder.
  strRBC = "ROBOCOPY " & strSourceDir & " " & strDestDirNew & " /MIR /XA:SH /R:2 /W:5 /NP /LOG:" & strBKUPDIR & "Backup_Passport.log"

' Remove the Hidden and System attributes ascribed by Robocopy to the most recent archive folder.
  strDelAttrib = "ATTRIB -H -S " & strDestDirNew

  strRun = CMDEXE & CHGDIR & LINK & strRBC & LINK & strDelAttrib

' Launch RoboCopy backup.
' objShell already set, so no need to reset:  Set objShell = CreateObject("WScript.Shell")
  objShell.Run strRun, NormalFocus

End If
