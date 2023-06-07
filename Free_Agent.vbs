' CAUTION:  THIS SCRIPT IS INCOMPLETE.  DO NOT RUN IT.


''''''''''''''''''''''''''''''
'''' SCRIPT CONFIGURATION ''''
''''''''''''''''''''''''''''''

Const TCD = "truecrypt /l w /v d:\qb /p RHAMbackup2008 /q"
Const RBC = "robocopy H:\ W:\Backup-2Q10\ /MIR /R:0 /V"


''''''''''''''''''''''''''''''
''''''' BACKUP  SCRIPT '''''''
''''''''''''''''''''''''''''''

' Note:  The /C parameter makes the cmd.exe window close after completing the command sequence.
'        The /K parameter, alternatively, keeps the command window open.
Const CMDEXE = "CMD /C "

Const CHGDIR = "CD\ & C:"
Const LINK = " & "
Const NormalFocus = 1

Set objShell = CreateObject("WScript.Shell")

' Mount TrueCrypt Device and call Robocopy backup.
strRun = CMDEXE & CHGDIR & LINK & TCD & LINK & RBC
objShell.Run strRun, NormalFocus


' W:\Backup_2010_Q2\H
' W:\Backup_2010_Q2\K
' X:\Backups\-0_BKUP_2011-12-23_Fri\H
' X:\Backups\-0_BKUP_2011-12-23_Fri\K
' No need to backup H & K separately, just back up parent folder
' Backup_Free_Agent.log, similar to Backup_Passbook.log
' Find a way to identify the Free Agent drive and create the folder ' Backup_YYYY_Q# to which the backup will be made.
' Prompt user for drive location?
' Drive should have file QB saved to its root.
' Ask user to identify drive, then test if QB file exists and with size around 487,587,840 KB.  Just check if it's greater than 400 GB.

