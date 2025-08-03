Attribute VB_Name = "modShell"
'******************************************************
'*  modShell - various global declarations
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Public Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
 End Type
 Public Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
 End Type
 Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long
 Public Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long
 'Public Declare Function CloseHandle Lib "kernel32" (ByVal _
 '     hObject As Long) As Long
 Public Const NORMAL_PRIORITY_CLASS = &H20&
 Public Const INFINITE = -1&

 Public Sub gShellAndWait(ByVal RunProg As String)
      Dim RetVal As Long
      Dim proc As PROCESS_INFORMATION
      Dim StartInf As STARTUPINFO
      StartInf.cb = Len(StartInf)
      'Execute the given path
      RetVal = CreateProcessA(0&, RunProg, 0&, 0&, 1&, _
           NORMAL_PRIORITY_CLASS, 0&, 0&, StartInf, proc)

      'Disable this app until the executed one is done
      RetVal = WaitForSingleObject(proc.hProcess, INFINITE)
      RetVal = CloseHandle(proc.hProcess)
 End Sub

