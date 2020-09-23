VERSION 4.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warcraft II Trainer"
   ClientHeight    =   1815
   ClientLeft      =   3015
   ClientTop       =   3450
   ClientWidth     =   2655
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   2220
   Icon            =   "frmMain.frx":0000
   Left            =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   1815
   ScaleWidth      =   2655
   Top             =   3105
   Width           =   2775
   Begin VB.Frame Frame1 
      Caption         =   "Made by Warp."
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command9 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Minimize"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "A&bout"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ship Ca&nnons"
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Catapults"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ship A&rmor"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S&heilds"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Arrows"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Swords"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
hwnd = FindWindow(vbNullString, "Warcraft II")
If (hwnd = 0) Then
MsgBox "Warcraft II BNE is not loaded.", vbCritical, "Warcraft II Trainer"
Exit Sub
End If
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (pHandle = 0) Then
Exit Sub
End If
WriteProcessMemory pHandle, &H443330, "êêêêêê", 6, 0&
MsgBox "Swords Constant Upgrade Enabled.", 64, "Warcraft II Trainer"
CloseHandle hProcess
End Sub

Private Sub Command2_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
hwnd = FindWindow(vbNullString, "Warcraft II")
If (hwnd = 0) Then
MsgBox "Warcraft II BNE is not loaded.", vbCritical, "Warcraft II Trainer"
Exit Sub
End If
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (pHandle = 0) Then
Exit Sub
End If
WriteProcessMemory pHandle, &H443580, "êêêêêê", 6, 0&
MsgBox "Sheilds Constant Upgrade Enabled.", 64, "Warcraft II Trainer"
CloseHandle hProcess
End Sub


Private Sub Command3_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
hwnd = FindWindow(vbNullString, "Warcraft II")
If (hwnd = 0) Then
MsgBox "Warcraft II BNE is not loaded.", vbCritical, "Warcraft II Trainer"
Exit Sub
End If
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (pHandle = 0) Then
Exit Sub
End If
WriteProcessMemory pHandle, &H443380, "êêêêêê", 6, 0&
MsgBox "Catapults Constant Upgrade Enabled.", 64, "Warcraft II Trainer"
CloseHandle hProcess
End Sub

Private Sub Command4_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
hwnd = FindWindow(vbNullString, "Warcraft II")
If (hwnd = 0) Then
MsgBox "Warcraft II BNE is not loaded.", vbCritical, "Warcraft II Trainer"
Exit Sub
End If
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (pHandle = 0) Then
Exit Sub
End If
WriteProcessMemory pHandle, &H4432C0, "êêêêêê", 6, 0&
MsgBox "Arrows Constant Upgrade Enabled.", 64, "Warcraft II Trainer"
CloseHandle hProcess
End Sub


Private Sub Command5_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
hwnd = FindWindow(vbNullString, "Warcraft II")
If (hwnd = 0) Then
MsgBox "Warcraft II BNE is not loaded.", vbCritical, "Warcraft II Trainer"
Exit Sub
End If
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (pHandle = 0) Then
Exit Sub
End If
WriteProcessMemory pHandle, &H443660, "êêêêêê", 6, 0&
MsgBox "Ship Armor Constant Upgrade Enabled.", 64, "Warcraft II Trainer"
CloseHandle hProcess
End Sub

Private Sub Command6_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
hwnd = FindWindow(vbNullString, "Warcraft II")
If (hwnd = 0) Then
MsgBox "Warcraft II BNE is not loaded.", vbCritical, "Warcraft II Trainer"
Exit Sub
End If
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (pHandle = 0) Then
Exit Sub
End If
WriteProcessMemory pHandle, &H4435F0, "êêêêêê", 6, 0&
MsgBox "Ship Cannons Constant Upgrade Enabled.", 64, "Warcraft II Trainer"
CloseHandle hProcess
End Sub

Private Sub Command7_Click()
MsgBox "Warcraft II Trainer." & Chr(13) & Chr(10) + "For - Warcraft II BNE" & Chr(13) & Chr(10) + "Author - Warp" & Chr(13) & Chr(10) + "E-Mail - Warp_X@Lycos.com" & Chr(13) & Chr(10) + "ICQ Number - 103902367" & Chr(13) & Chr(10) + "" & Chr(13) & Chr(10) + "If you use these codes give me credit! DO NOT be lame!" & Chr(13) & Chr(10) + "ONLY use this trainer for defense!", 64, "About Warcraft II Trainer"
End Sub

Private Sub Command8_Click()
Me.WindowState = 1
End Sub

Private Sub Command9_Click()
End
End Sub


Private Sub Timer1_Timer()
' Declare some variables we need
   Dim hwnd As Long        ' Holds the handle returned by FindWindow
   Dim pid As Long         ' Used to hold the Process Id
   Dim pHandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   hwnd = FindWindow(vbNullString, "Warcraft II")
   If (hwnd = 0) Then
      MsgBox "Warcraft II is not loaded!", vbCritical, "Game not Found"
      Exit Sub
   End If
   
   ' We can now get the pid
   GetWindowThreadProcessId hwnd, pid
   
   ' Use the pid to get a Process Handle
   pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
   If (pHandle = 0) Then
      MsgBox "Couldn't get a process handle!"
      Exit Sub
   End If
    
   ' Now we can write to our address in memory
   WriteProcessMemory pHandle, &H443330, "êêêêê", 5, 0&
    WriteProcessMemory pHandle, &H443335, "ê", 5, 0&

   ' Close the Process Handle
   CloseHandle hProcess
End Sub


