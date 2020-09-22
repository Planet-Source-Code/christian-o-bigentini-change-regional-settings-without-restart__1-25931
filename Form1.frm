VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1890
      TabIndex        =   9
      Top             =   1245
      Width           =   2355
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1890
      TabIndex        =   8
      Top             =   945
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set  Values"
      Height          =   450
      Left            =   4410
      TabIndex        =   5
      Top             =   1005
      Width           =   1590
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1890
      TabIndex        =   4
      Top             =   450
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1890
      TabIndex        =   3
      Top             =   150
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get  Values"
      Height          =   450
      Left            =   4410
      TabIndex        =   0
      Top             =   210
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Thousand Separator"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   1275
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Decimal Separator"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Thousand Separator"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Decimal Separator"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   195
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LOCALE_SDECIMAL = &HE
Private Const LOCALE_STHOUSAND = &HF
Private Const WM_SETTINGCHANGE = &H1A
      
Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Sub Command1_Click()
Dim SeparadorDecimal As String
Dim SeparadorMiles As String
         
   Dim Symbol As String
   Dim iRet1 As Long
   Dim iRet2 As Long
   Dim lpLCDataVar As String
   Dim Pos As Integer
   Dim Locale As Long
         
   Locale = GetSystemDefaultLCID()
   
   iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)
   Symbol = String$(iRet1, 0)
   iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
   Pos = InStr(Symbol, Chr$(0))
   If Pos > 0 Then
      Symbol = Left$(Symbol, Pos - 1)
      SeparadorDecimal = Symbol
   End If

   iRet1 = GetLocaleInfo(Locale, LOCALE_STHOUSAND, lpLCDataVar, 0)
   Symbol = String$(iRet1, 0)
   iRet2 = GetLocaleInfo(Locale, LOCALE_STHOUSAND, Symbol, iRet1)
   Pos = InStr(Symbol, Chr$(0))
   If Pos > 0 Then
      Symbol = Left$(Symbol, Pos - 1)
      SeparadorMiles = Symbol
   End If
    
    Text1.Text = SeparadorDecimal
    Text2.Text = SeparadorMiles
End Sub


Private Sub Command2_Click()

Dim dwLCID As Long
Dim SeparadorDecimal As String
Dim SeparadorMiles As String
         
         dwLCID = GetSystemDefaultLCID()
         
         SeparadorDecimal = Trim(Text1.Text)
         SeparadorMiles = Trim(Text2.Text)
         
         If SetLocaleInfo(dwLCID, LOCALE_SDECIMAL, SeparadorDecimal) = False Then
            MsgBox "Error Set Decimal Separator"
            Exit Sub
         End If
         
         If SetLocaleInfo(dwLCID, LOCALE_STHOUSAND, SeparadorMiles) = False Then
            MsgBox "Error Set Thousand Separator"
            Exit Sub
         End If
         
         PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0


End Sub


