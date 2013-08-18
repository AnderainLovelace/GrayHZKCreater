VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gray HZK Creater - by Anderain"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer tmrProc 
      Left            =   -120
      Top             =   2040
   End
   Begin VB.PictureBox picBack2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   13
      Top             =   2280
      Width           =   4095
      Begin VB.PictureBox picOutput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " 4-bit Gray"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.Frame frame1 
      Caption         =   "Font"
      Height          =   4215
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtSize 
         Height          =   270
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Start!"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   2295
      End
      Begin VB.CommandButton cmdGetAndTest 
         Caption         =   "Get && Test"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtYOffset 
         Height          =   270
         Left            =   1080
         TabIndex        =   12
         Text            =   "0"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtXOffset 
         Height          =   270
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtChs 
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdChangeFont 
         Caption         =   "Change Font"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtHeight 
         Height          =   270
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtWidth 
         Height          =   270
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "byte(s)"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Y Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "X Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.PictureBox picOriginal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1680
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Original"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Dim Mat(0 To 128, 0 To 128) As Byte
Dim gL As Long, gP As Long


Private Sub PicCenter()

    picOriginal.Left = (picBack.Width - picOriginal.Width) / 2
    picOriginal.Top = (picBack.Height - picOriginal.Height) / 2
    
    
    picOutput.Width = picOriginal.Width
    picOutput.Height = picOriginal.Height
    
    picOutput.Left = (picBack2.Width - picOriginal.Width) / 2
    picOutput.Top = (picBack2.Height - picOriginal.Height) / 2
    
End Sub


Private Sub cmdChangeFont_Click()
    On Error Resume Next
    
    cmnDlg.FontBold = picOriginal.Font.Bold
    cmnDlg.FontItalic = picOriginal.Font.Italic
    cmnDlg.FontName = picOriginal.Font.Name
    cmnDlg.FontSize = picOriginal.Font.size
    cmnDlg.FontStrikethru = picOriginal.Font.Strikethru
    cmnDlg.FontUnderline = picOriginal.Font.Underline
    
    cmnDlg.ShowFont
    
    picOriginal.Font.Bold = cmnDlg.FontBold
    picOriginal.Font.Italic = cmnDlg.FontItalic
    picOriginal.Font.Name = cmnDlg.FontName
    picOriginal.Font.size = cmnDlg.FontSize
    picOriginal.Font.Strikethru = cmnDlg.FontStrikethru
    picOriginal.Font.Underline = cmnDlg.FontUnderline
    
    txtChs_Change

End Sub

Private Sub cmdCreate_Click()
    
    Open App.Path + "\out.gzk" For Binary As #2
    Open App.Path + "\gbk.txt" For Binary As #1
    
    Dim c1 As Byte, c2 As Byte
    Dim FileLen As Long
    Dim s As String
    Dim xOffset As Long, yOffset As Long
    Dim w As Long, h As Long, i As Long, j As Long
    Dim b As Byte
    Dim size As Long
    Dim sec As Long, pot As Long
    
    
    h = picOriginal.Height
    w = picOriginal.Width \ 2 + picOriginal.Width Mod 2
    size = Val(txtSize)
    
    FileLen = LOF(1)
    
    xOffset = Val(txtXOffset)
    yOffset = Val(txtYOffset)
    
    gL = FileLen
    
    While Loc(1) < FileLen
        Get #1, , c1
        Get #1, , c2
        s = Chr("&H" + Hex(c1) + Hex(c2))
        DrawText xOffset, yOffset, s
        GetFont
        DoEvents
        sec = c1 - &HA0
        pot = c2 - &HA0
        Seek #2, (sec * 94 + pot) * size + 1
        For i = 0 To h - 1
            For j = 0 To w - 1
                b = Mat(i, j)
                Put #2, , b
            Next j
        Next i
        gP = Loc(1)
    Wend
    
    Close #2
    Close #1
    
    tmrProc.Interval = 0
    
    MsgBox "'out.gzk' successfully created."
    
    lblStatus = "Done."
End Sub

Private Sub GetFont()
    Dim X As Long
    Dim Y As Long
    Dim color As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim w As Long, h As Long
    Dim f As Boolean
    Dim depth As Long
    
    depth = 16
    h = picOriginal.Height
    w = picOriginal.Width \ 2 + picOriginal.Width Mod 2
    
    For Y = 0 To picOriginal.Height - 1
        f = True
        For X = 0 To picOriginal.Width - 1
            color = GetPixel(picOriginal.hdc, X, Y)
            r = color Mod 256
            g = (color \ 256) Mod 256
            b = color \ 256 \ 256
            a = (r + g + b) / 3
            a = a \ (256 \ depth)
            If f Then
                Mat(Y, X \ 2) = a
            Else
                Mat(Y, X \ 2) = Mat(Y, X \ 2) + a * depth
            End If
            f = Not f
        Next X
    Next Y
End Sub

Private Sub cmdGetAndTest_Click()

    Dim w As Long, h As Long
    Dim Y As Long, X As Long, a As Long
    Dim f As Boolean
    Dim depth As Long
    
    depth = 16
    
    h = picOriginal.Height
    w = picOriginal.Width \ 2 + picOriginal.Width Mod 2
    
    GetFont
    
    picOutput.Cls
    For Y = 0 To picOriginal.Height - 1
        f = True
        For X = 0 To picOriginal.Width - 1
            If f Then
                a = Mat(Y, X \ 2) Mod depth
            Else
                a = Mat(Y, X \ 2) / depth
            End If
            picOutput.PSet (X, Y), RGB(a * 256 / depth, a * 256 / depth, a * 256 / depth)
            f = Not f
        Next X
    Next Y
    
End Sub

Private Sub Command1_Click()
    'On Error Resume Next
    Open App.Path + "\out.gzk" For Binary As 1
    
    Dim sec As Long, pot As Long
    Dim w As Long, h As Long, i As Long, j As Long
    Dim b As Byte
    Dim Y As Long, X As Long, f As Boolean, depth As Long, a As Long
    Dim c1 As Long, c2 As Long
    
    depth = 16
    
    h = picOriginal.Height
    w = picOriginal.Width \ 2 + picOriginal.Width Mod 2
    
    c1 = &HEE
    c2 = &HF2
    
    sec = c1 - &HA0
    pot = c2 - &HA0
    
    Seek #1, (94 * sec + pot) * 128 + 1
    
    For i = 0 To h - 1
        For j = 0 To w - 1
            Get #1, , b
            Mat(i, j) = b
        Next j
    Next i
    
    picOutput.Cls
    
    For Y = 0 To picOriginal.Height - 1
        f = True
        For X = 0 To picOriginal.Width - 1
            If f Then
                a = Mat(Y, X \ 2) Mod depth
            Else
                a = Mat(Y, X \ 2) / depth
            End If
            picOutput.PSet (X, Y), RGB(a * 256 / depth, a * 256 / depth, a * 256 / depth)
            f = Not f
        Next X
    Next Y
    Close #1
    
End Sub

Private Sub Form_Load()
    txtWidth = "16"
    txtHeight = "16"
    txtChs = "ºº"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub tmrProc_Timer()
    lblStatus = CStr(gP * 100 \ gL) & "%"
End Sub

Private Sub txtChs_Change()
    DrawText Val(txtXOffset), Val(txtYOffset), Mid(txtChs, 1, 1)
End Sub

Private Sub DrawText(xOffset As Long, yOffset As Long, s As String)
    picOriginal.Cls
    TextOut picOriginal.hdc, xOffset, yOffset, s, 2
End Sub

Private Sub txtWidth_Change()
    Dim t As Long
    t = Val(txtWidth)
    If t = 0 Then
        t = 1
        txtWidth = "1"
    End If
    picOriginal.Width = t
    PicCenter
    RecalcSize
End Sub

Private Sub txtHeight_Change()
    Dim t As Long
    t = Val(txtHeight)
    If t = 0 Then
        t = 1
        txtHeight = "1"
    End If
    picOriginal.Height = t
    PicCenter
    RecalcSize
End Sub


Private Sub RecalcSize()
    Dim w As Long, h As Long
    
    h = picOriginal.Height
    w = picOriginal.Width \ 2 + picOriginal.Width Mod 2
    
    txtSize = w * h
End Sub

Private Sub txtXOffset_Change()
    txtChs_Change
End Sub

Private Sub txtYOffset_Change()
    txtChs_Change
End Sub
