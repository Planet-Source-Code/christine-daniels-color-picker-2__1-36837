VERSION 5.00
Begin VB.Form ColorPicker 
   Caption         =   "ColorPicker"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picKleur 
      Height          =   3345
      Left            =   3870
      ScaleHeight     =   3285
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   990
      Width           =   555
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   330
      LargeChange     =   10
      Left            =   90
      Max             =   360
      TabIndex        =   1
      Top             =   180
      Width           =   5910
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click right to copy value to clipboard"
      Height          =   285
      Left            =   405
      TabIndex        =   18
      Top             =   4680
      Width           =   3030
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Move mousepointer over colorfield."
      Height          =   285
      Left            =   405
      TabIndex        =   16
      Top             =   4455
      Width           =   3030
   End
   Begin VB.Label lblKleur 
      Height          =   195
      Index           =   0
      Left            =   90
      MousePointer    =   2  'Cross
      TabIndex        =   15
      Top             =   945
      Width           =   195
   End
   Begin VB.Label lblHUE 
      Alignment       =   2  'Center
      Caption         =   "HUE: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2475
      TabIndex        =   14
      Top             =   540
      Width           =   1140
   End
   Begin VB.Label lblHTML 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4500
      TabIndex        =   13
      Top             =   2430
      Width           =   1500
   End
   Begin VB.Label lblHSV 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5265
      TabIndex        =   12
      Top             =   3915
      Width           =   600
   End
   Begin VB.Label lblHSV 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5265
      TabIndex        =   11
      Top             =   3465
      Width           =   600
   End
   Begin VB.Label lblHSV 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5265
      TabIndex        =   10
      Top             =   3015
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4545
      TabIndex        =   9
      Top             =   3915
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4545
      TabIndex        =   8
      Top             =   3465
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4545
      TabIndex        =   7
      Top             =   3015
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4545
      TabIndex        =   6
      Top             =   1890
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4545
      TabIndex        =   5
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4545
      TabIndex        =   4
      Top             =   990
      Width           =   420
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5265
      TabIndex        =   3
      Top             =   1890
      Width           =   600
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5265
      TabIndex        =   2
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblRGB 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5265
      TabIndex        =   0
      Top             =   990
      Width           =   600
   End
   Begin VB.Menu CopyToClipboard 
      Caption         =   "Copy to Clipboard"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy value to Clipboard"
         Index           =   0
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&RGB-value"
         Index           =   2
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&html-value"
         Index           =   3
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "H&SV-value"
         Index           =   4
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
' Colorpicker by Christine Daniels
' maart 2002
'************************************************
Option Explicit
' de 3 RGB kleuren
Dim r As Double, g As Double, b As Double
' de HSV-waarden
Dim h As Double, s As Double, v As Double
'Toon de kleurwaarden?
Dim Toon As Boolean

'*********************************************
'het kleurvlak bestaat uit een array
' van labeltjes
' het aantal labeltjes wordt bepaald
' door de ArrayGrootte
' de afmetingen hoogte en breedte  (in pixels)
' van één labeltje de waarde AFM
'**********************************************
Const AFM = 3
Const ArrayGROOTTE = 75
' plaats (top en left-waarde)
' van het kleurvlak
Const LBOVEN = 70
Const LLINKS = 25

'***********************************
' toon de kleurwaarden in de labels
'***********************************
Private Sub ShowValues()
Dim rR As Integer, gg As Integer, bb As Integer
Dim ss As Double, vv As Integer
Dim i As Integer, j As Integer
Dim rij As Integer, kol As Integer

lblHSV(0).Caption = Str(h)

For i = 0 To lblKleur.Count - 1
        s = rij / ArrayGROOTTE
        v = kol / ArrayGROOTTE
        HSV2RGB h, s, v, r, g, b
        lblKleur(i).BackColor = RGB(r * 255, g * 255, b * 255)
        kol = kol + 1
        If kol > ArrayGROOTTE - 1 Then
            kol = 0
            rij = rij + 1
        End If
    Next

    lblHSV(1).Visible = False
    lblHSV(2).Visible = False

    lblRGB(0).Visible = False
    lblRGB(1).Visible = False
    lblRGB(2).Visible = False
    lblHTML.Visible = False
    picKleur.Visible = False
End Sub

'****************************
' initialisatie
'****************************
Private Sub Form_Load()
Dim i As Integer
Dim LINKS As Integer, BOVEN As Integer

AutoRedraw = True
lblKleur(0).Width = AFM
lblKleur(0).Height = AFM
lblKleur(0).Top = LBOVEN
lblKleur(0).Left = LLINKS
HScroll1.Value = 0
h = 0
For i = 1 To ArrayGROOTTE * ArrayGROOTTE - 1
    Load lblKleur(i)
    lblKleur(i).Visible = True
    LINKS = LINKS + AFM
    If LINKS >= ArrayGROOTTE * AFM Then
        LINKS = 0
        BOVEN = BOVEN + AFM
    End If
    lblKleur(i).Top = BOVEN + LBOVEN
    lblKleur(i).Left = LINKS + LLINKS
DoEvents
Next
picKleur.Top = LBOVEN
picKleur.Height = ArrayGROOTTE * AFM
ShowValues
End Sub

'*******************************************
' laelwaarden terug uitzetten
' als de muis het kleurveld verlaat
'*******************************************
Private Sub Form_MouseMove(Button As Integer, _
                            Shift As Integer, _
                            X As Single, Y As Single)
    lblHSV(1).Visible = False
    lblHSV(2).Visible = False
    
    lblRGB(0).Visible = False
    lblRGB(1).Visible = False
    lblRGB(2).Visible = False
    
    lblHTML.Visible = False
    picKleur.Visible = False
    Toon = False
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub


'***************************************
' de h-waarde scrollen
'***************************************
Private Sub HScroll1_Scroll()
Dim i As Integer, j As Integer

h = HScroll1.Value
lblHUE.Caption = "HUE: " + Format(h, "000")
If h > 359 Then
    h = h - 360
End If
ShowValues

End Sub


'**********************************
' popupmenu
'**********************************
Private Sub lblKleur_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu CopyToClipboard
End If
End Sub

'*******************************
' bewegen over het kleurvlak
'*******************************
Private Sub lblKleur_MouseMove(Index As Integer, _
            Button As Integer, Shift As Integer, _
            X As Single, Y As Single)
Dim rij As Integer
Dim kol As Integer
    

    lblHSV(1).Visible = True
    lblHSV(2).Visible = True
    
    lblRGB(0).Visible = True
    lblRGB(1).Visible = True
    lblRGB(2).Visible = True

    lblHTML.Visible = True
    picKleur.Visible = True

lblHSV(0).Caption = Str(h)

'bereken rij en kolom van aangeklikt label
kol = Index Mod ArrayGROOTTE
rij = (Index - kol) / ArrayGROOTTE
s = rij / (ArrayGROOTTE - 1)
v = kol / (ArrayGROOTTE - 1)

lblHSV(1).Caption = Format(s, "0.000")
lblHSV(2).Caption = Format(v, "0.000")

' HSV omrekenen naar RGB
HSV2RGB h, s, v, r, g, b

lblRGB(0).Caption = Str(Round(r * 255))
lblRGB(1).Caption = Str(Round(g * 255))
lblRGB(2).Caption = Str(Round(b * 255))
lblHTML.Caption = "#" & Dec2Hex(Round(r * 255)) _
                    & Dec2Hex(Round(g * 255)) _
                    & Dec2Hex(Round(b * 255))

' toon de kleur in het kleurvlak
picKleur.BackColor = RGB(Round(r * 255), Round(g * 255), Round(b * 255))
End Sub


'********************************
' kopiëren naar het clipboard
'********************************
Private Sub mnucopy_Click(Index As Integer)
Dim ss As String
Select Case Index
    Case 2: Clipboard.Clear
            ss = Str(Round(r * 255)) & "," & Str(Round(g * 255)) & "," & Str(Round(b * 255))
            Clipboard.SetText (ss)
            
    Case 3: Clipboard.Clear
            ss = "#" & Dec2Hex(Round(r * 255)) & Dec2Hex(Round(g * 255)) & Dec2Hex(Round(b * 255))
            Clipboard.SetText (ss)
            
    Case 4: Clipboard.Clear
            ss = Str(h) & "," & Format(s, "0.000") & "," & Format(v, "0.000")
            Clipboard.SetText (ss)
End Select
End Sub
