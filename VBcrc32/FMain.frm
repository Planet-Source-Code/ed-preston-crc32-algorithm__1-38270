VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CRC32 Test Application"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcCRC32 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   2340
      TabIndex        =   2
      Top             =   60
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2235
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   2235
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   3075
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variable to hold the instance of the CRC32 Class
Private objCRC32 As clsCRC32

Const CHUNK_SIZE = 2048

' Calculate and display CRC Checksum of the file currently
' selected.
Private Sub cmdCalcCRC32_Click()
    Dim strTempPath As String
    
    ' Make sure the directory path is valid
    If Right$(Dir1.Path, 1) <> "\" Then
        strTempPath = Dir1.Path + "\"
    Else
        strTempPath = Dir1.Path
    End If
    
    ' Clear the current value
    lblDisplay.Caption = vbNullString
    
    ' Make sure a file has been selected.
    If File1.FileName <> vbNullString Then
        ' Get the checksum and update the display
        lblDisplay.Caption = CRCFromFile(strTempPath & File1.FileName)
    Else
        ' No file selected
        MsgBox "Please select a file to use.", vbInformation, App.Title
    End If
End Sub

' When the directory selection control changes update the file
' list control.
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

' When the drive selection changes update the directory selection
' control.
Private Sub Drive1_Change()
    On Error GoTo ErrorHandler
    Dir1.Path = Drive1.Drive & "\"
ErrorHandler:
End Sub

Private Sub Form_Load()
    ' Create an instance of the class
    Set objCRC32 = New clsCRC32
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cleanup
    Set objCRC32 = Nothing
End Sub

' ----------------------------
' Support Routines
' ----------------------------

' Class method accepts byte arrays, so we will need to read the file,
' turn it into a byte array, and pass to the method.  The return value
' is numeric but we want to display it as text.  We will have to convert
' the return value before returning the result.
Private Function CRCFromFile(ByVal strFilePath As String) As String
    Dim bArrayFile() As Byte
    Dim lngCRC32 As Long

    Dim lngChunkSize As Long
    Dim lngSize As Long

    lngSize = FileLen(strFilePath)
    lngChunkSize = CHUNK_SIZE

    If lngSize <> 0 Then

        ' Read byte array from file
        Open strFilePath For Binary Access Read As #1

        Do While Seek(1) < lngSize

            If (lngSize - Seek(1)) > lngChunkSize Then
                ' Process data in chunks. Chunky!
                Do While Seek(1) < (lngSize - lngChunkSize)
                    ReDim bArrayFile(lngChunkSize - 1)
                    Get #1, , bArrayFile()
                    lngCRC32 = objCRC32.CRC32(lngCRC32, bArrayFile, lngChunkSize - 1)
                Loop
            Else
                ' Blast it at them
                ReDim bArrayFile(lngSize - Seek(1))
                Get #1, , bArrayFile()
                
                lngCRC32 = objCRC32.CRC32(lngCRC32, bArrayFile, UBound(bArrayFile))
            End If

        Loop

        Close #1

        ' Everyone expects to view checksums in Hex strings.  Add buffer zeros if
        ' needed by smaller values.
        CRCFromFile = Right$("00000000" & Hex$(lngCRC32), 8)
    Else
        ' File of zero bytes has a CRC of 0
        CRCFromFile = "00000000"
    End If
End Function
