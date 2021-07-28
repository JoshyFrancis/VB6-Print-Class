VERSION 5.00
Begin VB.Form frmAPIPrint 
   AutoRedraw      =   -1  'True
   Caption         =   "API Print"
   ClientHeight    =   6192
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10572
   LinkTopic       =   "Form1"
   ScaleHeight     =   6192
   ScaleWidth      =   10572
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPreview 
      Caption         =   "Preview"
      Height          =   372
      Left            =   5280
      TabIndex        =   26
      Top             =   5760
      Value           =   1  'Checked
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unicode Printing"
      Height          =   492
      Left            =   6960
      TabIndex        =   25
      Top             =   5640
      Width           =   1572
   End
   Begin VB.HScrollBar HSPages 
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   735
      Left            =   9120
      Picture         =   "frmAPIPrint.frx":0000
      ScaleHeight     =   684
      ScaleWidth      =   1284
      TabIndex        =   23
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   9120
      Picture         =   "frmAPIPrint.frx":018A
      ScaleHeight     =   684
      ScaleWidth      =   1284
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeleteForm 
      Caption         =   "Delete Form"
      Height          =   495
      Left            =   9240
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddForm 
      Caption         =   "Add Form"
      Height          =   495
      Left            =   7920
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtcy 
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtcx 
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtFormName 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton cmdAPIPrintOK 
      Caption         =   "API Print OK"
      Height          =   495
      Left            =   7680
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdVBPrintProblem 
      Caption         =   "VB Print Problem"
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdVBPrint 
      Caption         =   "VB Print"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrinterSetup 
      Caption         =   "Printer Setup"
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cdmPageSetup 
      Caption         =   "Page Setup"
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox cboPaperSize 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   600
      Width           =   6375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set As Default"
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboPrinters 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "Print Preview"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   419
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Height(inches)"
      Height          =   195
      Left            =   5040
      TabIndex        =   18
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Width(inches)"
      Height          =   195
      Left            =   5040
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "New Form Name"
      Height          =   195
      Left            =   5040
      TabIndex        =   14
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label lblFormSize 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   " x"
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Forms"
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Printer"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmAPIPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Const HALFTONE = 4
Private Type SIZEAPI
    cX                  As Long
    cY                  As Long
End Type
Private Declare Function GetTextExtentPointW Lib "gdi32" (ByVal hdc As Long, ByVal lpszString As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long

Dim cp As New cPrinter
Dim Pages() As IPicture
Dim PageCount As Long
Private Sub cboPaperSize_Change()
If cboPaperSize.ListIndex = -1 Then Exit Sub
    '    cp.PaperSize = 11
    cp.PaperSize = cboPaperSize.ItemData(cboPaperSize.ListIndex)
Dim cX As Long, cY As Long
Const mMMPerInch As Single = 25.4
    Call cp.PrinterGetFormSize(cboPaperSize.Text, cX, cY)
lblFormSize.Caption = Format$(cX / mMMPerInch / 1000, "#.00") & " x " & Format$(cY / mMMPerInch / 1000, "#.00")
    txtFormName.Text = cboPaperSize.Text
    txtcx.Text = Format$(cX / mMMPerInch / 1000, "#.00")
    txtcy.Text = Format$(cY / mMMPerInch / 1000, "#.00")
End Sub

Private Sub cboPaperSize_Click()
cboPaperSize_Change
End Sub

Private Sub cboPrinters_Change()
    cp.Printer = cboPrinters.Text
'    cp.PaperSize = 11
        Picture1.Cls
        Picture1.Print "Server name :" & cp.ServerName
        Picture1.Print "Printer name :" & cp.Printer
        Picture1.Print "Share name :" & cp.ShareName
        Picture1.Print "Port name :" & cp.Port
        Picture1.Print "Driver name :" & cp.DriverName
        Picture1.Print "Comment :" & cp.Comment
        Picture1.Print "Location :" & cp.Location
        Picture1.Print "Print Processor :" & cp.PrintProcessor
        Picture1.Print "Default Data Type :" & cp.DefaultDataType
       
       
cboPaperSize.Clear
Dim c As Long, NumForms As Long, sNames() As String, cX() As Long, cY() As Long
    NumForms = cp.PrinterGetForms(sNames, cX, cY)
For c = 0 To NumForms - 1
    cboPaperSize.AddItem sNames(c)
    cboPaperSize.ItemData(cboPaperSize.NewIndex) = c + 1
'        If c + 1 = 11 Then
'            cboPaperSize.ListIndex = c
'        End If
Next
    If NumForms Then
'        cboPaperSize.ListIndex = 10
        For c = 0 To NumForms - 1
            If cboPaperSize.ItemData(c) = 11 Then
                cboPaperSize.ListIndex = c
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cboPrinters_Click()
cboPrinters_Change
End Sub


Private Sub cdmPageSetup_Click()
    cp.PageSetup
End Sub

Private Sub cmdAddForm_Click()
Const mMMPerInch As Single = 25.4
    cp.PrinterAddNewForm Val(txtcx.Text) * mMMPerInch * 1000, Val(txtcy.Text) * mMMPerInch * 1000, txtFormName.Text
cboPrinters_Change
End Sub


Private Sub cmdAPIPrintOK_Click()
Dim c As Long, n As Long
'    cp.Printer = cboPrinters.Text
'    cp.PaperSize = 11
If cp.PrinterStartDoc Then
        cp.PrinterStartPage
                    cp.PrintText "Page 1", cp.Width \ 2 - cp.TextWidth("Page 1") \ 2, 1
                n = 1
            For c = 1 To 100
                cp.PrintText "Line " & n, 4, (c - 1) * cp.TextHeight("A") + 1
                n = n + 1
            Next
        cp.PrinterEndPage
    cp.PrinterEndDoc
End If
End Sub

Private Sub cmdDeleteForm_Click()
    If MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
cp.PrinterDeleteForm txtFormName.Text
    cboPrinters_Change
End Sub


Private Sub cmdPrint_Click()
Dim c As Long, n As Long
'    cp.Printer = cboPrinters.Text
'    cp.PaperSize = 11
If cp.PrinterStartDoc Then
        cp.PrinterStartPage
                'cp.PrintText "Hello", 10, 10
            cp.Rectangle 1, 1, cp.Width - 2, cp.Height - 2
                    cp.PrintText "Page 1", cp.Width \ 2 - cp.TextWidth("Page 1") \ 2, 1
                    cp.Font.Bold = True
                    cp.Font.Italic = True
                    cp.PrintText "Bold Itelaic", 0, 0
            cp.PaintPicture Me.Icon, cp.Width \ 2 - (32), 32, 64, 64
                n = 1
            For c = 1 To (cp.Height \ cp.TextHeight("A"))
                cp.PrintText "Line " & n, 4, (c - 1) * cp.TextHeight("A") + 1
                n = n + 1
            Next
                cp.PaintPicture Picture3.Picture, cp.Width \ 2, 96
                cp.PaintPicture Picture2.Picture, 4, 128
                
        cp.PrinterEndPage
        cp.PrinterStartPage
                    
                    cp.PrintText "Page 2", cp.Width \ 2 - cp.TextWidth("Page 1") \ 2, 1
            For c = 1 To (cp.Height \ cp.TextHeight("A"))
                cp.PrintText "Line " & n, 4, (c - 1) * cp.TextHeight("A") + 1
                n = n + 1
            Next
        cp.PrinterEndPage
    cp.PrinterEndDoc
End If
End Sub

Private Sub cmdPrinterSetup_Click()
    cp.PrinterSetup
End Sub

Private Sub cmdPrintPreview_Click()
Dim c As Long, n As Long
'    cp.Printer = cboPrinters.Text
'    cp.PaperSize = 11
                Erase Pages
                PageCount = 0
        cp.Preview = True
If cp.PrinterStartDoc Then
        cp.PrinterStartPage
                'cp.PrintText "Hello", 10, 10
            cp.Rectangle 1, 1, cp.Width - 2, cp.Height - 2
                    cp.PrintText "Page 1", cp.Width \ 2 - cp.TextWidth("Page 1") \ 2, 1
            cp.PaintPicture Me.Icon, cp.Width \ 2 - (32), 32, 64, 64
                n = 1
            For c = 1 To (cp.Height \ cp.TextHeight("A"))
                cp.PrintText "Line " & n, 4, (c - 1) * cp.TextHeight("A") + 1
                n = n + 1
            Next
                cp.PaintPicture Picture3.Picture, cp.Width \ 2, 96
                cp.PaintPicture Picture2.Picture, 4, 128
                    Picture1.Cls
            cp.PaintTo Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight
                Picture1.Refresh
                    ReDim Preserve Pages(PageCount)
                        Set Pages(PageCount) = Picture1.Image
                        PageCount = PageCount + 1
        cp.PrinterEndPage
        cp.PrinterStartPage
                    cp.PrintText "Page 2", cp.Width \ 2 - cp.TextWidth("Page 1") \ 2, 1
            For c = 1 To (cp.Height \ cp.TextHeight("A"))
                cp.PrintText "Line " & n, 4, (c - 1) * cp.TextHeight("A") + 1
                n = n + 1
            Next
                    Picture1.Cls
            cp.PaintTo Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight
                Picture1.Refresh
                    ReDim Preserve Pages(PageCount)
                        Set Pages(PageCount) = Picture1.Image
                        PageCount = PageCount + 1
        cp.PrinterEndPage
    cp.PrinterEndDoc
End If
        cp.Preview = False
If PageCount > 0 Then
    HSPages.Min = 0
    HSPages.Max = PageCount - 1
    HSPages.Value = 0
    HSPages_Change
End If
End Sub

Private Sub cmdVBPrint_Click()
'Dim PageWidth As Single, PageHeight As Single
'    PageWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbPixels)
'    PageHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbPixels)
'        XFac = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
'        YFac = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
Dim c As Long, n As Long
    Dim prn As Printer
    For Each prn In Printers
        If prn.DeviceName = cboPrinters.Text Then Set Printer = prn: Exit For
    Next
    Set Printer.Font = Me.Font
'    Printer.PaperSize = 11
If cboPaperSize.ListIndex = -1 Then Exit Sub
    Printer.PaperSize = cboPaperSize.ItemData(cboPaperSize.ListIndex)
                Printer.CurrentX = Printer.ScaleWidth \ 2 - Printer.TextHeight("Page 1") \ 2
                Printer.CurrentY = 0
                Printer.Print "Page 1"
                Dim W As Single, H As Single
                    W = Printer.ScaleX(Me.Icon.Width, vbHimetric, Printer.ScaleMode)
                    H = Printer.ScaleY(Me.Icon.Height, vbHimetric, Printer.ScaleMode)
                    
            Printer.PaintPicture Me.Icon, Printer.ScaleWidth \ 2 - (W), H, W * 2, H * 2
                n = 1
            For c = 1 To (Printer.ScaleHeight \ Printer.TextHeight("A"))
                Printer.CurrentX = 0
                Printer.CurrentY = (c - 1) * Printer.TextHeight("A") '+ 1
                Printer.Print "Line " & n
                n = n + 1
            Next
                Printer.PaintPicture Picture3.Picture, Printer.ScaleWidth \ 2, Printer.ScaleY(128, 3, Printer.ScaleMode)
                Printer.PaintPicture Picture2.Picture, 10, Printer.ScaleY(256, 3, Printer.ScaleMode)
            
    Printer.NewPage
                Printer.CurrentX = Printer.ScaleWidth \ 2 - Printer.TextHeight("Page 1") \ 2
                Printer.CurrentY = 0
                Printer.Print "Page 2"
            For c = 1 To (Printer.ScaleHeight \ Printer.TextHeight("A"))
                Printer.CurrentX = 0
                Printer.CurrentY = (c - 1) * Printer.TextHeight("A") ' + 1
                Printer.Print "Line " & n
                n = n + 1
            Next
    Printer.EndDoc
End Sub

Private Sub cmdVBPrintProblem_Click()
Dim c As Long, n As Long
    Dim prn As Printer
    For Each prn In Printers
        If prn.DeviceName = cboPrinters.Text Then Set Printer = prn: Exit For
    Next
    Set Printer.Font = Me.Font
'    Printer.PaperSize = 11
If cboPaperSize.ListIndex = -1 Then Exit Sub
    Printer.PaperSize = cboPaperSize.ItemData(cboPaperSize.ListIndex)
                Printer.CurrentX = Printer.ScaleWidth \ 2 - Printer.TextHeight("Page 1") \ 2
                Printer.CurrentY = 0
                Printer.Print "Page 1"
                n = 1
            For c = 1 To 100
                Printer.CurrentX = 0
                Printer.CurrentY = (c - 1) * Printer.TextHeight("A") + 1
                Printer.Print "Line " & n
                n = n + 1
            Next
    Printer.EndDoc

End Sub

Private Sub Command1_Click()
Dim sz As SIZEAPI
                Erase Pages
                PageCount = 0
        cp.Preview = chkPreview.Value = 1 'True
If cp.PrinterStartDoc Then
        cp.PrinterStartPage
            SetBkMode cp.hdc, TRANSPARENT
                'cp.PrintText "Hello", 10, 10
            cp.Rectangle 1, 1, cp.Width - 2, cp.Height - 2
            cp.Rectangle 10, 10, cp.Width - 20, cp.Height - 20
        Dim f As New StdFont
            f.Name = "Times New Roman"
            f.Size = 50
            f.Bold = True
            f.Underline = True
                    
                    Set cp.Font = f
            Dim str As String
                str = "Unicode " + ChrW(3374) + ChrW(3378) + ChrW(3375) + ChrW(3390) + ChrW(3379) + ChrW(3330)
                GetTextExtentPointW cp.hdc, StrPtr(str), Len(str), sz
           TextOutW cp.hdc, cp.WidthPrint \ 2 - sz.cX \ 2, 40, StrPtr(str), Len(str)
            
'            BitBlt cp.hdc, 20, 100, 32, 32, Picture3.hdc, 0, 0, vbSrcCopy
            cp.PaintPicture Picture3.Picture, 20, 100
            cp.Rectangle 20, 140, 32, 32, vbRed, vbGreen
            
            f.Size = 80
            f.Bold = False
            f.Underline = False
                    Set cp.Font = f
                str = "Center"
                    GetTextExtentPointW cp.hdc, StrPtr(str), Len(str), sz
           TextOutW cp.hdc, cp.WidthPrint \ 2 - sz.cX \ 2, cp.HeightPrint \ 2 - sz.cY \ 2, StrPtr(str), Len(str)

            

            f.Size = 80
            f.Bold = False
            f.Underline = False
                    Set cp.Font = f
                str = "Bottom"
                    GetTextExtentPointW cp.hdc, StrPtr(str), Len(str), sz
           TextOutW cp.hdc, cp.WidthPrint \ 2 - sz.cX \ 2, (cp.HeightPrint - 20) - sz.cY, StrPtr(str), Len(str)


'            cp.PaintTo Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight
            SetStretchBltMode Picture1.hdc, 4
'            BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, cp.hdc, 0, 0, vbSrcCopy
            StretchBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, cp.hdc, 0, 0, cp.WidthPrint, cp.HeightPrint, vbSrcCopy
                Picture1.Refresh
                   
        cp.PrinterEndPage
    cp.PrinterEndDoc
End If
        cp.Preview = False
        
End Sub

Private Sub Command3_Click()
    cp.SetPrinterDefault cboPrinters.Text
End Sub

Private Sub Form_Load()
Dim numprinters As Long
Dim sPrinters() As String
Dim sPrinterDispNames() As String
Dim c As Long
        cp.Init Me
    sPrinters = cp.GetPrinteres(numprinters, sPrinterDispNames)
If numprinters > 0 Then
    For c = 0 To numprinters - 1
        cboPrinters.AddItem sPrinters(c)
    Next
        Erase sPrinters
        Erase sPrinterDispNames
End If
        cboPrinters.Text = cp.PrinterDefault
        
    cboPrinters_Change
Const mMMPerInch As Single = 25.4
    cp.PrinterAddNewForm 4 * mMMPerInch * 1000, 4 * mMMPerInch * 1000, "Small4x4"
    cboPrinters_Change
For c = 0 To cboPaperSize.ListCount - 1
    If cboPaperSize.List(c) = "Small4x4" Then
        cboPaperSize.ListIndex = c
        Exit For
    End If
Next

''  sc_Subclass Me.hwnd                                                       'Subclass a window... or three
'' sc_AddMsg Me.hwnd, ALL_MESSAGES, MSG_AFTER                         'Add messages of interest

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase Pages
Set cp = Nothing
'  sc_Terminate                                                              'Terminate all subclassing
End Sub


Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'DO NOT USE BREAKPOINT!!!!!
MsgBox "HWND:" & hWnd & ",MSG:" & uMsg & ",WPARAM:" & wParam & ",LPARAM:" & lParam

 WndProc = 1
End Function

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
'    Cls
'    Debug.Print IIf(bBefore, "Before", "After") & "&H" & Hex$(lng_hWnd) & "&H" & Hex$(uMsg) & "&H" & Hex$(wParam) & "&H" & Hex$(lParam) & "&H" & Hex$(lParamUser) & IIf(bBefore, vbNullString, "&H" & Hex$(lReturn))
End Sub
'-End Subclass callback, usually ordinal #1, the last method in this source file----------------------

Private Sub HSPages_Change()
    If PageCount Then
        Picture1.Cls
        Picture1.PaintPicture Pages(HSPages.Value), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    End If
End Sub

Private Sub HSPages_Scroll()
HSPages_Change
End Sub
 
