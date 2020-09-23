VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame #3"
      ForeColor       =   &H000000C0&
      Height          =   2220
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   4470
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame # 2"
      ForeColor       =   &H0000C000&
      Height          =   2220
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   4470
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2220
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   4470
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   2340
      Width           =   4605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
'its a good idea to place api declarations in alphabetical order to find one easier
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hbrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
 
 
Private Type POINTAPI
     x As Long
     y As Long
End Type
 
Dim hRgn(2) As Long 'the tabs region
Dim m_activeTab As Integer 'the tab that has current focus and is foreground
Private Const TAB_COUNT = 3 'how many tabs were using
 

Private Sub Form_Load()
    
    ' set properties that only need to be set once
    With Picture1
         .BorderStyle = 0
         .ScaleMode = vbPixels
         'helps minimize flickering
         .AutoRedraw = True
         'set the active tab if you wish
         Frame1(0).Visible = False
         Frame1(2).Visible = False
         m_activeTab = 1
    End With
    
End Sub
 
Private Sub FrameHide(showIndex As Integer)
  
  'hide all the frames
  Dim i As Integer
 
  For i = 0 To 2
     If showIndex <> i Then
          Frame1(i).Visible = False
    Else
          Frame1(i).Visible = True
    End If
  Next i
 
End Sub
 
Private Sub Form_Resize()
    
  Call Picture1_Paint
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  ' in creating regions your actually creating objects and you want
  ' to destroy all objects before exiting  otherwise your computer
  ' will start acting like a windows 98 system
  Dim i As Integer
  For i = 0 To (TAB_COUNT - 1)
     DeleteObject hRgn(i)
  Next i
  
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     
     'here our goal is to determine if we are mouseing down within a tabs region
     'or inbetween the tabs.
     Dim i As Integer
     Dim pt As POINTAPI
     
     'where the cursor is mouseing down at relative to the screen
     GetCursorPos pt
     
     'lets change the relativity of the mousedown to the picture box
     ScreenToClient Picture1.hwnd, pt
     
     'loop through the 3 regions that represent the tabs and check if
     'were mouseing down within one of them
     For i = 0 To (TAB_COUNT - 1)
         If PtInRegion(hRgn(i), pt.x, pt.y) Then
         
            'if we are AND the tab were mouseing down in has changed
            'the we need to repaint
            If m_activeTab <> i Then
                  m_activeTab = i
                  Call Picture1_Paint
                  
                  'code execution for tab click
                  Call FrameHide(i)
                  Debug.Print i
                  Exit For
                  
             End If
         End If
    Next i
    
End Sub

Private Sub Picture1_Paint()
  
     With Picture1
          ' erase the blackboard so we can redraw fresh
          .Cls
          ' always reference the scale width as apposed to width becuase
          ' scale width is compatible with whatever scale mode were using
          '(remember we set picture1 scalemode to pixels on form load)
          ' also scalewidth represents the actual paintable portion of the control
          ' and doesnt include things like a controls borders which arent paintable
          Dim workingwid As Integer
          workingwid = Picture1.ScaleWidth
          
          ' in this example were using 3 tabs
          Dim tabwid As Integer
          tabwid = (workingwid / TAB_COUNT)
          
          ' we need to "dip the brush" that will paint our tabs for us
          Dim i As Integer, hbrush As Long

          'first define the regions that make up the tabs
          For i = 0 To (TAB_COUNT - 1)
              hRgn(i) = CreateRoundRectRgn((i * tabwid), -.ScaleHeight, ((i + 1) * tabwid), (.ScaleHeight), 30, 50)
              
              If i = m_activeTab Then
                  'darker brush
                  hbrush = CreateSolidBrush(RGB(110, 110, 140))
              Else
                 'less dark brush
                  hbrush = CreateSolidBrush(RGB(150, 150, 170))
              End If
              
              'now paint the tab
              FrameRgn .hdc, hRgn(i), hbrush, 2, 2
          Next i
 
          '"clean the brush
          DeleteObject hbrush
          ' to show indication of which tab is active we want to paint horizontal line
          ' to the beginning of the currently active tab (which is set/changed when
          ' we click one of the tabs)...
          Picture1.Line (0, 0)-((m_activeTab * tabwid), 0), RGB(110, 110, 140)
          'we also want to paint a line from right edge of active tab to right edge of form
          Picture1.Line (((m_activeTab + 1) * tabwid), 0)-(.ScaleWidth, 0), RGB(110, 110, 140)
           
    End With
    
End Sub
