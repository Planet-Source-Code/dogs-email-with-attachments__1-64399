VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Email 
   CanGetFocus     =   0   'False
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   PropertyPages   =   "UserControl1.ctx":31E5
   ScaleHeight     =   630
   ScaleWidth      =   840
   ToolboxBitmap   =   "UserControl1.ctx":31F9
   Windowless      =   -1  'True
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'---------------------------------------------------------------------
' Email ocx Developed By Andy Hughes Email:andy@andythughes.co.uk
'
' I know this code isnt ground breaking but It will show other VB6 users
' how to create a user control, how to use property Pages, how to send Emails
' and more importantly how to add and send attachments.
'
' the code used to create the unicode file isnt mine, and I am unable to
' give proper credit to the original coder as I have had it in my module
' collection for a few years now and I cant remember who wrote it.
'
' to use add the component into the project, and drop onto the form
' you will see a nice little object called Email1
'
' use the code thus...
'
'    Email1.MailFrom = "Dogsbollox@topbloke.com"
'    Email1.MailMessage = "This is a test"
'    Email1.MailSubject = "This is a test"
'    Email1.MailTo = "andy@andythughes.co.uk"
'    Email1.PortNumber = 25
'    Email1.ServerName = "put your smtp server name here"
'    Email1.Attachment = "C:\andy test web\Newsletter-v1.jpg"
'    Email1.SendMail

' once the mail has been sent, an event is raised with the status
' you get many stages of failed, or Mail Sent Successfully
'
' to check on the event use the code below
'
'    Private Sub Email1_MailFailed(ByVal MessageStatus As Integer)

'       MsgBox "status " & Email1.DisplayError

'    End Sub
'---------------------------------------------------------------------

Private New_Server As String
Private New_Port As Integer
Private New_MailFrom As String
Private New_MailTo As String
Private New_MailSubject As String
Private New_MailMessage As String
Public ThisError As Boolean
Private CurrentStep As Integer
Private New_DisplayError As String
Private New_SendComplete As Integer
Private New_Attachment As String
Private AttachedFile As String


    Private Enum ErrorDisplay
        No_Server = 1
        No_Port = 2
        No_MailFrom = 3
        No_MailTo = 4
        No_MailMessage = 5
        Connect_To_Server = 6
        Send_Error = 7
        Disconnected_From_Server = 8
        Mail_Away_Successfully = 9
        Invalid_Attachment = 10
    End Enum

Public Event MailFailed(ByVal MessageStatus As Integer)


Public Property Get SendComplete() As Integer

    SendComplete = New_SendComplete

End Property

Public Property Let SendComplete(New_Value As Integer)

    New_SendComplete = New_Value
    PropertyChanged "SendComplete"

End Property

Public Property Get Attachment() As String
Attribute Attachment.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    Attachment = New_Attachment

End Property

Public Property Let Attachment(New_Value As String)

    New_Attachment = New_Value
    PropertyChanged "Attachment"


End Property



Public Property Get MailFrom() As String
Attribute MailFrom.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    MailFrom = New_MailFrom

End Property

Public Property Let MailFrom(New_Value As String)

    New_MailFrom = New_Value
    PropertyChanged "MailFrom"

End Property
Public Property Get DisplayError() As String
Attribute DisplayError.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    DisplayError = New_DisplayError

End Property

Public Property Let DisplayError(New_Value As String)

    New_DisplayError = New_Value
    PropertyChanged "DislayError"

End Property

Public Property Get MailTo() As String
Attribute MailTo.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    MailTo = New_MailTo

End Property

Public Property Let MailTo(New_Value As String)

    New_MailTo = New_Value
    PropertyChanged "MailTo"

End Property

Public Property Get MailSubject() As String
Attribute MailSubject.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    MailSubject = New_MailSubject

End Property

Public Property Let MailSubject(New_Value As String)

    New_MailSubject = New_Value
    PropertyChanged "MailSubject"

End Property

Public Property Get MailMessage() As String
Attribute MailMessage.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    MailMessage = New_MailMessage

End Property

Public Property Let MailMessage(New_Value As String)

    New_MailMessage = New_Value
    PropertyChanged "MailMessage"

End Property



Public Property Get ServerName() As String
Attribute ServerName.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    ServerName = New_Server

End Property

Public Property Let ServerName(New_Value As String)

    New_Server = New_Value
    PropertyChanged "ServerName"

End Property


Public Property Get PortNumber() As Integer
Attribute PortNumber.VB_ProcData.VB_Invoke_Property = "EmailProperty"

    PortNumber = New_Port

End Property

Public Property Let PortNumber(New_Value As Integer)

    New_Port = New_Value
    PropertyChanged "PortNumber"

End Property

Private Sub UserControl_Initialize()

    ThisError = False
End Sub

Private Sub UserControl_Resize()

    UserControl.Height = 630
    UserControl.Width = 840
    UserControl.Refresh

End Sub


Public Function SendMail()

    SendComplete = 0

    If Trim(ServerName) = "" Then
        ReportError (ErrorDisplay.No_Server)
        Exit Function
    End If
    
    If PortNumber = 0 Then
        ReportError (ErrorDisplay.No_Port)
        Exit Function
    End If
    
    If Trim(MailFrom) = "" Then
        ReportError (ErrorDisplay.No_MailFrom)
        Exit Function
    End If

    If Trim(MailTo) = "" Then
        ReportError (ErrorDisplay.No_MailTo)
        Exit Function
    End If

    If Trim(MailMessage) = "" Then
        ReportError (ErrorDisplay.No_MailMessage)
        ' can still continue even if no message body
    '    Exit Function
    End If

    ' next if there is an attachment we need to make sure it has a valid path
    If Trim(Attachment) <> "" Then
        If FileEx(Attachment) = False Then
            ReportError (ErrorDisplay.Invalid_Attachment)
            Exit Function
        End If
    End If
    
    ws.RemoteHost = ServerName
    ws.RemotePort = PortNumber                    ' Port 25
    ws.Connect                              ' Connect that shit up

End Function
Private Function ReportError(ByVal ErrorNum As Integer)


    Select Case ErrorNum
    Case 1
        DisplayError = "Invalid Mail Server"
        RaiseEvent MailFailed(2)
    Case 2
        DisplayError = "Invalid Port Number"
        RaiseEvent MailFailed(2)
    Case 3
        DisplayError = "Invalid Mail From"
        RaiseEvent MailFailed(2)
    Case 4
        DisplayError = "Invalid Mail Recipient"
        RaiseEvent MailFailed(2)
    Case 5
        DisplayError = "Mail Message Body Not Supplied"
        RaiseEvent MailFailed(2)
    Case 6
        DisplayError = "Connected To Mail Server"
    Case 7
        DisplayError = "Error Sending Mail"
        RaiseEvent MailFailed(2)
    Case 8
        DisplayError = "Disconnected From Mail Server"
    Case 9
        DisplayError = "Mail Sent Successfully"
        RaiseEvent MailFailed(1)
    Case 10
        DisplayError = "Invalid Attachment Specified"
        RaiseEvent MailFailed(2)
    End Select
    
End Function


Private Sub ws_Connect()
    
    CurrentStep = 0
    t1.Enabled = True

End Sub

Private Sub T1_Timer()

    On Error GoTo TimerError

    Select Case CurrentStep
    Case 0
        ws.SendData "HELO" + vbCrLf
        CurrentStep = CurrentStep + 1
    Case 1
        ' put mail code here
        ws.SendData "MAIL FROM:" & Trim(MailFrom) & vbCrLf
        CurrentStep = CurrentStep + 1
    Case 2
        ws.SendData "RCPT TO:" & Trim(MailTo) & vbCrLf
        CurrentStep = CurrentStep + 1
    Case 3
        ws.SendData "DATA" & vbCrLf
        CurrentStep = CurrentStep + 1
    Case 4
        If Trim(MailSubject) <> "" Then
            ws.SendData "SUBJECT:" & Trim(MailSubject) & vbCrLf
        End If
        CurrentStep = CurrentStep + 1
    Case 5
        ws.SendData vbCrLf
        CurrentStep = CurrentStep + 1
    Case 6
        ws.SendData vbCrLf
        CurrentStep = CurrentStep + 1
    Case 7
        ws.SendData MailMessage & vbCrLf
        CurrentStep = CurrentStep + 1
    Case 8
        If (Attachment) <> "" Then
            AttachedFile = Attach(Attachment)
            ws.SendData AttachedFile & vbCrLf
        End If
        CurrentStep = CurrentStep + 1
    Case 9
        ws.SendData "." & vbCrLf
        CurrentStep = CurrentStep + 1
    Case 10
        ws.Close
        ReportError (ErrorDisplay.Mail_Away_Successfully)
        t1.Enabled = False
    End Select

    ' if you want to add more recipients
    ' just change the case select statement to allow more of the RCPT
    ' commands to be added
       
  '  Case 3
  '      ws.SendData "RCPT TO:davidplant@hotmail.co.uk" & vbCrLf
  '      CurrentStep = CurrentStep + 1
  '  Case 4
  '      WS.SendData "RCPT TO:andy.hughes@the.co.uk" & vbCrLf
  '      CurrentStep = CurrentStep + 1

    
    Exit Sub
    
TimerError:
    ws.Close
    DisplayError = Err.Description
    RaiseEvent MailFailed(2)
    t1.Enabled = False

End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    DisplayError = Description
    RaiseEvent MailFailed(2)

End Sub

Private Function Attach(strFilePath As String) As String

    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    Dim strResult       As String
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    'This important: "begin 664"
    strResult = "begin 664 " + strFileName + vbLf
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize / 45 + 1
    'you need to encode every 45 bytes
    strFileData = Space(45)
    intFile = FreeFile
    'open the output file
    Open strFilePath For Binary As intFile
    For i = 1 To lEncodedLines
        If i = lEncodedLines Then
            strFileData = Space(lFileSize Mod 45)
        End If
        'get data
        Get intFile, , strFileData
        'the first byte in a line is a char, which number describes
        'how many bytes are in the line
        strTempLine = Chr(Len(strFileData) + 32)
        If i = lEncodedLines And (Len(strFileData) Mod 3) Then
            strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
        End If
        'now some encoding
        For j = 1 To Len(strFileData) Step 3
            strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
            strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                           + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
            strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                           + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
            strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
        Next j
        strResult = strResult + strTempLine + vbLf
        strTempLine = ""
        'get next line
    Next i
        'close the file
    Close intFile
    'add the "end" string
    strResult = strResult & "'" & vbLf + "end" + vbLf
    'return the encoded string
    Attach = strResult
End Function



Private Function FileEx(strPath As String) As Boolean

    FileEx = Not (Dir(strPath) = "")

End Function


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Attachment = PropBag.ReadProperty("Attachment", "")
    MailFrom = PropBag.ReadProperty("MailFrom", "")
    MailTo = PropBag.ReadProperty("MailTo", "")
    MailSubject = PropBag.ReadProperty("MailSubject", "")
    MailMessage = PropBag.ReadProperty("MailMessage", "")
    ServerName = PropBag.ReadProperty("ServerName", "")
    PortNumber = PropBag.ReadProperty("PortNumber", "25")
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Attachment", Attachment, "")
    Call PropBag.WriteProperty("MailFrom", MailFrom, "")
    Call PropBag.WriteProperty("MailTo", MailTo, "")
    Call PropBag.WriteProperty("MailSubject", MailSubject, "")
    Call PropBag.WriteProperty("MailMessage", MailMessage, "")
    Call PropBag.WriteProperty("ServerName", ServerName, "")
    Call PropBag.WriteProperty("PortNumber", PortNumber, "25")

End Sub



'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    Attachment = ""
    MailFrom = ""
    MailTo = ""
    MailSubject = ""
    MailMessage = ""
    ServerName = ""
    PortNumber = "25"

End Sub

