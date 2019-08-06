VERSION 5.00
Begin VB.Form OnnoRokomSms 
   Caption         =   "OnnoRokomSMS"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Sms_Text 
      Height          =   1095
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton Send_Button 
      Caption         =   "Click"
      Height          =   855
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Mobile_Number 
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "SMS Text"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Number"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "OnnoRokomSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function SMS()

    Dim DataToSend As String
    Dim objXML As Object
    
    Dim apiKey As String
    Dim mobileNumber As String
    Dim smsText As String
    Dim op As String
    Dim smsType As String
    Dim URL As String
    Dim maskName As String
    Dim campaignName As String
    
    apiKey = "" 'API KEY OnnoRokomSMS Panel
    
    mobileNumber = Mobile_Number.Text
    smsText = URLEncode(Sms_Text.Text)
    
    op = "NumberSms" 'OneToOne and OneToMany
    smsType = "TEXT"
    
    URL = "https://api2.onnorokomsms.com/HttpSendSms.ashx?"
    
    Set objXML = CreateObject("MSXML2.serverXMLHTTP")
    objXML.Open "POST", URL, False
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    objXML.send "op=" + op + "&apiKey=" + apiKey + "&type=" + smsType + "&mobile=" + mobileNumber + "&smsText=" + smsText + "&maskName=&campaignName"
    
    If Len(objXML.responseText) > 0 Then
            MsgBox objXML.responseText
    End If

End Function

Private Sub Send_Button_Click()

    Call SMS

End Sub

Function URLEncode(ByVal Text As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim char As String

    URLEncode = Text

    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next

End Function

Function GetBalance()
    
    Dim apiKey As String
    Dim op As String
    Dim URL As String
    
    apiKey = "" 'API KEY
    
    op = "GetCurrentBalance"
    
    URL = "https://api2.onnorokomsms.com/HttpSendSms.ashx?"
    
    Set objXML = CreateObject("MSXML2.serverXMLHTTP")
    objXML.Open "POST", URL, False
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    objXML.send "op=" + op + "&apiKey=" + apiKey
    
    If Len(objXML.responseText) > 0 Then
            MsgBox objXML.responseText
    End If
    
End Function
