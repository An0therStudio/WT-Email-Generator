VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "Walkthrough Email Generator"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   OleObjectBlob   =   "FrmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public INIPath
Public WTTelNo
Public WTEmail
Public WTStartDay
Public WTEndDay
Public WTStartTime
Public WTEndTime
Public WTAHTelNo
Public WTAHEmail
Public intLegs As Integer
Public blRecordings As Boolean

Private Sub UserForm_Activate()
    arrUserName = Split(ReturnUserName(), " ")
    txtName.Text = arrUserName(0)
    blRecordings = False
    INIPath = "\\YGK01CFP01\Operations\Call Execution\WT Team\Macros\EmailGenerator\EmailGenerator.ini"
    LoadINIData
End Sub

Private Sub BtnClearAll_Click()
    TxtDataSheet.Text = ""
End Sub

Private Sub BtnSettings_Click()
    FrmSettings.Show
    If FrmSettings.ChkSave.Value = True Then
        WTTelNo = FrmSettings.TxtTelNo.Text
        WTEmail = FrmSettings.TxtEmail.Text
        WTStartDay = FrmSettings.TxtStartDay.Text
        WTEndDay = FrmSettings.TxtEndDay.Text
        WTStartTime = FrmSettings.TxtStartTime.Text
        WTEndTime = FrmSettings.TxtEndTime.Text
        WTAHTelNo = FrmSettings.TxtAHTelNo.Text
        WTAHEmail = FrmSettings.TxtAHEmail.Text
        SaveINIData
    End If
    

End Sub

Private Sub BtnGenerate_Click()
'On Error Resume Next
    Dim EmailContent As String
' First paragraph:
    EmailContent = "<html><head><style>body {font-family:tahoma; line-height: 1; font-size:11pt}</style></head><body><p>"
    EmailContent = EmailContent & "Hello,<br><br>"
    EmailContent = EmailContent & "My name is " & txtName.Text
    EmailContent = EmailContent & " from the Conferencing Center.  I am contacting you regarding a conference call that you have scheduled for <span style='font-weight:bold'>"
    EmailContent = EmailContent & GetDate() & "</span> with ID number <span style='font-weight:bold'>" & GetConfId() & ".</span>  "
    EmailContent = EmailContent & "Please review the conference details and highlighted sections below to ensure everything is set up correctly. You may then answer the <span style='background-color:yellow'>highlighted</span> questions, <span style='font-weight:bold'>update and highlight any changes</span> you would like to make and send it to "
    EmailContent = EmailContent & "<a href=" & WTEmail & ">" & WTEmail & "</a>.<br><br><br>"
    EmailContent = EmailContent & "<span style='background-color:#66ff33; font-size:13pt; font-weight:bold' >Over the phone:</span><br><br>"
    EmailContent = EmailContent & "If you would like to review the details over the phone please fill out the below “Scheduled Walkthrough Request” template and send it to "
    EmailContent = EmailContent & "<a href=" & WTEmail & ">" & WTEmail & "</a>.<br><br>"
    
' Create the Table
    EmailTable = "<table style='border:2px solid black;border-collapse: collapse;font-family:cambria;'><tr><td style='width:400px;text-align:center;border:1px solid black;padding:5px;background-color:#fcf010' colspan=2><span style='font-size:115%;font-weight:bold;'>Scheduled Walkthrough Request</span></td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-weight:bold;'>Conference ID</span></td><td style='border:1px solid black;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#8ac5ff'>"
    EmailTable = EmailTable & "</td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#d7f1ff'><span style='font-weight:bold;'>Contact Name</span></td><td style='border:1px solid black;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#d7f1ff'>"
    EmailTable = EmailTable & "</td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-weight:bold;'>Contact Number</span></td><td style='border:1px solid black;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#8ac5ff'>"
    EmailTable = EmailTable & "</td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#d7f1ff'><span style='font-weight:bold;'>Requested Date</span></td><td style='border:1px solid black;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#d7f1ff'>"
    EmailTable = EmailTable & "</td></tr><tr><td style='border:1px solid black;width:150px;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#8ac5ff'><span style='font-weight:bold;'>Requested Time*</span></td><td style='border:1px solid black;padding-left:15px;padding-top:5px;padding-bottom:5px;background-color:#8ac5ff'>"
    EmailTable = EmailTable & "</td></tr><tr><td style='width:400px;text-align:center;border:1px solid black;padding:5px;background-color:#fcf010' colspan=2><span style='font-size:85%;font-weight:bold;'>*please be sure to include time zone.</span></td></tr></table><br>"
    EmailContent = EmailContent & EmailTable
    EmailContent = EmailContent & "<span style='background-color:#66ff33; font-size:16pt; font-weight:bold' >Or</span><br>"
    EmailContent = EmailContent & "<span style='background-color:#66ff33; font-size:13pt; font-weight:bold' >Via email:</span><br><br>"

' Basic details:
    BasicDetails = "<span style='font-weight:bold'>" & GetCompanyName() & " Conference Call:</span><br>"
    BasicDetails = BasicDetails & "<span style='font-weight:bold'>Call Start Time:</span> " & GetCallTime() & "<br>"
    BasicDetails = BasicDetails & "<span style='font-weight:bold'>Duration:</span> " & GetDuration() & " minutes" & "<br>"
    BasicDetails = BasicDetails & "<span style='font-weight:bold'>Call Topic:</span> " & GetTopic() & "<br><br>"
    EmailContent = EmailContent & BasicDetails

' Speaker list:
    SpeakerList = "<span style='font-weight:bold'>Approved Speakers List on the Call: </span>" & GetSpeakerList()
    SpeakerList = SpeakerList & "<ul>"
    SpeakerList = SpeakerList & "<li><span style='background-color:yellow'>Will presenters dial in on a single line or separately?</span>" & Space(1)
    SpeakerList = SpeakerList & "<li><span style='background-color:yellow'>Approximately how early will presenters begin to dial-in?</span>" & Space(1)
    SpeakerList = SpeakerList & "<li><span style='background-color:yellow'>Will there be any additional persons joining the speaker line?</span>" & Space(1)
    SpeakerList = SpeakerList & "</ul>"
    EmailContent = EmailContent & SpeakerList

' Comm line Details:
    If GetCommScheduled() = True Then
        CommList = "<span style='font-weight:bold'>Communications Line Dialing In:</span> " & GetCommList() & "<br>"
        CommList = CommList & "<ul>"
        CommList = CommList & "<li><span style='background-color:#b3e6ff'>A communications line is back-end communication with a secondary operator - separate from your lead operator -  that your designated communications line contact can speak with during your conference without call interruption. A communications line should be  someone who is *not* a speaker on your call and often dial in from a different room than the presenters.</span>" & Space(1)
        CommList = CommList & "<li><span style='background-color:yellow'>Purpose of Communications Line for your call:</span>" & Space(1)
        CommList = CommList & "</ul>"
        EmailContent = EmailContent & CommList
    End If

' Expected parts Details:
    ExpectedParts = "<span style='font-weight:bold'>Total Participants Expected:</span> " & GetTotalLegs() & "<br>"
    EmailContent = EmailContent & ExpectedParts

' Dial-in numbers:
    EmailContent = EmailContent & vbCrLf & GenerateDialInNums() & vbCrLf
    
' Call Type:
    EmailContent = EmailContent & "<span style='font-weight:bold'>Call Type:</span><br>" & GenerateCallType() & "<br>"

' Call Features:
    CallFeatures = "<span style='font-weight:bold'>Call Features:</span><br>"
    CallFeatures = CallFeatures & "Please ensure all Special Scripts, Special Annunciators, Approved Participants Lists and/or Polling Questions are emailed to our reservations department at: reservations@teleconferencingcenter.com at least 24 hours prior to call start time.<br><br>"
    CallFeatures = CallFeatures & GetEntry()            ' Entry type
    CallFeatures = CallFeatures & GetVoiceTalent()      ' Voice Talent
    CallFeatures = CallFeatures & GetSpecialized()      ' Specialized operator
    CallFeatures = CallFeatures & GetLeaderView()       ' Leaderview?
    CallFeatures = CallFeatures & GetPartEntry()        ' What parts will reference.
    CallFeatures = CallFeatures & GetPartReport()       ' Participant report.
    CallFeatures = CallFeatures & GetSpecialScript()    ' Special script.
    CallFeatures = CallFeatures & GetSpecialAnn()       ' Special Annunciator.
    CallFeatures = CallFeatures & GetAPL()              ' APL
    EmailContent = EmailContent & CallFeatures
    
' Additional Call Features:
    AddFeatures = AddFeatures & GetSilentRecord()       ' Silent record line
    AddFeatures = AddFeatures & GetIWS()                ' Web Audio streaming
    AddFeatures = AddFeatures & GetPolling()            ' Polling
    AddFeatures = AddFeatures & GetRollCall()           ' Roll Call
    AddFeatures = AddFeatures & GetMonitor()            ' Call Monitor
    AddFeatures = AddFeatures & GetAnnounceAll()        ' Announce all
    AddFeatures = AddFeatures & GetTones()              ' Entry/Exit Tones
    ' Dial out
    If AddFeatures <> "" Then
        EmailContent = EmailContent & "<br><br><span style='font-weight:bold'>Additional Call Features:</span><br>"
        EmailContent = EmailContent & AddFeatures
    End If
    
' Recording and transcription services:
    Recordings = "<br><br><span style='font-weight:bold'>Recording & Transcription services:</span><br>"
    Recordings = Recordings & GetRecordings()           ' Recordings Requested.
    Recordings = Recordings & GetEncore()               ' Encore
    'Check for Recordings
    If blRecordings = False Then
        Recordings = Recordings & ChrW(&H2612) & "<span style='background-color:yellow'>NO Recordings requested</span><br>" & Space(1)
    End If
    Recordings = Recordings & GetTranscription()        ' call transcription
    EmailContent = EmailContent & Recordings & "<br><br>"
    
' Final message:
    EmailContent = EmailContent & "If you have any questions or concerns, please do not hesitate to contact us " & WTStartDay & " to " & WTEndDay & " between the hours of " & WTStartTime & " and " & WTEndTime & " eastern time via email at " & WTEmail & " or via phone at " & WTTelNo & ". If outside these hours and urgent, you may also contact us via email at " & WTAHEmail & " or via phone at " & WTAHTelNo & ".<br><br><br> Thank you and have a great day,<br><br></p>"
' Sign Off
    UserSignature = ""
    UserSignature = "<span style='font-weight:bold; color:#002b80'>" & ReturnUserName() & "</span><br>" & Space(1)
    UserSignature = UserSignature & "<span style='color:##808080; font-size:95%'>Walkthrough Support</span><br>" & Space(1)
    UserSignature = UserSignature & "<span style='color:#002b80; font-size:95%'>----------------------------------------</span><br>" & Space(1)
    UserSignature = UserSignature & "<span style='color:##808080; font-size:95%'>866-528-4699 - Toll Free<br></span>" & Space(1)
    UserSignature = UserSignature & "<a style='font-size:95%' href=" & WTEmail & ">" & WTEmail & "</a><br><br><br><br>"
    EmailContent = EmailContent & UserSignature
    EmailContent = EmailContent & "</body></html>"
' Create email message:
    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)
    With objMsg
      .Subject = "Conference Detail Walkthrough for CID " & GetConfId()
      .SentOnBehalfOfName = WTEmail
      .BodyFormat = olFormatHTML
      .HTMLBody = EmailContent
      .Display

    End With
    Set objMsg = Nothing
    
End Sub

Private Sub BtnGenShort_Click()
    'Gather Required information:
    theCID = InputBox(Prompt:="Please enter the CID of the call.", Title:="Short WT Email")
    theDate = InputBox(Prompt:="Please enter the date of the call.", Title:="Short WT Email")
    ' Create the First Paragraph.
    EmailContent = "Hello,<br><br>My name is " & txtName.Text
    EmailContent = EmailContent & " and I am contacting you regarding a conference call that you have scheduled for "
    EmailContent = EmailContent & theDate & " with the ID number listed above.  "
    EmailContent = EmailContent & "I wanted to quickly go over the conference call details with you to ensure that everything was set up correctly and to answer any questions that you may have."
    EmailContent = EmailContent & "<br><br>"
    EmailContent = EmailContent & "If you would like to go over your call details, please leave a message - including your name, contact number, conference ID number and a preferred date & time to review - on our Walkthrough Support Line by dialing "
    EmailContent = EmailContent & WTTelNo & " or you may complete the form below and reply directly to this email to schedule a time with one of our Walkthrough Specialists."
    EmailContent = EmailContent & "<br><br>"
    EmailContent = EmailContent & "Our Walkthrough Support Line hours are " & WTStartDay & " to " & WTEndDay & " between the hours of " & WTStartTime & " and " & WTEndTime & " eastern time. If you have an urgent request outside these hours, you may call " & WTAHTelNo & " or send an email to " & WTAHEmail & "."
    EmailContent = EmailContent & "<br><br>"
    ' Create the Table CSS Style
    EmailTableStyle = "<style> table, th, td { border: 1px solid black; border-collapse: collapse; }</style>"
    ' Create the Table
    EmailTable = "<table cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>"
    EmailTable = EmailTable & "<tr><td colspan='2' valign=top style='width:255.35pt;border:solid windowtext 2.25pt;border-bottom:solid windowtext 1.0pt;background:#4F81BD;padding:0in 5.4pt 0in 5.4pt'>"
    EmailTable = EmailTable & "<u><b><span style='color:white'><p align=center style='text-align:center'>SCHEDULED WALKTHROUGH REQUEST</p></b></u></span></td></tr>"
    EmailTable = EmailTable & "<tr><td valign=top style='width:133.85pt;border-top:none;border-left:solid windowtext 2.25pt;border-bottom:solid windowtext 1.0pt;border-right:none;background:#4F81BD;padding:0in 5.4pt 0in 5.4pt'>"
    EmailTable = EmailTable & "<span style='color:white'><small><b>Contact Name:</b></small></span></td>"
    EmailTable = EmailTable & "<td valign=top style='width:121.5pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 2.25pt;background:#A7BFDE;padding:0in 5.4pt 0in 5.4pt'></td></tr>"
    EmailTable = EmailTable & "<tr><td valign=top style='width:133.85pt;border-top:none;border-left:solid windowtext 2.25pt;border-bottom:solid windowtext 1.0pt;border-right:none;background:#4F81BD;padding:0in 5.4pt 0in 5.4pt'>"
    EmailTable = EmailTable & "<span style='color:white'><small><b>Contact Number:</b></small></span></td>"
    EmailTable = EmailTable & "<td valign=top style='width:121.5pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 2.25pt;background:#D3DFEE;padding:0in 5.4pt 0in 5.4pt'></td>"
    EmailTable = EmailTable & "<tr><td valign=top style='width:133.85pt;border-top:none;border-left:solid windowtext 2.25pt;border-bottom:solid windowtext 2.25pt;border-right:none;background:#4F81BD;padding:0in 5.4pt 0in 5.4pt'>"
    EmailTable = EmailTable & "<span style='color:white'><small><b>Date & Time* to Review:</b></small></span></td>"
    EmailTable = EmailTable & "<td valign=top style='width:121.5pt;border-top:none;border-left:none;border-bottom:solid windowtext 2.25pt;border-right:solid windowtext 2.25pt;background:#A7BFDE;padding:0in 5.4pt 0in 5.4pt'></td></tr>"
    EmailTable = EmailTable & "</table>"
    ' Compile Email body.
    HTMLBody = "<html><head>"
    HTMLBody = HTMLBody & EmailTableStyle
    HTMLBody = HTMLBody & "</head><body>"
    HTMLBody = HTMLBody & EmailContent
    HTMLBody = HTMLBody & "<br>" & EmailTable
    HTMLBody = HTMLBody & "<small><i>*please be sure to include time zone.</i></small><br>"
    HTMLBody = HTMLBody & "<br><br>Thank you very much and I look forward to hearing from you soon!"
    UserSignature = ""
    UserSignature = "<br><br><span style='font-weight:bold; color:#002b80'>" & ReturnUserName() & "</span><br>" & Space(1)
    UserSignature = UserSignature & "<span style='color:##808080; font-size:95%'>Walkthrough Support</span><br>" & Space(1)
    UserSignature = UserSignature & "<span style='color:#002b80; font-size:95%'>----------------------------------------</span><br>" & Space(1)
    UserSignature = UserSignature & "<span style='color:##808080; font-size:95%'>866-528-4699 - Toll Free<br></span>" & Space(1)
    UserSignature = UserSignature & "<a style='font-size:95%' href=" & WTEmail & ">" & WTEmail & "</a><br><br><br><br>"
    HTMLBody = HTMLBody & UserSignature
    HTMLBody = HTMLBody & "</body></html>"
    
    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)
    With objMsg
      .Subject = "Conference Detail Walkthrough for CID " & theCID
      .SentOnBehalfOfName = WTEmail
      .BodyFormat = olFormatHTML
      .HTMLBody = HTMLBody
      .Display
    End With
    Set objMsg = Nothing
    
End Sub

Function GetTranscription()
    RetStr = ""
    hasTranscription = RunGeneralRegex("Call Transcription: (\w+)", TxtDataSheet.Text)
    If hasTranscription <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Call Transcription<br>"
        Reply = MsgBox(Prompt:="Is the entire call being transcribed?", Buttons:=vbYesNo, Title:="Transcription")
        If Reply = vbYes Then
            RetStr = RetStr & ChrW(&H2612) & " Entire call"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Q&A Session only<br>"
        Else
            'RetStr = RetStr & ChrW(&H2610) & " Entire call"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Q&A Session only<br>"
        End If
        'MsgBox hasTranscription
        If hasTranscription = "Standard" Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Standard: 48 hour business day turnaround" & "<br>"
        'Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Standard: 48 hour business day turnaround"
        End If
        'RetStr = RetStr & "<br>"
        If hasTranscription = "Priority" Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Priority Service: 24 hour business day turnaround" & "<br>"
        'Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Priority Service: 24 hour business day turnaround"
       End If
        'RetStr = RetStr & "<br>"
        If hasTranscription = "Express" Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Express Service: 12 hour business day turnaround" & "<br>"
        'Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Express Service: 12 hour business day turnaround"
        End If
        'RetStr = RetStr & "<br>"
        If hasTranscription = "Rush" Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Rush Service: 3 hour business day turnaround" & "<br>"
        'Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Rush Service: 3 hour business day turnaround"
        End If
        '*************************
        'RetStr = RetStr & "<br>"
        'CapInfo = ""
        'CapInfo = InputBox(Prompt:="Who is the recipient of the report? (First and Last Name)", Title:="Transcription Recipient(s)")
        'RetStr = RetStr & "<ul><li>Transcription Recipient(s):<ul><li>Name:&nbsp;" & CapInfo
        'CapInfo = ""
        'CapInfo = InputBox(Prompt:="What email will receive the report?", Title:="Transcription Recipient(s)")
        'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;Email:&nbsp;" & CapInfo & "</li>"
        'Reply = ""
        'Reply = MsgBox(Prompt:="Are there multiple recipients?", Buttons:=vbYesNo, Title:="Transcription Recipient(s)")
        'blMulti = False
        'If Reply = vbYes Then
        '   blMulti = True
        'Else
        '   blMulti = False
        'End If
        '*************************
        RetStr = RetStr & "<ul><li>Transcription Recipient(s):<ul>"
        Do
            blMulti = False
            CapInfo = ""
            CapInfo = InputBox(Prompt:="Who is the recipient of the report? (First and Last Name)", Title:="Transcription Recipient(s)")
            RetStr = RetStr & "<li>Name:&nbsp;" & CapInfo
            CapInfo = ""
            CapInfo = InputBox(Prompt:="What email will receive the report?", Title:="Transcription Recipient(s)")
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;Email:&nbsp;" & CapInfo & "</li>"
            Reply = ""
            Reply = MsgBox(Prompt:="Are there additional recipients?", Buttons:=vbYesNo, Title:="Transcription Recipient(s)")
            If Reply = vbYes Then
                blMulti = True
            Else
                blMulti = False
            End If
        Loop While blMulti = True
        RetStr = RetStr & "</ul></ul>"
    End If
    GetTranscription = RetStr
End Function

Function GetEncore()
    RetStr = ""
    hasEncore = RunGeneralRegex("(Encore Settings:)", TxtDataSheet.Text)
    If hasEncore <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Encore playback<br>"
        StartDate = RunGeneralRegex("Start Date: ([0-9/]+)", TxtDataSheet.Text)
        StartTime = RunGeneralRegex("Start Date: [0-9/]+\s+(\d\d:\d\d)", TxtDataSheet.Text)
        EndDate = RunGeneralRegex("End Date: ([0-9/]+)", TxtDataSheet.Text)
        EndTime = RunGeneralRegex("End Date: [0-9/]+\s+(\d\d:\d\d)", TxtDataSheet.Text)
        RetStr = RetStr & "Starting Date: " & StartDate & "&nbsp;&nbsp;&nbsp;&nbsp;Time: " & StartTime & "<br>"
        RetStr = RetStr & "Ending Date: " & EndDate & "&nbsp;&nbsp;&nbsp;&nbsp;Time: " & EndTime & "<br>"
        RetStr = RetStr & "Encore Playback Numbers: Toll Free: 800-585-8367 or 855-859-2056 Toll: 404-537-3406<br><br>"
        
        reportRequest = MsgBox(Prompt:="Does this call have an encore report?", Buttons:=vbYesNo, Title:="Encore")
        
        'ReportRequest = RunGeneralRegex("Report: (Yes|No)", TxtDataSheet.Text) ' ICRM Does not show correctly.
        If reportRequest = vbYes Then
            RetStr = RetStr & ChrW(&H2612) & "Encore Report<br>"
            CapInfo = InputBox(Prompt:="Please enter the information captured for the encore report.", Title:="Encore")
            priorPrompt = MsgBox(Prompt:="Does this call have a prior prompt?", Buttons:=vbYesNo, Title:="Encore")
            afterPrompt = MsgBox(Prompt:="Does this call have an after prompt?", Buttons:=vbYesNo, Title:="Encore")
            If priorPrompt = vbYes Then
                RetStr = RetStr & ChrW(&H2612) & "Prior Prompt"
            Else
                'RetStr = RetStr & ChrW(&H2610) & "Prior Prompt"
            End If
            
            If afterPrompt = vbYes Then
                RetStr = RetStr & ChrW(&H2612) & "&nbsp;&nbsp;&nbsp;&nbsp;After Prompt<br>"
            Else
                'RetStr = RetStr & ChrW(&H2610) & "&nbsp;&nbsp;&nbsp;&nbsp;After Prompt<br>"
            End If
            RetStr = RetStr & "<ul><li>Information Captured: " & CapInfo
            RetStr = RetStr & "<li>Encore Report Recipient(s):<ul>"
        Do
            blMulti = False
            CapInfo = ""
            CapInfo = InputBox(Prompt:="Who is the recipient of the report? (First and Last Name)", Title:="Encore Report Recipient(s)")
            RetStr = RetStr & "<li>Name:&nbsp;" & CapInfo
            CapInfo = ""
            CapInfo = InputBox(Prompt:="What email will receive the report?", Title:="Encore Report Recipient(s)")
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;Email:&nbsp;" & CapInfo & "</li>"
            Reply = ""
            Reply = MsgBox(Prompt:="Are there additional recipients?", Buttons:=vbYesNo, Title:="Encore Report Recipient(s)")
            If Reply = vbYes Then
                blMulti = True
            Else
                blMulti = False
            End If
        Loop While blMulti = True
        RetStr = RetStr & "</ul></ul>"
        Else
            RetStr = RetStr & ChrW(&H2612) & "Encore Report - NOT REQUESTED<br>"
        End If
        blRecordings = True
    End If
    GetEncore = RetStr
End Function

Function GetRecordings()
    RetStr = ""
    isRecording = RunGeneralRegex("Audio Recording: (\w+)", TxtDataSheet.Text)
    If isRecording <> "ERROR NO MATCH FOUND" Then ' There is recordings
        'RetStr = RetStr & ChrW(&H2612) & " Recordings requested<br>"
        isFTP = RunGeneralRegex("(https://ftp.icallinc.com)", TxtDataSheet.Text)
        If isFTP <> "ERROR NO MATCH FOUND" Then
            Reply = vbNo ' Found ftp site.
        Else
            Reply = MsgBox(Prompt:="Is the recording being mailed to the customer?", Buttons:=vbYesNo, Title:="Recordings")
        End If
        
        If Reply = vbYes Then
            ' CD Mailed out, gather address info.
            RetStr = RetStr & GetMailingInfo()
        Else
            ' Recording uploaded to FTP, gather FTP info.
            RetStr = RetStr & GetFTPInfo()
        End If
        blRecordings = True
    End If
    GetRecordings = RetStr
End Function

Function GetMailingInfo()
    RetStr = ""
    RetStr = RetStr & ChrW(&H2612) & " CD Record<br>"
    ' Record format:
    RetStr = RetStr & GetRecordFormat()
    ' Quantity of CD's to ship.
    qty = RunGeneralRegex("Tape/CD Type:\s+[0-9a-zA-Z -]+\s+Qty:\s+(\d)", TxtDataSheet.Text)
    RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;Qty:" & qty & "<br>"
    'Shipping method
    shipMethod = RunGeneralRegex("Shipping Method: ([0-9a-zA-Z -]+)", TxtDataSheet.Text)
    RetStr = RetStr & "Shipping Method: " & shipMethod & "<br>" ' TODO: FIX SO THERE ARE BOXES WITH OPTIONS.
    mailAddr = InputBox(Prompt:="Please enter the Address:", Title:="CD Mailing Information")
    mailCity = InputBox(Prompt:="Please enter the City:", Title:="CD Mailing Information")
    mailState = InputBox(Prompt:="Please enter the State / Province:", Title:="CD Mailing Information")
    mailZip = InputBox(Prompt:="Please enter the Zip / Postal Code:", Title:="CD Mailing Information")
    RetStr = RetStr & "Mailing Address: " & mailAddr & "<br>"
    RetStr = RetStr & "City: " & mailCity & "&nbsp;&nbsp;&nbsp;&nbsp;State/Province: " & mailState & "<br>"
    RetStr = RetStr & "Zip/Postal Code: " & mailZip & "<br><br><br>"
    
    GetMailingInfo = RetStr
End Function

Function GetFTPInfo()
    RetStr = ""
    RetStr = RetStr & ChrW(&H2612) & " FTP Upload<br>"
    ' Record Format
    RetStr = RetStr & "" & GetRecordFormat() & "<br>"
    ' Ftp Site
    FTPSite = RunGeneralRegex("(https://ftp.icallinc.com)", TxtDataSheet.Text)
    If FTPSite = "ERROR NO MATCH FOUND" Then         ' Reuse the data captured.
        FTPSite = InputBox(Prompt:="Please Enter the ftp site.", Title:="FTP Infomation")
    End If
    RetStr = RetStr & "Website: " & FTPSite & "<br>"
    RetStr = RetStr & "<ul><li>Recording Recipients:"
    RetStr = RetStr & "<ul>"
    Do
        blMulti = False
        CapInfo = ""
        CapInfo = InputBox(Prompt:="What email will receive the report?", Title:="FTP Recipient(s)")
        RetStr = RetStr & "<li>Email:&nbsp;" & CapInfo & "</li>"
        Reply = ""
        Reply = MsgBox(Prompt:="Are there additional recipients?", Buttons:=vbYesNo, Title:="FTP Recipient(s)")
        If Reply = vbYes Then
            blMulti = True
        Else
            blMulti = False
        End If
    Loop While blMulti = True
    RetStr = RetStr & "</ul></ul>"
    GetFTPInfo = RetStr
End Function

Function GetRecordFormat()
    RetStr = ""
    recordFormat = RunGeneralRegex("Tape/CD Type: ([0-9a-zA-Z -]+)\s+Qty:", TxtDataSheet.Text)
    RetStr = RetStr & "Format: " & recordFormat
    GetRecordFormat = RetStr
End Function

Function GetIWS()
    RetStr = ""
    isIWS = RunGeneralRegex("(- Audio Streaming:)", TxtDataSheet.Text)
    If isIWS <> "ERROR NO MATCH FOUND" Then
        webcastID = RunGeneralRegex("Audio Webcast Id: (\d+)", TxtDataSheet.Text)
        RetStr = RetStr & ChrW(&H2612) & " Audio Only Web Streaming – <span style='background-color:yellow'>Will you be using this feature on the live call and/or for recordings after the call?</span>" & Space(1)
        RetStr = RetStr & "<ul>"
        If webcastID <> "ERROR NO MATCH FOUND" Then
            RetStr = RetStr & "<li> Webcast ID: " & webcastID
            archiveDate = InputBox(Prompt:="Please enter the IWS archived until date.", Title:="IWS Streaming")
            RetStr = RetStr & "<li> Archived Until: " & archiveDate
        End If
        RetStr = RetStr & "</ul>"
    End If
    GetIWS = RetStr
End Function

Function GetTones()
    RetStr = ""
    isTones = RunGeneralRegex("Entry / Exit Tones: (\w+)", TxtDataSheet.Text)
    If isTones = "Both" Then
        RetStr = RetStr & ChrW(&H2612) & " Entry Tones<br>"
        RetStr = RetStr & ChrW(&H2612) & " Exit Tones<br>"
    ElseIf isTones = "Entry" Then
        RetStr = RetStr & ChrW(&H2612) & " Entry Tones<br>"
    ElseIf isTones = "Exit" Then
        RetStr = RetStr & ChrW(&H2612) & " Exit Tones<br>"
    End If
    GetTones = RetStr
End Function

Function GetAnnounceAll()
    RetStr = ""
    isAnn = RunGeneralRegex("(Announce)", TxtDataSheet.Text)
    If isAnn <> "ERROR NO MATCH FOUND" Then
        RetStr = ChrW(&H2612) & " Announce All<br>"
    End If
    GetAnnounceAll = RetStr

End Function

Function GetMonitor()
    RetStr = ""
    isMon = RunGeneralRegex("(Call Monitor)", TxtDataSheet.Text)
    If isMon <> "ERROR NO MATCH FOUND" Then
        RetStr = ChrW(&H2612) & " Call Monitor<br>"
    End If
    GetMonitor = RetStr
End Function

Function GetRollCall()
    RetStr = ""
    isRollCall = RunGeneralRegex("(Rollcall)", TxtDataSheet.Text)
    If isRollCall <> "ERROR NO MATCH FOUND" Then
        RetStr = ChrW(&H2612) & " Roll Call<br>"
    End If
    GetRollCall = RetStr
End Function

Function GetSilentRecord()
    Reply = MsgBox(Prompt:="Is There a silent record line required for this call?", Buttons:=vbYesNo, Title:="Silent Record Line")
    If Reply = vbYes Then
        RetStr = ChrW(&H2612) & " Silent Record Line<br>"
    Else
        RetStr = ""
    End If
    GetSilentRecord = RetStr
End Function

Function GetAPL()
    RetStr = ""
    isAPL = RunGeneralRegex("(Approved Participant List)", TxtDataSheet.Text)
    If isAPL <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Approved Participant List"
        ' If apl received
        Reply = MsgBox(Prompt:="Has the participant list been received?", Buttons:=vbYesNo, Title:="Approved Participants")
        If Reply = vbYes Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Received"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Not yet received<br>"
        Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Received"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Not yet received<br>"
        End If
        
        ' Reason for apl:
        Reply = MsgBox(Prompt:="Is the APL used to determine if participants can join the call?", Buttons:=vbYesNo, _
                                Title:="Approved Participants")
        If Reply = vbYes Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Join Call<br>"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Q&A Session Only<br>"
        Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Join Call<br>"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Q&A Session Only<br>"
        End If
    
    Else
        RetStr = ""
    End If
    GetAPL = RetStr
End Function

Function GetPolling()
    RetStr = ""
    isPolling = RunGeneralRegex("(Polling:)", TxtDataSheet.Text)
    If isPolling <> "ERROR NO MATCH FOUND" Then
        RetStr = ChrW(&H2612) & " Polling<br>&nbsp;&nbsp;&nbsp;&nbsp;Questions/Responses:&nbsp;&nbsp;&nbsp;&nbsp;"
        Reply = MsgBox(Prompt:="Have the polling questions / responses been received?", Buttons:=vbYesNo, Title:="Polling")
        If Reply = vbYes Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Received"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Not yet received"
        Else
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Not yet received"
        End If
    Else
        RetStr = ""
    End If
    GetPolling = RetStr
End Function

Function GetSpecialAnn()
    RetStr = ""
    isAnnunciator = RunGeneralRegex("(Special Annunciator)", TxtDataSheet.Text)
    If isAnnunciator <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Special Annunciator"
        Reply = MsgBox(Prompt:="Has the special annunciator been received?", Buttons:=vbYesNo, Title:="Special Annunciator")
        If Reply = vbYes Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Received"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Not yet received"
        Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Received"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Not yet received"
        End If
    Else
        RetStr = ""
    End If
    GetSpecialAnn = RetStr
End Function

Function GetSpecialScript()
    RetStr = ""
    isSpecialScript = RunGeneralRegex("(Special Script)", TxtDataSheet.Text)
    If isSpecialScript <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Special Script "
        Reply = MsgBox(Prompt:="Has the special script been received?", Buttons:=vbYesNo, Title:="Special Script")
        If Reply = vbYes Then
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Received"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Not yet received"
        Else
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Received"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Not yet received"
        End If
    Else
        RetStr = ""
    End If
    GetSpecialScript = RetStr
End Function

Function GetPartReport()
    RetStr = ""
    factsComplete = RunGeneralRegex("(Facts Complete)", TxtDataSheet.Text)
    If factsComplete <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & "Participant report"
        gatheredInfo = InputBox(Prompt:="Please enter the information gathered.", Title:="Facts complete information")
        RetStr = RetStr & "<ul><li>Information captured: " & gatheredInfo
        ' Search for recipients:
        RetStr = RetStr & "<li>Report Recipient(s):<ul>"
        For Each Line In Split(TxtDataSheet.Text, vbCrLf)
            Recipient = RunGeneralRegex("(.+ null)", CStr(Line))
            
            If Recipient <> "ERROR NO MATCH FOUND" And Recipient <> "Fax Number: null" And Recipient <> "Comments: null" Then
                RecipientData = Split(Recipient, " ")
                RecipientName = RecipientData(0) & " " & RecipientData(1)
                RecipientEmail = RecipientData(2)
                RetStr = RetStr & "<li>Name: " & RecipientName & "&nbsp;&nbsp;&nbsp;&nbsp;Email: " & RecipientEmail
            End If
        Next
        RetStr = RetStr & "</ul></ul>"
    Else
        RetStr = RetStr & ChrW(&H2612) & "<span style='background-color:yellow'>Participant report - NO PARTICIPANT REPORT REQUESTED.</span><br>" & Space(1)
    End If
    GetPartReport = RetStr
End Function

Function GetPassword()
    RetStr = ""
    PWord = RunGeneralRegex("Password: (.+)", TxtDataSheet.Text)
    If PWord <> "ERROR NO MATCH FOUND" Then
        RetStr = PWord
    Else
        RetStr = "NO PASSWORD FOUND"
    End If
    GetPassword = RetStr
End Function

Function GetPartEntry()
    RetStr = ""
    TollFreeEPPart = RunGeneralRegex("EP Participant (\d+)", TxtDataSheet.Text)
    If TollFreeEPPart <> "ERROR NO MATCH FOUND" Then
        ' Is an ep
        RetStr = ChrW(&H2612) & " Event Plus - Automated Entry<br>"
        Passcode = RunGeneralRegex("Passcode: (\d+)\r", TxtDataSheet.Text)
        If Passcode <> "ERROR NO MATCH FOUND" Then  ' EP with passcode
            RetStr = RetStr & ChrW(&H2612) & " Participants will reference:<br>"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Conference ID Number<br>"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Passcode: <span style='background-color:yellow'>" & Passcode & "</span>" & Space(1) & "<br>"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Other:<br>"
        Else
            RetStr = RetStr & ChrW(&H2612) & " Participants will reference: NO PASSCODE - DIRECT ENTRY<br>" ' Direct entry EP
            
        End If
    Else
        'is an oa
        RetStr = ChrW(&H2612) & " Participants will reference:<br>"
        Password = GetPassword()
        If Password <> "NO PASSWORD FOUND" Then
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Conference ID Number<br>"
            RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Password: <span style='background-color:yellow'>" & Password & "</span>" & Space(1) & "<br>"
            'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Other:<br>"
        Else
            Reply = MsgBox(Prompt:="Will Participants be referencing Conference ID for entry?", Buttons:=vbYesNo, _
                            Title:="Paricipant Entry")
            If Reply = vbYes Then
                RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Conference ID Number<br>"
                'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Password:<br>"
                'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Other:<br>"
            Else
                otherStr = InputBox(Prompt:="What will Participants reference to join the call?", Title:="Participant reference entry", _
                                        Default:="Topic? etc....")
                'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Conference ID Number<br>"
                'RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2610) & " Password:<br>"
                RetStr = RetStr & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChrW(&H2612) & " Other: " & otherStr & "<br>"
            End If
        End If
    End If
    GetPartEntry = RetStr
End Function

Function GetLeaderView()
    RetStr = ""
    LeaderView = RunGeneralRegex("(Leader View: Option On)", TxtDataSheet.Text)
    If LeaderView <> "ERROR NO MATCH FOUND" Then
        WebPin = RunGeneralRegex("User Web Pin: (\d+)", TxtDataSheet.Text)
        RetStr = ChrW(&H2612) & " Leader View&nbsp;&nbsp;&nbsp;&nbsp;Web Pin: #" & WebPin & "<span style='background-color:yellow'>  Will you be using Leaderview?</span><br>" & Space(1)
    End If
    GetLeaderView = RetStr
End Function

Function GetSpecialized()
    RetStr = ""
    Specialized = RunGeneralRegex("(Specialized Operator)", TxtDataSheet.Text)
    If Specialized <> "ERROR NO MATCH FOUND" Then
        RetStr = ChrW(&H2612) & " Specialized Operator<br>"
    End If
    GetSpecialized = RetStr
End Function

Function GetVoiceTalent()
    RetStr = ""
    VoiceTalent = RunGeneralRegex("(Voice Talent)", TxtDataSheet.Text)
    If VoiceTalent <> "ERROR NO MATCH FOUND" Then
        Language = RunGeneralRegex("Agent Role: Lead Operator Language: (.+)", TxtDataSheet.Text)
        If Language = "ERROR NO MATCH FOUND" Then
            Language = "English"
        End If
        RetStr = ChrW(&H2612) & " Voice Talent: " & Language & "<br>"
    End If
    GetVoiceTalent = RetStr
End Function

Function GetEntry()
    RetStr = ""
    EntryType = RunGeneralRegex("Participant Entry: (\w+)", TxtDataSheet.Text)
    If EntryType = "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Music Hold<br>"
        RetStr = RetStr & ChrW(&H2612) & " Silent Hold<br>"
        RetStr = RetStr & ChrW(&H2612) & " News Hold<br>"
    ElseIf EntryType = "Music" Then
        RetStr = RetStr & ChrW(&H2612) & " Music Hold<br>"
        'RetStr = RetStr & ChrW(&H2610) & " Silent Hold<br>"
    ElseIf EntryType = "Silent" Then
        'RetStr = RetStr & ChrW(&H2610) & " Music Hold<br>"
        RetStr = RetStr & ChrW(&H2612) & " Silent Hold<br>"
    ElseIf EntryType = "Direct" Then
        'RetStr = RetStr & ChrW(&H2612) & " Music Hold<br>"
        'RetStr = RetStr & ChrW(&H2610) & " Silent Hold<br>"
        RetStr = RetStr & ChrW(&H2612) & " Direct Entry<br>"
    ElseIf EntryType = "News" Then
        'RetStr = RetStr & ChrW(&H2612) & " Music Hold<br>"
        'RetStr = RetStr & ChrW(&H2610) & " Silent Hold<br>"
        RetStr = RetStr & ChrW(&H2612) & " News Hold<br>"
    Else
        RetStr = "ERROR IN GET ENTRY"
        MsgBox EntryType
    End If
    GetEntry = RetStr
End Function

Function GenerateCallType()
    RetStr = ""
    IS_QA = RunGeneralRegex("- (Q&A):", TxtDataSheet.Text)
    IS_LEC = RunGeneralRegex("- (Lecture):", TxtDataSheet.Text)
    If IS_QA <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & ChrW(&H2612) & " Question and Answer (all participant lines muted during presentation; lines opened individually for Q&A)"
        RetStr = RetStr & "<ul><li><span style='background-color:yellow'>Will there be One Q&A session near the end of the Call or multiple Q&As throughout the call?</span></ul>" & Space(1)
        'RetStr = RetStr & ChrW(&H2610) & " Lecture Call (all participant lines throughout the call muted)<br>"
        'RetStr = RetStr & ChrW(&H2610) & " Open Call (all lines open)<br>"
    ElseIf IS_LEC <> "ERROR NO MATCH FOUND" Then
        'RetStr = RetStr & ChrW(&H2610) & " Question and Answer (all participant lines muted during presentation; lines opened individually for Q&A)"
        'RetStr = RetStr & "<ul><li><span style='background-color:yellow'>Will there be One Q&A session near the end of the Call or multiple Q&As throughout the call?</span></ul>"
        RetStr = RetStr & ChrW(&H2612) & " Lecture Call (all participant lines throughout the call muted)<br>"
        'RetStr = RetStr & ChrW(&H2610) & " Open Call (all lines open)<br>"
    Else
        'RetStr = RetStr & ChrW(&H2610) & " Question and Answer (all participant lines muted during presentation; lines opened individually for Q&A)"
        'RetStr = RetStr & "<ul><li><span style='background-color:yellow'>Will there be One Q&A session near the end of the Call or multiple Q&As throughout the call?</span></ul>"
        'RetStr = RetStr & ChrW(&H2610) & " Lecture Call (all participant lines throughout the call muted)<br>"
        If intLegs > 35 Then
            RetStr = RetStr & ChrW(&H2612) & "<span style='background-color:yellow'> Open Call (all lines open)<b> <-- Is this correct?</b></span>" & Space(1) & "<br>"
        Else
            RetStr = RetStr & ChrW(&H2612) & " Open Call (all lines open) <br>"
        End If
    End If
    GenerateCallType = RetStr
End Function

Function GenerateDialInNums()
    RetStr = "<b>Dial-In Numbers:</b><br><ul>"
'Toll Free Leader:
    TollFreeLeader = RunGeneralRegex("MM Leader (\d+)", TxtDataSheet.Text)
    If TollFreeLeader <> "ERROR NO MATCH FOUND" Then
        TollFreeLeader = HyphenateNumber(TollFreeLeader)
        RetStr = RetStr & "<li><b>Toll Free Leader Dial-In Number:</b> " & TollFreeLeader
    End If
'Toll Leader:
    TollLeader = RunGeneralRegex("LM Leader (\d+)", TxtDataSheet.Text)
    If TollLeader <> "ERROR NO MATCH FOUND" Then
        TollLeader = HyphenateNumber(TollLeader)
        RetStr = RetStr & "<li><b>Local/Toll Leader Dial-in Number:</b> " & TollLeader
    End If
'Toll Free Comm:
    TollFreeComm = RunGeneralRegex("MM Comm Line (\d+)", TxtDataSheet.Text)
    If TollFreeComm <> "ERROR NO MATCH FOUND" Then
        TollFreeComm = HyphenateNumber(TollFreeComm)
        RetStr = RetStr & "<li><b>Toll Free Communications Dial-In Number:</b> " & TollFreeComm
    End If
'Toll Comm:
    TollComm = RunGeneralRegex("LM Comm Line (\d+)", TxtDataSheet.Text)
    If TollComm <> "ERROR NO MATCH FOUND" Then
        TollComm = HyphenateNumber(TollComm)
        RetStr = RetStr & "<li><b>Local/Toll Communications Dial-In Number:</b> " & TollComm
    End If
'EP Participant:
    TollFreeEPPart = RunGeneralRegex("EP Participant (\d+)", TxtDataSheet.Text)
    If TollFreeEPPart <> "ERROR NO MATCH FOUND" Then
        TollFreeEPPart = HyphenateNumber(TollFreeEPPart)
        RetStr = RetStr & "<li><b>EP Toll Free Participant Dial-in Number:</b> " & TollFreeEPPart
        TollEPPart = RunGeneralRegex("LM Participant (\d+)", TxtDataSheet.Text)
        If TollEPPart <> "ERROR NO MATCH FOUND" Then
            TollEPPart = HyphenateNumber(TollEPPart)
            RetStr = RetStr & "<li><b>EP Local/Toll Participant Dial-in Number:</b> " & TollEPPart
        End If
        TollEPPart = RunGeneralRegex("EL Participant (\d+)", TxtDataSheet.Text)
        If TollEPPart <> "ERROR NO MATCH FOUND" Then
            TollEPPart = HyphenateNumber(TollEPPart)
            RetStr = RetStr & "<li><b>EP Local/Toll Participant Dial-in Number:</b> " & TollEPPart
        End If
' OA Participant:
    Else
        TollFreeOAPart = RunGeneralRegex("MM Participant (\d+)", TxtDataSheet.Text)
        If TollFreeOAPart <> "ERROR NO MATCH FOUND" Then
            TollFreeOAPart = HyphenateNumber(TollFreeOAPart)
            RetStr = RetStr & "<li><b>Toll Free Participant Dial-in Number:</b> " & TollFreeOAPart
        End If
        TollOAPart = RunGeneralRegex("LM Participant (\d+)", TxtDataSheet.Text)
        If TollOAPart <> "ERROR NO MATCH FOUND" Then
            TollOAPart = HyphenateNumber(TollOAPart)
            RetStr = RetStr & "<li><b>Local/Toll Participant Dial-in Number:</b> " & TollOAPart
        End If
    End If
' OA Press only participant:
    'TODO: ADD PRESS ONLY PARTS HERE!!!!!!!!!!!!!
' EP Pass code:
    Passcode = RunGeneralRegex("Passcode: (\d+)\r", TxtDataSheet.Text)
    If Passcode <> "ERROR NO MATCH FOUND" Then
        RetStr = RetStr & "<li><b>EP Passcode:</b> <span style='background-color:yellow'>" & Passcode & "</span>" & Space(1)
    End If
' ITFS Numbers:
    ITFSUsed = RunGeneralRegex("ITFS/ITL: (\w+)", TxtDataSheet.Text)
    If ITFSUsed <> "ERROR NO MATCH FOUND" Then
        ' Test if it's yes or no.
        RetStr = RetStr & "<li><b>ITFS (International Toll Free Service) Numbers:</b> " & ITFSUsed
    Else
        ' This should never happen.
        RetStr = RetStr & "<li>ITFS THIS SHOULD NEVER HAPPEN"
    End If
    RetStr = RetStr & "</ul>"
    GenerateDialInNums = RetStr
End Function

Function GetTotalLegs()
    RetStr = RunGeneralRegex("Legs: (\d+)", TxtDataSheet.Text)
    intLegs = CInt(RetStr)
    GetTotalLegs = RetStr
End Function

Function GetCommScheduled()
    CommNeeded = RunGeneralRegex("(Comm Line)", TxtDataSheet.Text)
    If CommNeeded = "ERROR NO MATCH FOUND" Then
        GetCommScheduled = False
    Else
        GetCommScheduled = True
    End If
End Function

Function GetCommList()
    Dim CommArray() As String
    Dim counter As Integer
    counter = 0
    For Each Line In Split(TxtDataSheet.Text, vbCrLf)
        CommName = RunGeneralRegex("^C (\w+, \w+)", CStr(Line))
        If CommName <> "ERROR NO MATCH FOUND" Then
            LastFirst = Split(CommName, ", ")
            CommName = LastFirst(1) & " " & LastFirst(0)
            ReDim Preserve CommArray(counter)
            CommArray(counter) = CommName
            counter = counter + 1
        End If
    Next
    On Error GoTo ErrHandler
    If UBound(CommArray) > 0 Then
        GetCommList = Join(CommArray, ", ")
    Else
        GetCommList = CommArray(0)
    End If
    Exit Function
ErrHandler:
    GetCommList = "<span style='background-color:yellow'>Will you be using?</span><br>" & Space(1)
End Function

Function GetSpeakerList()
    Dim SpeakerArray() As String
    Dim counter As Integer
    counter = 0
    For Each Line In Split(TxtDataSheet.Text, vbCrLf)
'        SpkName = RunGeneralRegex("^\* (\w+, \w+)", CStr(Line)) ' Old doesn't work for \. or (
        SpkName = RunGeneralRegex("^\* ([0-9a-zA-Z \-'(.]+, [0-9a-zA-Z \-'(.]+) (EP|MM|DO)", CStr(Line))
        If SpkName <> "ERROR NO MATCH FOUND" Then
            LastFirst = Split(SpkName, ", ")
            SpkName = LastFirst(1) & " " & LastFirst(0)
            ReDim Preserve SpeakerArray(counter)
            SpeakerArray(counter) = SpkName
            counter = counter + 1
        End If
    Next
    If UBound(SpeakerArray) > 0 Then
        GetSpeakerList = Join(SpeakerArray, ", ")
    Else
        GetSpeakerList = SpeakerArray(0)
    End If
End Function

Function GetTopic()
    Dim RetStr As String
    RetStr = RunGeneralRegex("Topic: ([\s\S]+)Call Type:", TxtDataSheet.Text)
    If RetStr <> "ERROR NO MATCH FOUND" Then
        RetStr = Replace(RetStr, vbCrLf, " ") ' Remove newline if in topic.
        RetStr = Trim(RetStr)
    Else
        RetStr = "n/a"
    End If
    GetTopic = RetStr
End Function

Function GetDuration()
    Dim RetStr
    RetStr = RunGeneralRegex("Duration: (.+)", TxtDataSheet.Text)
    GetDuration = RetStr
End Function

Function GetCallTime()
    Dim RetStr
    RetStr = RunGeneralRegex("Call Time: (.+)", TxtDataSheet.Text)
    GetCallTime = RetStr
End Function

Function GetCompanyName()
    Dim RetStr
    RetStr = RunGeneralRegex("Company: (.+)", TxtDataSheet.Text)
    GetCompanyName = RetStr
End Function

Function GetConfId()
    Dim RetStr
    RetStr = RunGeneralRegex("Conference ID: (.+)", TxtDataSheet.Text)
    GetConfId = RetStr
End Function

Function GetDate()
    Dim RetStr
    RetStr = RunGeneralRegex("Call Date: (.+)", TxtDataSheet.Text)
    GetDate = RetStr
End Function

Function ReturnUserName()
    Dim strUserName As String
    strUserName = ""
    'Create user name from object
    arrUserName = Split(Application.Session.CurrentUser.Name, ",")
    If UBound(arrUserName) >= 1 Then
        strUserName = Trim(arrUserName(1)) & Space(1) & Trim(arrUserName(0))
        arrUserName = Split(strUserName, " ")
        If UBound(arrUserName) >= 1 Then
            strUserName = Trim(arrUserName(0)) & Space(1) & Trim(arrUserName(UBound(arrUserName)))
        End If
    Else
        arrUserName = Split(Application.Session.CurrentUser.Name, " ")
        If UBound(arrUserName) >= 1 Then
            strUserName = Trim(arrUserName(0)) & Space(1) & Trim(arrUserName(UBound(arrUserName)))
        End If
    End If
    ReturnUserName = strUserName
End Function

' Run a given regex on the given text.
Function RunGeneralRegex(regex As String, input_data As String)
    Dim ConfDate As String
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    Set objRegExp = New RegExp
    objRegExp.Pattern = regex
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    If (objRegExp.test(input_data)) Then
        Set colMatches = objRegExp.Execute(input_data)
        For Each objMatch In colMatches
            RetStr = objMatch.SubMatches(0)
            If Right(RetStr, 1) = vbCrLf Or Right(RetStr, 1) = vbCr Then ' Trim any trailing newlines
                RetStr = Left(RetStr, (Len(RetStr) - 1))
            End If
        Next
    Else
        RetStr = "ERROR NO MATCH FOUND"
    End If
    RunGeneralRegex = RetStr
End Function

Sub SaveINIData()
    'MsgBox Prompt:="TODO: Write INI File."
    EmailGen.WriteIni INIPath, "Contact Info", "Tel No", WTTelNo
    EmailGen.WriteIni INIPath, "Contact Info", "Email Addr", WTEmail
    EmailGen.WriteIni INIPath, "Contact Hours", "Start Day", WTStartDay
    EmailGen.WriteIni INIPath, "Contact Hours", "End Day", WTEndDay
    EmailGen.WriteIni INIPath, "Contact Hours", "Start Time", WTStartTime
    EmailGen.WriteIni INIPath, "Contact Hours", "End Time", WTEndTime
    EmailGen.WriteIni INIPath, "After Hours Contact", "Tel No", WTAHTelNo
    EmailGen.WriteIni INIPath, "After Hours Contact", "Email Addr", WTAHEmail
    MsgBox Prompt:="Settings saved successfully."
End Sub

Sub LoadINIData()
    Dim DefaultsUsed As Boolean
    DefaultsUsed = False
    
    WTTelNo = EmailGen.ReadIni(INIPath, "Contact Info", "Tel No")
    If WTTelNo = "FAILED" Then
        WTTelNo = "1-866-528-4699"
        DefaultsUsed = True
    End If
    
    WTEmail = EmailGen.ReadIni(INIPath, "Contact Info", "Email Addr")
    If WTEmail = "FAILED" Then
        WTEmail = "walkthroughsupport@teleconferencingcenter.com"
        DefaultsUsed = True
    End If
    
    WTStartDay = EmailGen.ReadIni(INIPath, "Contact Hours", "Start Day")
    If WTStartDay = "FAILED" Then
        WTStartDay = "Monday"
        DefaultsUsed = True
    End If
    
    WTEndDay = EmailGen.ReadIni(INIPath, "Contact Hours", "End Day")
    If WTEndDay = "FAILED" Then
        WTEndDay = "Friday"
        DefaultsUsed = True
    End If
    
    WTStartTime = EmailGen.ReadIni(INIPath, "Contact Hours", "Start Time")
    If WTStartTime = "FAILED" Then
        WTStartTime = "7:30 AM"
        DefaultsUsed = True
    End If
    
    WTEndTime = EmailGen.ReadIni(INIPath, "Contact Hours", "End Time")
    If WTEndTime = "FAILED" Then
        WTEndTime = "4:30 PM"
        DefaultsUsed = True
    End If
    
    WTAHTelNo = EmailGen.ReadIni(INIPath, "After Hours Contact", "Tel No")
    If WTAHTelNo = "FAILED" Then
        WTAHTelNo = "1-866-248-7712"
        DefaultsUsed = True
    End If
    
    WTAHEmail = EmailGen.ReadIni(INIPath, "After Hours Contact", "Email Addr")
    If WTAHEmail = "FAILED" Then
        WTAHEmail = "reservations@teleconferencingcenter.com"
        DefaultsUsed = True
    End If
    
    If DefaultsUsed = True Then
        MsgBox ("Default values were used for the settings, please check the settings, and save them.")
    End If

End Sub

Function HyphenateNumber(ByVal inNumber As String) As String
    Dim strPart1, strPart2, strPart3 As String
    strPart1 = Mid(inNumber, 1, 3)
    strPart2 = Mid(inNumber, 4, 3)
    strPart3 = Mid(inNumber, 7)
    HyphenateNumber = strPart1 & "-" & strPart2 & "-" & strPart3
End Function


