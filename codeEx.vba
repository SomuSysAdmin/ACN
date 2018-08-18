'Code Example
' Please note that the folling is just a snippet of the complete code contained in the file ICG.xls
' And isn't intended to be executed stand-alone. It contains several examples from various modules.

' The code below is to auto-generate the formatting for the raw data from customers and create log entries
' during status updates. The code to auto-generate jQuery code isn't included since it's in development.

' Author -- Somenath Sinha, NetOps, ICG.

' A Screenshot of the excel sheet that this code deals with can be found
' at https://github.com/SomuSysAdmin/ACN/blob/master/interface_su.png

Sub CopyCell(cell As Range)
    Dim objData As New DataObject
    Dim strTemp As String
    strTemp = cell.Value
    objData.SetText (strTemp)
    objData.PutInClipboard
End Sub

Private Sub clearSU_Click()
    Range("C2:D3, G2:G3").Value = ""
    Range("C6:C11, E6:G10").Value = ""
    Range("E11, G11, D12").Value = ""
    Range("su_callLog, su_escLog, su_ocode, su_rcode").Value = ""
End Sub

Private Sub su_callLogCopy_Click()
    CopyCell Range("su_callLog")
End Sub

Private Sub su_escCopy_Click()
    CopyCell Range("su_escLog")
End Sub

Private Sub su_oCodeCopy_Click()
    CopyCell Range("su_ocode")
End Sub

Private Sub su_process_Click()
Dim callLogText As String
Dim fullName As String
Dim pronoun As String
Dim oCodeLabel As String
Dim rCodeLabel As String
oCodeLabel = ""
rCodeLabel = ""

'Variable Value Assignments
If Range("su_gender") = "M" Or Range("su_gender") = "m" Then
    pronoun = "him"
Else
    pronoun = "her"
End If
fullName = Range("su_callerName") & " " & Range("su_callerLastName")

'Making the Call Log
callLogText = fullName & " called in to get a status update." & vbCrLf & _
    "Action Taken:" & vbCrLf & _
    "* Gave the latest status update to " & pronoun

' Status Advise
If Range("su_updated").Value = "Y" Or Range("su_updated").Value = "y" _
  And isNotEmpty(Range("su_advise")) Then
    callLogText = callLogText & _
    " and told " & pronoun & " that " & Range("su_advise") & "."
Else
    callLogText = callLogText & "."
End If

' OPS Notify
If Range("su_opsNotify").Value = "Y" Or Range("su_opsNotify").Value = "y" _
  And isNotEmpty(Range("su_opsNotifyAbout")) Then
    callLogText = callLogText & vbCrLf & _
    "* Notified the tech team that " & Range("su_opsNotifyAbout") & "."
    'Adding to O-Code
    oCodeLabel = fullName & " called in to let the tech team know that " & Range("su_opsNotifyAbout") & "."
End If

' TMG Notify
If Range("su_tmgNotify").Value = "Y" Or Range("su_tmgNotify").Value = "y" _
  And isNotEmpty(Range("su_tmgNotifyAbout")) Then
    callLogText = callLogText & vbCrLf & _
    "* Notified the TMG team that " & Range("su_tmgNotifyAbout") & "."
    'Adding to R-Code
    rCodeLabel = fullName & " called in to let the TMG team know that " & Range("su_tmgNotifyAbout") & "."
End If

' Escalation
If Range("su_esc").Value = "Y" Or Range("su_esc").Value = "y" _
  And isNotEmpty(Range("su_escLvl")) Then
    Dim escLog As String
    escLog = "Upon the request of " & Range("su_callerName") & ", escalated the ticket to level " & Range("su_escLvl") & "."
    callLogText = callLogText & vbCrLf & _
    "* " & escLog

    Dim escLabel As String
    escLabel = ""
    For i = 1 To 20
        escLabel = escLabel & Range("su_escLvl").Value
    Next i
    escLabel = escLabel & "  Escalation  " & escLabel & vbCrLf & vbCrLf
    escLog = escLabel & escLog

    Range("su_escLog").Value = escLog
End If

' Callback
If Range("su_callBack").Value = "T" Or Range("su_callBack").Value = "t" _
  And isNotEmpty(Range("su_callBackNo")) Then
    callLogText = callLogText & vbCrLf & _
    "* Arranged a callback from the Tech team to " & Range("su_callerName") & " on " & Range("su_callBackNo")
    If isNotEmpty(Range("su_opsNotifyAbout")) Then
        oCodeLabel = oCodeLabel & vbCrLf & vbCrLf
    End If
    oCodeLabel = oCodeLabel & _
        "Please callback " & Range("su_callerName") & " on " & Range("su_callBackNo")
    If isNotEmpty(Range("su_reason")) Then
        oCodeLabel = oCodeLabel & " because " & Range("su_reason") & "."
    End If
ElseIf Range("su_callBack").Value = "M" Or Range("su_callBack").Value = "m" _
  And isNotEmpty(Range("su_callBackNo")) Then
    callLogText = callLogText & vbCrLf & _
    "* Arranged a callback from the TMG team to " & Range("su_callerName") & " on " & Range("su_callBackNo")
    If isNotEmpty(Range("su_tmgNotifyAbout")) Then
        rCodeLabel = rCodeLabel & vbCrLf & vbCrLf
    End If
    rCodeLabel = rCodeLabel & _
        "Please callback " & Range("su_callerName") & " on " & Range("su_callBackNo")
    If isNotEmpty(Range("su_reason")) Then
        rCodeLabel = rCodeLabel & " because " & Range("su_reason") & "."
    End If
End If
If isNotEmpty(Range("su_callBackNo")) And isNotEmpty(Range("su_reason")) Then
    callLogText = callLogText & _
        " because " & Range("su_reason") & "."
End If

' Bridge
If Range("su_bridge").Value = "T" Or Range("su_bridge").Value = "t" _
  And isNotEmpty(Range("su_agentId")) Then
    callLogText = callLogText & vbCrLf & _
    "* Bridged the call to the Tech team so that " & Range("su_callerName") & " could talk to " & Range("su_agentId")
ElseIf Range("su_bridge").Value = "M" Or Range("su_bridge").Value = "m" _
  And isNotEmpty(Range("su_agentId")) Then
    callLogText = callLogText & vbCrLf & _
    "* Bridged the call to the TMG team so that " & Range("su_callerName") & " could talk to " & Range("su_agentId")
End If
' Adding Agent Name if available
If isNotEmpty(Range("su_bridge")) Then
    If isNotEmpty(Range("su_agentName")) Then
        callLogText = callLogText & _
            " [" & Range("su_agentName") & "] "
    End If
    ' Adding reason for bridge if available
    If isNotEmpty(Range("su_reason")) Then
        callLogText = callLogText & _
            " because " & Range("su_reason") & "."
    End If
End If

'Dumping contents of R/O Code Labels
If oCodeLabel <> "" Then
    Range("su_ocode").Value = oCodeLabel
End If
If rCodeLabel <> "" Then
    Range("su_rcode").Value = rCodeLabel
End If

'VEC Portal Promotion
callLogText = callLogText & vbCrLf & _
    "* Asked " & Range("su_callerName") & " to visit the VEC portal [https://enterprisecenter.verizon.com] for further updates."

Range("su_callLog").Value = callLogText
End Sub

Private Sub su_rCodeCopy_Click()
    CopyCell Range("su_rcode")
End Sub
