Attribute VB_Name = "PMQu"
Option Explicit
Global Const ver = "1.0.126"
' --------------------------------------------------------
' PMQu
' (c) David R Pratten (2013-2015)

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, Optional ByVal lpParameters As String, Optional ByVal lpDirectory As String, Optional ByVal nShowCmd As Long = 1) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, Optional ByVal lpParameters As String, Optional ByVal lpDirectory As String, Optional ByVal nShowCmd As Long = 1) As Long
#End If
    
Global ActualResultsFieldID As Long
Global ExpectedResultsFieldId As Long
Global ActualAsExpectedFieldID As Long
Global Testing As Boolean
Global LowID As Long
Global HighID As Long
Dim Lookaside As Dictionary
Const htmlCrLf = "<br />"
Const minPerDay = 60 * 8
Const maxTest = 53
Const maxpertestplus1 = 100 ' 4 of on screen message
Dim numOf() As Integer
Dim details() As String
Private Function FieldIDofCustomField(CustomFieldName As String, CustomFieldType As String) As Long
  
  Dim i As Long
  Dim CustomFieldID As Long
  CustomFieldID = 0
  Select Case CustomFieldType
    Case "Number"
        For i = pjTaskNumber1 To pjTaskNumber5
          If CustomFieldGetName(i) = CustomFieldName Then
              CustomFieldID = i
              Exit For
          End If
        Next
        If CustomFieldID = 0 Then
            For i = pjTaskNumber6 To pjTaskNumber20
              If CustomFieldGetName(i) = CustomFieldName Then
                  CustomFieldID = i
                  Exit For
              End If
            Next
        End If
    Case "Text"
        Dim TextFieldIDs(0 To 29) As Long
        TextFieldIDs(0) = 188743731
        TextFieldIDs(1) = 188743734
        TextFieldIDs(2) = 188743737
        TextFieldIDs(3) = 188743740
        TextFieldIDs(4) = 188743743
        TextFieldIDs(5) = 188743746
        TextFieldIDs(6) = 188743747
        TextFieldIDs(7) = 188743748
        TextFieldIDs(8) = 188743749
        TextFieldIDs(9) = 188743750
        TextFieldIDs(10) = 188743997
        TextFieldIDs(11) = 188743998
        TextFieldIDs(12) = 188743999
        TextFieldIDs(13) = 188744000
        TextFieldIDs(14) = 188744001
        TextFieldIDs(15) = 188744002
        TextFieldIDs(16) = 188744003
        TextFieldIDs(17) = 188744004
        TextFieldIDs(18) = 188744005
        TextFieldIDs(19) = 188744006
        TextFieldIDs(20) = 188744007
        TextFieldIDs(21) = 188744008
        TextFieldIDs(22) = 188744009
        TextFieldIDs(23) = 188744010
        TextFieldIDs(24) = 188744011
        TextFieldIDs(25) = 188744012
        TextFieldIDs(26) = 188744013
        TextFieldIDs(27) = 188744014
        TextFieldIDs(28) = 188744015
        TextFieldIDs(29) = 188744016
        For i = 0 To 29
          If CustomFieldGetName(TextFieldIDs(i)) = CustomFieldName Then
              CustomFieldID = TextFieldIDs(i)
              Exit For
          End If
        Next
  End Select
  FieldIDofCustomField = CustomFieldID
  End Function

Public Sub OpenReport(fn As String)
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", fn)
End Sub
Sub ScheduleHealthCheck()
Dim SubPlans As Dictionary
Dim Res As Dictionary
Dim i As Long
Dim tsk As Task
Dim tskID As Variant

' Initialise Global Variables.
ActualResultsFieldID = FieldIDofCustomField("Actual Results", "Text")
ExpectedResultsFieldId = FieldIDofCustomField("Expected Results", "Text")
ActualAsExpectedFieldID = FieldIDofCustomField("Actual As Expected", "Text")
ReDim numOf(maxTest)
ReDim details(maxTest)
Set Lookaside = New Dictionary

If ActualResultsFieldID > 0 And ExpectedResultsFieldId > 0 And ActualAsExpectedFieldID > 0 Then ' this is a test file
    Set SubPlans = SubPlans_set()
    If SubPlans.Count = 0 Then
        MsgBox "This test plan does not contain any SubPlans to run tests on."
    Else
        ' clear out previous test results.
        For Each tsk In ActiveProject.tasks
            tsk.SetField ActualAsExpectedFieldID, ""
            tsk.SetField ActualResultsFieldID, ""
        Next
        For i = 0 To UBound(SubPlans.Keys)
            LowID = Val(SubPlans.Keys(i))
            If i = UBound(SubPlans.Keys) Then
                HighID = ActiveProject.tasks.Count
            Else
                HighID = Val(SubPlans.Keys(i + 1)) - 1
            End If
            Testing = True
            Set Res = CheckAnalyse("All", "PMQu - Project Information Quality Check")
        Next
        ' Compare Actual with Expected updating results
        For Each tsk In ActiveProject.tasks
            If tsk.GetField(ActualResultsFieldID) <> tsk.GetField(ExpectedResultsFieldId) Then
                tsk.SetField ActualAsExpectedFieldID, "FAIL"
            Else
                tsk.SetField ActualAsExpectedFieldID, "OK"
            End If
        Next
    End If
Else
    Testing = False
    LowID = 1
    HighID = ActiveProject.tasks.Count
    Set Res = CheckAnalyse("All", "PMQu - Project Information Quality Check")
End If
Dim chkPathName As String
If Res("Linked to Disk File") Then
    chkPathName = CreateReport("Check", Res("message"))
    OpenReport (chkPathName)
Else
    MsgBox "The project must be first saved to disk."
End If

End Sub

Private Function CreateReport(Suffix As String, message As String) As String
    Dim msgStyle As VbMsgBoxStyle
    Dim FSO As Variant
    Dim oFile As Variant

    Dim chkPathName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim myFilePath As String
    
    ' Make the folder for the reports if necessary.
    chkPathName = CStr(Environ("appdata")) & "\PMQu"
    If Not FileFolderExists(chkPathName) Then
        MkDir chkPathName
    End If
    
    chkPathName = chkPathName & "\" & ActiveProject.Name & " " & Suffix & ".html"
    Set oFile = FSO.CreateTextFile(chkPathName)

    oFile.Write "<html><head><style>body {  font-family: Verdana, Arial, sans-serif; }   div.details {   margin-left: 4em; margin-bottom:1em;} h1,h2,h3,h4 {     color: #234F32;  margin-top:.8em;     font-family:""Trebuchet MS"",sans-serif;     font-weight:normal; } h1 {     font-size:218%;     margin-top:.6em;     margin-bottom:.6em;     line-height:1.1em; } h2 {     font-size:150%;     margin-top:1em;     margin-bottom:.2em;     line-height:1.2em; p, ul, dl {     margin-top:.6em;     margin-bottom:.8em; }</style></head><body>"
    oFile.Write message
    oFile.Write "</body></html>"
    oFile.Close
    CreateReport = chkPathName
End Function
Sub LogErrorConfig(testNo As Integer, Setting As String)
    numOf(testNo) = numOf(testNo) + 1
    details(testNo) = details(testNo) & "    " & Setting & htmlCrLf
    ' There is currently no testing version of this.
End Sub
Sub LogErrorProject(testNo As Integer, message As String)
    ' Default to logging project wide issues against the Level 1 Summary Component.
    LogErrorTask testNo, ActiveProject.tasks(LowID), message
End Sub
Sub ClearActualResults(testNo As Integer)
    Dim tsk As Task
    For Each tsk In ActiveProject.tasks
        If Testing And tsk.ID >= LowID And tsk.ID <= HighID Then
            tsk.SetField ActualResultsFieldID, Replace(tsk.GetField(ActualResultsFieldID), "{" & Trim(Str(testNo)) & "}", "")
        End If
    Next
    numOf(testNo) = 0
    details(testNo) = ""
End Sub
Sub LogErrorTask(testNo As Integer, tsk As Task, message As String, Optional tsk2 As Task = Nothing)
    If Testing Then
        tsk.SetField ActualResultsFieldID, tsk.GetField(ActualResultsFieldID) & "{" & Trim(Str(testNo)) & "}"
        If Not tsk2 Is Nothing Then
            tsk2.SetField ActualResultsFieldID, tsk2.GetField(ActualResultsFieldID) & "{" & Trim(Str(testNo)) & "}"
        End If
    End If
    numOf(testNo) = numOf(testNo) + 1
    If message <> "" Then
        Dim buildMsg As String
        buildMsg = message
        buildMsg = Replace(buildMsg, "!numOf(testNo)!", Str(numOf(testNo)))
        buildMsg = Replace(buildMsg, "!NameID!", tsk.Name & "[" & Trim(Str(tsk.ID)) & "]")
        If Not tsk2 Is Nothing Then
            buildMsg = Replace(buildMsg, "!NameID2!", tsk2.Name & "[" & Trim(Str(tsk2.ID)) & "]")
        End If
        If numOf(testNo) < maxpertestplus1 Then details(testNo) = details(testNo) & buildMsg & htmlCrLf
    End If
End Sub
Private Function CheckAnalyse(IncludedTests As String, ReportName As String) As Dictionary
' IncludedTests "All" = all, otherwise a comma separated list of tests to perform
    Application.StatusBar = "Schedule Health Check ..."
    Dim tsk As Task
    Dim tsk2 As Task
    Dim message As String
    Dim preamble As String
    Dim descOf(maxTest) As String
    Dim sevOf(maxTest) As Integer
    Dim bandOf(maxTest) As Integer
    
    Dim IncludedOf(maxTest) As Boolean
    Dim maxSev As Integer
    Const sevWarning = 1
    Const sevError = 3
    Dim testNo As Integer
    Dim i As Integer
    Dim j As Integer
    Dim TotalFound As Integer
    Dim startFound As Boolean
    Dim startCount As Integer
    Dim finishFound As Boolean
    Dim finishCount As Integer
    Dim TaskID As Variant
    Dim nonMilestoneChildren As Integer
    Dim thisSuccessor As Dictionary
    Dim successorsLessOne As Dictionary
    Dim severitymessage As String
    Dim reportable As Boolean
    Dim Res As New Dictionary
    Dim StartMilestoneID As New Dictionary
    Dim FinishMilestoneID As New Dictionary
    Dim PermanentIDFieldID As Long
    Dim tskParentsID As Dictionary
    Dim tsk2ParentsID As Dictionary
    Dim thistsk As Dictionary
    Dim thistsk2 As Dictionary
    Dim cola As Dictionary
    Dim MaxPermID As Integer
    Dim HealthCheckOptionsID As Long
    Dim settings36(16) As String
    Dim settings47(2) As String
    Dim StartMilestones As New Dictionary
    Dim FinishMilestones As New Dictionary
    Dim Scratch As String
    
    message = ""
    'details = ""
    Dim healthcheckoptionsFieldName As String
    healthcheckoptionsFieldName = "PMQu Check Exclusions"
    HealthCheckOptionsID = FieldIDofCustomField(healthcheckoptionsFieldName, "Text")
    
    
    
    For i = 1 To maxTest
        sevOf(i) = sevError
        IncludedOf(i) = True
        If IncludedTests <> "All" Then
            If UboundFilterExactMatch(IncludedTests, i) < 0 Then
                IncludedOf(i) = False
            End If
        End If
        If HealthCheckOptionsID <> 0 Then
            If UboundFilterExactMatch(ActiveProject.ProjectSummaryTask.GetField(HealthCheckOptionsID), i) >= 0 Then
                IncludedOf(i) = False
            End If
        End If
    Next
    
  

    
    descOf(1) = "Summary with Resource Assignments" ' Item.
    bandOf(1) = 50
    descOf(2) = "Summary with Successors" ' Item.
    bandOf(2) = 40
    descOf(3) = "Summary with Predecessors" ' Item.
    bandOf(3) = 40
    descOf(4) = "Task with elapsed time > 30 days" ' Item.
    sevOf(4) = sevWarning
    bandOf(4) = 60
    'descOf(5) = "Milestone with constraint type other than ASAP, MSO, SNET, or FNLT" ' Item.  ' Harris 2010 c11
    'sevOf(5) = sevWarning
    bandOf(5) = 60
    descOf(6) = "Task with constraint type other than ASAP. Use Deadlines or put constraint on a milestone" ' Item.
    bandOf(6) = 60
    descOf(7) = "Summary with constraint type other than ASAP" ' Item.
    bandOf(7) = 60
    descOf(8) = "Manually scheduled" ' Item.
    bandOf(8) = 60
    descOf(9) = "Task Type other than Fixed Units" ' Item.
    bandOf(9) = 50
    descOf(10) = "Tasks/Milestones without predecessor" ' Network.  ' (excl. External and SNET Milestones)"
    bandOf(10) = 40
    descOf(11) = "Tasks/Milestones without successor" ' Network.  ' (excl. External and FNLT Milestones)"
    bandOf(11) = 40
    descOf(12) = "Tasks without duration specified" ' Item.
    bandOf(12) = 60
    descOf(13) = "Summary without Start or Finish milestones" ' WBS/PBS.
    bandOf(13) = 30
    descOf(14) = "Milestone with Resource Assignments" ' Item.
    bandOf(14) = 50
    'descOf(15) = "Start Milestone with no sibling successors" ' Network.
    'descOf(16) = "Start Milestone with no non-sibling predecessors" ' Network.
    'descOf(17) = "Finish Milestone with no sibling predecessors" ' Network.
    'descOf(18) = "Finish Milestone with no non-sibling successors" ' Network.
    descOf(19) = "Tasks with Duplicate Names" ' Item.
    bandOf(19) = 20
    descOf(20) = "Tasks with sub-day duration" ' Item.
    bandOf(20) = 60
    descOf(21) = "Tasks with Start date before Status Date with no Actual Start date" ' Progress.
    bandOf(21) = 70
    sevOf(21) = sevWarning
    descOf(22) = "Tasks with Finish date before Status Date with no Actual Finish date" ' Progress.
    bandOf(22) = 70
    sevOf(22) = sevWarning
    descOf(23) = "Start Milestone is not the predecessor of its siblings" ' Network.
    bandOf(23) = 30
    descOf(24) = "Finish Milestone is not the successor of its siblings" ' Network.
    bandOf(24) = 30
    descOf(25) = "Actual Start is after Status Date" ' Progress.
    bandOf(25) = 70
    descOf(26) = "Actual Finish is after Status Date" ' Progress.
    bandOf(26) = 70
    descOf(27) = "Summary has fewer than two children" ' WBS/PBS.
    bandOf(27) = 30
    descOf(28) = "Unmet constraint generating negative slack" ' Scheduling.
    bandOf(28) = 60
    sevOf(28) = sevWarning
    descOf(29) = "Task has more than 30 days slack" ' Scheduling.
    bandOf(29) = 60
    sevOf(29) = sevWarning
    descOf(30) = "Milestone %Complete must be 0% or 100%" ' Item.
    bandOf(30) = 70
    descOf(31) = "Dependency is redundant" ' Network.
    bandOf(31) = 40
    descOf(32) = "Project has more than one top level task (You may add a 'Status Date Milestone' at Outline Level 1 with Start No Earlier Than constraint)" ' WBS/PBS.
    bandOf(32) = 30
    descOf(33) = "Project Summary Task is visible" ' WBS/PBS.
    bandOf(33) = 10
'    descOf(34) = "Project does not contain a 'Status Date Milestone' at Outline Level 1 with Start No Earlier Than constraint" ' Item.
'    sevOf(34) = sevWarning
    descOf(35) = "A successor of Status Date Milestone has less than 10 days slack" ' Scheduling.
    sevOf(35) = sevWarning
    bandOf(35) = 60
    descOf(36) = "Recommend the following settings under File > Options > Schedule" ' Project.
    bandOf(36) = 10
    sevOf(36) = sevWarning
    descOf(37) = "You are not using MS Project 2010 or 2013.  Your mileage may vary." ' Project.
    bandOf(37) = 10
    sevOf(37) = sevWarning
    descOf(38) = "'Status Date Milestone' has a predecessor."
    bandOf(38) = 40
    descOf(39) = "'Status Date Milestone' has a successor that is not floating, it has already started."
    bandOf(39) = 40
    descOf(40) = "Task has blank name" ' Item.
    bandOf(40) = 20
    descOf(41) = "Only this Summary Task's Start Milestone's name may begin with 'Start '" ' WBS/PBS.
    bandOf(41) = 30
    descOf(42) = "Only this Summary Task's Finish Milestone's name may begin with 'Finish '" ' WBS/PBS.
    bandOf(42) = 30
    descOf(43) = "Dependency may not be with self, not with parent Tasks, nor with child Tasks" ' Network.
    bandOf(43) = 40
    descOf(44) = "Task with duplicate 'Permanent ID' field" ' Item.
    bandOf(44) = 20
    'descOf(45) = "Use an Interim Output Milestone here and make distant tasks dependent on the Milestone." ' Network.
    'bandOf(45) = 40
    'sevOf(45) = sevWarning
    descOf(46) = "Tasks with common predecessors suggests that an Interim Milestone is missing" ' Network.
    bandOf(46) = 40
    sevOf(46) = sevWarning
    descOf(47) = "Recommend the following settings under Project > Project Information" ' Project Information.
    sevOf(47) = sevWarning
    bandOf(47) = 10
    descOf(48) = "Use an Interim Milestone as the predecessor or successor." ' Network.
    sevOf(48) = sevWarning
    bandOf(48) = 40
    descOf(49) = "Permanent ID field must be 1 or greater" ' Item.
    bandOf(49) = 20
    descOf(50) = "Task is Effort Driven." ' Item.
    bandOf(50) = 50
    descOf(51) = "Start Milestone has a successor which is not a sibling"
    bandOf(51) = 40
    descOf(52) = "Finish Milestone has a predecessor which is not a sibling"
    bandOf(52) = 40
    descOf(53) = "Milestone must be zero Duration"
    bandOf(53) = 20
    
    
    Res.Add "Linked to Disk File", (UBound(Split(ActiveProject.FullName, ".")) > 0)
    
    
    PermanentIDFieldID = FieldIDofCustomField("Permanent ID", "Number")
    
    
    ' Version check
    testNo = 37
    If Application.Version <> "14.0" And Application.Version <> "15.0" And IncludedOf(testNo) Then
        LogErrorProject testNo, ""
    End If
    
    settings36(1) = "New tasks created: [Auto Scheduled]"
    settings36(2) = "Auto scheduled tasks scheduled on: [Project Start Date]"
    settings36(3) = "Duration is entered in: [Days]"
    settings36(4) = "Work is entered in: [Hours]"
    settings36(5) = "Default task type: [Fixed Units]"
    settings36(6) = "<b>&#x2610;</b> New tasks are effort driven"
    settings36(7) = "<b>&#x2610;</b> Autolink inserted or moved tasks"
    settings36(8) = "<b>&#x2610;</b> Split in-progress tasks"
    settings36(9) = "<b>&#x2610;</b> Update Manually Scheduled tasks when editing links"
    settings36(10) = "<b>&#x2610;</b> Tasks will always honor their constraint dates"
    settings36(11) = "<b>&#x2611;</b> Show that scheduled tasks have estimated durations"
    settings36(12) = "<b>&#x2611;</b> New scheduled tasks have estimated durations"
    settings36(13) = "<b>&#x2610;</b> Keep task on nearest working day when changing to Automatically Scheduled mode"
    settings36(14) = "<b>&#x2610;</b> Show task schedule warnings"
    settings36(15) = "<b>&#x2610;</b> Show task schedule suggestions"

    
    ' Check schedule options
    testNo = 36
    If IncludedOf(testNo) Then
        If ActiveProject.NewTasksCreatedAsManual Then
            LogErrorConfig testNo, settings36(1)
        End If
        If Not ActiveProject.ScheduleFromStart Then
            LogErrorConfig testNo, settings36(2)
        End If
        If Not ActiveProject.DefaultDurationUnits = pjDay Then
            LogErrorConfig testNo, settings36(3)
        End If
        If Not ActiveProject.DefaultWorkUnits = pjHour Then
            LogErrorConfig testNo, settings36(4)
        End If
        If Not ActiveProject.DefaultTaskType = pjFixedUnits Then
            LogErrorConfig testNo, settings36(5)
        End If
        If ActiveProject.DefaultEffortDriven Then
            LogErrorConfig testNo, settings36(6)
        End If
        If ActiveProject.AutoLinkTasks Then
            LogErrorConfig testNo, settings36(7)
        End If
        If ActiveProject.AutoSplitTasks Then
            LogErrorConfig testNo, settings36(8)
        End If
        If ActiveProject.ManuallyScheduledTasksAutoRespectLinks Then
            LogErrorConfig testNo, settings36(9)
        End If
        If ActiveProject.HonorConstraints Then
            LogErrorConfig testNo, settings36(10)
        End If
        If Not ActiveProject.ShowEstimatedDuration Then
            LogErrorConfig testNo, settings36(11)
        End If
        If Not ActiveProject.NewTasksEstimated Then
            LogErrorConfig testNo, settings36(12)
        End If
        If ActiveProject.KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled Then
            LogErrorConfig testNo, settings36(13)
        End If
        If ActiveProject.ShowTaskWarnings Then
            LogErrorConfig testNo, settings36(14)
        End If
        If ActiveProject.ShowTaskSuggestions Then
            LogErrorConfig testNo, settings36(15)
        End If
        
        
    End If
    
    settings47(1) = "Schedule From [Project Start Date]"
    
    testNo = 47
    If IncludedOf(testNo) Then
        If Not ActiveProject.ScheduleFromStart Then
            LogErrorConfig testNo, settings47(1)
        End If
        
    End If
    'Prior to analysis move the "Status Date Milestone to the ReallyStatusDate().
    
    For Each tsk In ActiveProject.tasks
        testNo = 40
        If tsk Is Nothing Then
            LogErrorProject testNo, "Blank task #!numOf(testNo)! found."
        ElseIf Trim(tsk.Name) = "" And tsk.ID >= LowID And tsk.ID <= HighID Then
            LogErrorTask testNo, tsk, "Blank task #!numOf(testNo)! found."
        End If
    Next
    
    If numOf(testNo) = 0 Then ' only do more tests if all tasks are not nothing and not blank ie have a non-blank name
    
        testNo = 34
        Dim StatusDateMilestoneID As Integer
        StatusDateMilestoneID = -1
        For Each tsk In ActiveProject.tasks
            If tsk.ID >= LowID And tsk.ID <= HighID Then
                If tsk.Name = "Status Date Milestone" And tsk.OutlineLevel = 1 And tsk.Milestone And tsk.ConstraintType = pjSNET Then
                    StatusDateMilestoneID = tsk.ID
                    tsk.ConstraintDate = ReallyStatusDate()
                End If
            End If
        Next
        If StatusDateMilestoneID = -1 Then
            ' pass
        Else
            ' check to see if any of the successors have low slack
            
            For Each tsk2 In ActiveProject.tasks(StatusDateMilestoneID).SuccessorTasks
                If tsk2.ID >= LowID And tsk2.ID <= HighID Then
                    ' check if successors have already started
                    testNo = 39
                    If tsk2.PercentComplete <> 0 Then
                        If IncludedOf(testNo) Then
                            LogErrorTask testNo, ActiveProject.tasks(StatusDateMilestoneID), "Remove predecessor !NameID! from !NameID2! it started on " & tsk2.ActualStart, tsk2
                        End If
                    Else ' only test if 39 passes
                        testNo = 35
                        If Min(tsk2.TotalSlack, tsk2.StartSlack) < 10 * minPerDay And IncludedOf(testNo) Then
                            LogErrorTask testNo, tsk2, "!NameID! has " & Min(tsk2.TotalSlack, tsk2.StartSlack) / minPerDay & " days slack"
                        End If
                    End If
                End If
            Next
            ' check if status date milestone has predecessors
            testNo = 38
            If ActiveProject.tasks(StatusDateMilestoneID).PredecessorTasks.Count > 0 Then
                LogErrorProject testNo, ""
            End If
        End If
    
        
       
        For Each tsk In ActiveProject.tasks
            If tsk.ID >= LowID And tsk.ID <= HighID Then
                If tsk.Summary Then
                    If tsk.Assignments.Count > 0 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 1) < 0 And IncludedOf(1) Then
                        LogErrorTask 1, tsk, "!NameID!"
                    End If
                End If
                
                If tsk.Summary Then
                    If tsk.SuccessorTasks.Count > 0 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 2) < 0 And IncludedOf(2) Then
                        LogErrorTask 2, tsk, "!NameID!"
                    End If
                End If
                
                If tsk.Summary Then
                    If tsk.PredecessorTasks.Count > 0 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 3) < 0 And IncludedOf(3) Then
                        LogErrorTask 3, tsk, "!NameID!"
                    End If
                End If
                
                If Not tsk.Summary Then
                    If (tsk.Finish - tsk.Start) > 30 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 4) < 0 And IncludedOf(4) Then
                        LogErrorTask 4, tsk, "!NameID!"
                    End If
                End If
                
                If Not tsk.Milestone And Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 6) < 0 And IncludedOf(6) Then
                    If Not (tsk.ConstraintType = pjASAP) Then
                        LogErrorTask 6, tsk, "!NameID!"
                    End If
                End If
                
                If tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 7) < 0 And IncludedOf(7) Then
                    If Not (tsk.ConstraintType = pjASAP) Then
                        LogErrorTask 7, tsk, "!NameID!"
                    End If
                End If
                
                If tsk.Manual And tskFieldExactMatch(tsk, HealthCheckOptionsID, 8) < 0 And IncludedOf(8) Then
                    LogErrorTask 8, tsk, "!NameID!"
                End If
                
                If Not tsk.Milestone And Not tsk.Summary And tsk.Type <> pjFixedUnits And tskFieldExactMatch(tsk, HealthCheckOptionsID, 9) < 0 And IncludedOf(9) Then
                    LogErrorTask 9, tsk, "!NameID!"
                End If
                
                If Not tsk.Milestone And Not tsk.Summary And tsk.EffortDriven And tskFieldExactMatch(tsk, HealthCheckOptionsID, 50) < 0 And IncludedOf(50) Then
                    LogErrorTask 50, tsk, "!NameID!"
                End If
                
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 10) < 0 And IncludedOf(10) Then
                    If tsk.PredecessorTasks.Count = 0 And Not ((Left(tsk.Name, 8) = "External" Or tsk.ConstraintType = pjSNET) And tsk.Milestone) And Not (tsk.OutlineLevel = 2 And Left(tsk.Name, 5) = "Start") Then 'ignore external milestones AND the origin "Start" Milestone
                        LogErrorTask 10, tsk, "!NameID!"
                    End If
                End If
            
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 11) < 0 And IncludedOf(11) Then
                    If tsk.SuccessorTasks.Count = 0 And Not ((InStr(tsk.Name, "External") <> 0 Or tsk.ConstraintType = pjFNLT Or tsk.ID = StatusDateMilestoneID) And tsk.Milestone) And Not (tsk.OutlineLevel = 2 And Left(tsk.Name, 6) = "Finish") Then
                        LogErrorTask 11, tsk, "!NameID!"
                    End If
                End If
                
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 12) < 0 And IncludedOf(12) Then
                    If InStr(tsk.DurationText, "?") > 0 Then
                        LogErrorTask 12, tsk, "!NameID!"
                    End If
                End If
        
                Dim chld As Task
                If tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 13) < 0 And IncludedOf(13) Then
                    startFound = False
                    finishFound = False
                    startCount = 0
                    finishCount = 0
                    For Each chld In tsk.OutlineChildren
                        If chld.Milestone And chld.Name = "Start " & tsk.Name Then
                            startFound = True
                            StartMilestoneID.Add tsk.ID, chld.ID
                        End If
                        If chld.Milestone And chld.Name = "Finish " & tsk.Name Then
                            FinishMilestoneID.Add tsk.ID, chld.ID
                            finishFound = True
                        End If
                        If Left(chld.Name, 6) = "Start " Then
                            startCount = startCount + 1
                        End If
                        If Left(chld.Name, 7) = "Finish " Then
                            finishCount = finishCount + 1
                        End If
                    Next
                    If Not startFound Or Not finishFound Then
                        Scratch = "!NameID! doesn't have "
                        If Not startFound Then
                            Scratch = Scratch & "Start "
                        End If
                        If Not finishFound Then
                            Scratch = Scratch & "Finish "
                        End If
                        Scratch = Scratch & "Start " & "Milestone(s)"
                        LogErrorTask 13, tsk, Scratch
                    End If
                    If numOf(13) = 0 And startCount > 1 Then
                        LogErrorTask 41, tsk, "Summary Task has multiple children with names beginning with 'Start ' - !NameID!"
                    End If
                    If numOf(13) = 0 And finishCount > 1 Then
                        LogErrorTask 42, tsk, "Summary Task has multiple children with names beginning with 'Finish ' - !NameID!"
                    End If
                End If
                
                If tsk.Milestone And tskFieldExactMatch(tsk, HealthCheckOptionsID, 14) < 0 And IncludedOf(14) Then
                    If tsk.Assignments.Count > 0 Then
                        LogErrorTask 14, tsk, "!NameID!"
                    End If
                End If
                For Each tsk2 In ActiveProject.tasks
                    If tsk2.ID >= LowID And tsk2.ID <= HighID Then
                        If tsk2.Name = tsk.Name And tsk2.ID <> tsk.ID And tskFieldExactMatch(tsk, HealthCheckOptionsID, 19) < 0 And IncludedOf(19) Then
                            LogErrorTask 19, tsk, "!NameID! and !NameID2! have duplicate names.", tsk2
                        End If
                    End If
                Next
                
                If PermanentIDFieldID <> 0 Then
                    If tsk.GetField(PermanentIDFieldID) = 0 Then
                        LogErrorTask 49, tsk, "!NameID! has a Permanent ID less than 1."
                    Else
                        
                        For Each tsk2 In ActiveProject.tasks
                            If tsk2.ID >= LowID And tsk2.ID <= HighID Then
                                If tsk2.GetField(PermanentIDFieldID) = tsk.GetField(PermanentIDFieldID) And tsk2.ID > tsk.ID And tskFieldExactMatch(tsk, HealthCheckOptionsID, 44) < 0 And IncludedOf(44) Then
                                    LogErrorTask 44, tsk, tsk.GetField(PermanentIDFieldID) & " is a duplicated 'Permanent ID' between !NameID! and !NameID2!", tsk2
                                End If
                            End If
                        Next
                        MaxPermID = Max(Int(tsk.GetField(PermanentIDFieldID)), MaxPermID)
                    End If
                End If
                
                
            
                If Not tsk.Summary And Not tsk.Milestone And tskFieldExactMatch(tsk, HealthCheckOptionsID, 20) < 0 And IncludedOf(20) Then
                    If tsk.Duration < 8 * 60 Then
                        LogErrorTask 20, tsk, "!NameID!"
                    End If
                End If
            
        
                testNo = 27
                If tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(27) Then
                    If tsk.OutlineChildren.Count < 4 Then
                        LogErrorTask testNo, tsk, "!NameID!"
                    End If
                End If
            'MSO, SNET, or FNLT
                testNo = 28
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(testNo) Then
                    If Min(tsk.StartSlack, Min(tsk.FinishSlack, tsk.TotalSlack)) < 0 And ((tsk.ConstraintType = pjMSO And tsk.Start <> tsk.ConstraintDate) Or (tsk.ConstraintType = pjSNET And tsk.Start < tsk.ConstraintDate) Or (tsk.ConstraintType = pjFNLT And tsk.Finish > tsk.ConstraintDate)) Then
                        LogErrorTask testNo, tsk, "!NameID! has " & Min(tsk.TotalSlack, tsk.StartSlack) / minPerDay & " days slack"
                    End If
                End If
            
                testNo = 29
                If Not tsk.Summary And Not tsk.Milestone And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(testNo) Then
                    If Max(tsk.StartSlack, Max(tsk.FinishSlack, tsk.TotalSlack)) > 30 * 8 * 60 Then
                        LogErrorTask testNo, tsk, "!NameID!"
                    End If
                End If
            
                testNo = 30
                If tsk.Milestone And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(testNo) Then
                    If tsk.PercentComplete <> 0 And tsk.PercentComplete <> 100 Then
                        LogErrorTask testNo, tsk, "!NameID!"
                    End If
                End If
           
                testNo = 53
                If tsk.Milestone And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(testNo) Then
                    If tsk.Duration <> 0 Then
                        LogErrorTask testNo, tsk, "!NameID!"
                    End If
                End If
           
                
                
                ' check for common predecessors
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 46) < 0 And IncludedOf(46) Then
                    Set thistsk = New Dictionary
                    thistsk.Add Str(tsk.ID), Str(tsk.ID)
                    For Each tsk2 In ActiveProject.tasks
                        If tsk.ID < tsk2.ID Then
                            Set thistsk2 = New Dictionary
                            thistsk2.Add Str(tsk2.ID), Str(tsk2.ID)
                            Set cola = Intersect(successors_set(thistsk), thistsk2)
                            If cola.Count() > 0 Then ' successors can't have interesting sets of common precessors
                                GoTo continue2332:
                            End If
                            Set cola = Intersect(successors_set(thistsk2), thistsk)
                            If cola.Count() > 0 Then ' successors can't have interesting sets of common precessors
                                GoTo continue2332:
                            End If
                            Set cola = Intersect(tasks_set(tsk.PredecessorTasks), tasks_set(tsk2.PredecessorTasks))
                            If cola.Count() > 1 Then
                                LogErrorTask 46, tsk, "!NameID! and !NameID2! have the " & cola.Count() & " predecessors [" & Trim(Join(cola.Keys(), ", ")) & "] in common.", tsk2
                            End If
                        End If
continue2332:
                    Next
                End If
                
                testNo = 32
                If tsk.OutlineLevel = 1 And tsk.ID <> StatusDateMilestoneID And IncludedOf(testNo) Then
                    LogErrorTask testNo, tsk, "!NameID!"
                End If
            End If
        Next tsk
        
        
    Set StartMilestones = InvertDictionary(StartMilestoneID)
    Set FinishMilestones = InvertDictionary(FinishMilestoneID)
     
    For Each tsk In ActiveProject.tasks
        If tsk.ID >= LowID And tsk.ID <= HighID Then
            If tsk.Summary Then
                For Each chld In tsk.OutlineChildren
                    If chld.Milestone And chld.Name = "Start " & tsk.Name Then
                        ' Check 15
                        'Set cola = Intersect(descendents_set(tasks_set(chld.SuccessorTasks)), descendents_set(tasks_set(tsk.OutlineChildren)))
                        'If cola.Count() = 0 And tskFieldExactMatch(chld, HealthCheckOptionsID, 15) < 0 And IncludedOf(15) Then
                        '   numOf(15) = numOf(15) + 1
                        '    If numOf(15) < maxpertestplus1 Then details(15) = details(15) & "    " & chld.name & "[" & chld.ID & "] has no peer successors" & htmlCrLf
                        'End If
                        ' Check 16
                        'Set cola = Subtract(descendents_set(tasks_set(chld.PredecessorTasks)), descendents_set(tasks_set(tsk.OutlineChildren)))
                        'If cola.Count() = 0 And tskFieldExactMatch(chld, HealthCheckOptionsID, 16) < 0 And IncludedOf(16) Then
                        '    numOf(16) = numOf(16) + 1
                        '    If numOf(16) < maxpertestplus1 Then details(16) = details(16) & "    " & chld.name & "[" & chld.ID & "] has no external predecessors" & htmlCrLf
                        'End If
                        ' Check 51
                        Set cola = Subtract(tasks_set(chld.SuccessorTasks), tasks_set(chld.OutlineParent.OutlineChildren))
                        If cola.Count() > 0 And tskFieldExactMatch(chld, HealthCheckOptionsID, 51) < 0 And IncludedOf(51) Then
                            For Each TaskID In cola.Keys
                                ' if not start item of a sibling then
                                If Not (Left(ActiveProject.tasks(Val(TaskID)).Name, 6) = "Start " And ActiveProject.tasks(Val(TaskID)).Milestone And ActiveProject.tasks(Val(TaskID)).OutlineParent.OutlineParent.ID = chld.OutlineParent.ID) Then
                                   LogErrorTask 51, chld, "!NameID! should not be the succeeded by !NameID2!.", ActiveProject.tasks(Val(TaskID))
                                End If
                            Next
                        End If
                        If numOf(10) = 0 And numOf(11) = 0 Then
                            Set cola = Subtract(descendents_set_except(tasks_set(tsk.OutlineChildren), Str(chld.ID), False), successors_set(tasks_set(chld.SuccessorTasks)))
                            If cola.Count() > 0 And tskFieldExactMatch(chld, HealthCheckOptionsID, 23) < 0 And IncludedOf(23) Then
                                LogErrorTask 23, tsk, ""
                                For Each TaskID In cola.Keys
                                    LogErrorTask 23, ActiveProject.tasks(Val(TaskID)), chld.Name & "[" & chld.ID & "] is not the predecessor of !NameID!"
                                Next
                            End If
                        End If
                    End If
                    If chld.Milestone And chld.Name = "Finish " & tsk.Name Then
                        
                        ' Check 52
                        Set cola = Subtract(tasks_set(chld.PredecessorTasks), tasks_set(chld.OutlineParent.OutlineChildren))
                        If cola.Count() > 0 And tskFieldExactMatch(chld, HealthCheckOptionsID, 52) < 0 And IncludedOf(52) Then
                            For Each TaskID In cola.Keys
                                ' if not start item of a sibling then
                                If Not (Left(ActiveProject.tasks(Val(TaskID)).Name, 7) = "Finish " And ActiveProject.tasks(Val(TaskID)).Milestone And ActiveProject.tasks(Val(TaskID)).OutlineParent.OutlineParent.ID = chld.OutlineParent.ID) Then
                                    LogErrorTask 52, chld, "!NameID! should not be the preceded by !NameID2!.", ActiveProject.tasks(Val(TaskID))
                                End If
                            Next
                        End If
                        If numOf(10) = 0 And numOf(11) = 0 And numOf(23) = 0 And IncludedOf(23) And IncludedOf(24) Then    'only do this test if the 23's are clear
                            Set cola = Subtract(descendents_set_except(tasks_set(tsk.OutlineChildren), Str(chld.ID), False), predecessors_set(tasks_set(chld.PredecessorTasks)))
                            If cola.Count() > 0 Then
                                
                                reportable = False
                                For Each TaskID In cola.Keys
                                    If numOf(24) < maxpertestplus1 And Val(TaskID) <> StatusDateMilestoneID Then
                                            reportable = True
                                            details(24) = details(24) & "    " & chld.Name & "[" & chld.ID & "] is not the successor of " & ActiveProject.tasks(Val(TaskID)).Name & "[" & ActiveProject.tasks(Val(TaskID)).ID & "]" & htmlCrLf
                                    End If
                                Next
                                If reportable Then
                                    LogErrorProject 24, ""
                                End If
                            End If
                        End If
                        
                   End If
                Next
            End If
        
            If numOf(2) = 0 And numOf(3) = 0 And numOf(10) = 0 And numOf(11) = 0 And numOf(31) = 0 And numOf(13) = 0 And numOf(41) = 0 And numOf(42) = 0 And numOf(51) = 0 And numOf(52) = 0 Then   ' Only perform this test if otherwise all OK.
        
                ' assumes that no dependencies on summaries
                If Not tsk.Summary And tsk.SuccessorTasks.Count > 0 Then
                    For Each tsk2 In tsk.SuccessorTasks
                        Set thisSuccessor = New Dictionary
                        thisSuccessor.Add Str(tsk2.ID), Str(tsk2.ID)
                        Set successorsLessOne = Subtract(tasks_set(tsk.SuccessorTasks), thisSuccessor)
                        Set cola = Intersect(successors_set(successorsLessOne), thisSuccessor)
                        If cola.Count() <> 0 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 31) < 0 And IncludedOf(31) Then
                            LogErrorTask 31, tsk, "!NameID! has redundant successor dependency to !NameID2!.", tsk2
                        End If
                        
                    Next
                End If
               testNo = 51
               If IncludedOf(testNo) And StartMilestones.Exists(tsk.ID) Then
                   
                   Set cola = Subtract(tasks_set(tsk.SuccessorTasks), start_successor_set(tsk, StartMilestoneID, FinishMilestones))
                   If cola.Count() > 0 Then
                       For Each TaskID In cola.Keys
                           LogErrorTask testNo, tsk, "!NameID! has invalid successor dependency to !NameID2!.", ActiveProject.tasks(Val(TaskID))
                       Next
                   End If
               End If
    
               testNo = 52
               If IncludedOf(testNo) And FinishMilestones.Exists(tsk.ID) Then
                   Set cola = Subtract(tasks_set(tsk.PredecessorTasks), finish_predecessor_set(tsk, FinishMilestoneID, StartMilestones))
                   If cola.Count() > 0 Then
                       For Each TaskID In cola.Keys
                            LogErrorTask testNo, tsk, "!NameID! has invalid predecessor dependency from !NameID2!.", ActiveProject.tasks(Val(TaskID))
                       Next
                   End If
               End If
       
            End If
        End If
    Next tsk
        
    testNo = 33
    If ActiveProject.DisplayProjectSummaryTask And IncludedOf(testNo) Then
        LogErrorProject testNo, ""
    End If
    
    
    If numOf(2) = 0 And numOf(3) = 0 And numOf(10) = 0 And numOf(11) = 0 And numOf(31) = 0 And numOf(13) = 0 And numOf(41) = 0 And numOf(42) = 0 Then   ' Only perform this test if otherwise all OK.
        
    
        testNo = 43
        
        If IncludedOf(testNo) Then
        
            Dim InError As Boolean
            Dim Taskto As Task
            Dim Taskfrom As Task
            Dim WBSfrom As Integer
            Dim WBSto As Integer
            Dim DepList As New Dictionary
            Dim DepCount As Integer
            Dim DepKey As Variant
            Dim FromToStartFinish As Boolean
            Const iWBSfrom = 0
            Const iWBSto = 1
            Const iTaskfromID = 2
            Const iTasktoID = 3
            DepCount = 0
            For Each Taskfrom In ActiveProject.tasks
                If Taskfrom.ID >= LowID And Taskfrom.ID <= HighID Then
                    For Each Taskto In Taskfrom.SuccessorTasks
                        ' translate Start / Finish Milestones to be represented by their Summary Task
                        FromToStartFinish = False
                        If StartMilestones.Exists(Taskfrom.ID) Then
                            WBSfrom = StartMilestones(Taskfrom.ID)
                            FromToStartFinish = True
                        ElseIf FinishMilestones.Exists(Taskfrom.ID) Then
                            WBSfrom = FinishMilestones(Taskfrom.ID)
                            FromToStartFinish = True
                        Else
                            WBSfrom = Taskfrom.ID
                        End If
                        If StartMilestones.Exists(Taskto.ID) Then
                            WBSto = StartMilestones(Taskto.ID)
                            FromToStartFinish = True
                        ElseIf FinishMilestones.Exists(Taskto.ID) Then
                            WBSto = FinishMilestones(Taskto.ID)
                            FromToStartFinish = True
                        Else
                            WBSto = Taskto.ID
                        End If
                        If FromToStartFinish Then
                            DepCount = DepCount + 1
                            DepList.Add DepCount, Array(WBSfrom, WBSto, Taskfrom.ID, Taskto.ID)
                        End If
                    Next
                End If
            Next
            For Each DepKey In DepList
                InError = False
                If DepList(DepKey)(iWBSfrom) = DepList(DepKey)(iWBSto) Then ' can't depend on itself
                    InError = True
                ElseIf wbs_descendents_set(ActiveProject.tasks(DepList(DepKey)(iWBSfrom))).Exists(Str(DepList(DepKey)(iWBSto))) Then 'if wbsto is a descendent of wbsfrom
                    InError = InError Or StartMilestoneID(DepList(DepKey)(iWBSfrom)) <> DepList(DepKey)(iTaskfromID) 'allowable if from  start milestone
                    If ActiveProject.tasks(DepList(DepKey)(iWBSto)).Summary Then
                        InError = InError Or StartMilestoneID(DepList(DepKey)(iWBSto)) <> DepList(DepKey)(iTasktoID) 'allowable if to  start milestone
                        InError = InError Or ActiveProject.tasks(DepList(DepKey)(iTaskfromID)).OutlineLevel + 1 <> ActiveProject.tasks(DepList(DepKey)(iTasktoID)).OutlineLevel ' immediate = one level different
                    Else
                        InError = InError Or ActiveProject.tasks(DepList(DepKey)(iTaskfromID)).OutlineLevel <> ActiveProject.tasks(DepList(DepKey)(iTasktoID)).OutlineLevel 'immediate means same level
                    End If
                ElseIf wbs_descendents_set(ActiveProject.tasks(DepList(DepKey)(iWBSto))).Exists(Str(DepList(DepKey)(iWBSfrom))) Then 'if wbsfrom is a descendent of wbsto
                    InError = InError Or FinishMilestoneID(DepList(DepKey)(iWBSto)) <> DepList(DepKey)(iTasktoID)  ' allowable if to finish milestone
                    If ActiveProject.tasks(DepList(DepKey)(iWBSfrom)).Summary Then
                        InError = InError Or FinishMilestoneID(DepList(DepKey)(iWBSfrom)) <> DepList(DepKey)(iTaskfromID) ' allowable if from finish milestone
                        InError = InError Or ActiveProject.tasks(DepList(DepKey)(iTasktoID)).OutlineLevel + 1 <> ActiveProject.tasks(DepList(DepKey)(iTaskfromID)).OutlineLevel ' immediate = one level different
                    Else
                        InError = InError Or ActiveProject.tasks(DepList(DepKey)(iTasktoID)).OutlineLevel <> ActiveProject.tasks(DepList(DepKey)(iTaskfromID)).OutlineLevel 'immediate means same level
                    End If
                End If
                If InError Then
                    LogErrorTask testNo, ActiveProject.tasks(DepList(DepKey)(iTaskfromID)), "!NameID! has invalid successor dependency to !NameID2!.", ActiveProject.tasks(DepList(DepKey)(iTasktoID))
                End If
            Next DepKey
        End If
        

    End If
    
    
    End If
    
    If numOf(40) = 0 And numOf(2) = 0 And numOf(3) = 0 And numOf(10) = 0 And numOf(11) = 0 And numOf(31) = 0 And numOf(13) = 0 And numOf(41) = 0 And numOf(42) = 0 And numOf(51) = 0 And numOf(52) = 0 Then
    
        For Each tsk In ActiveProject.tasks
            If tsk.ID >= LowID And tsk.ID <= HighID Then
                ' Check that distant dependencies are from an interim milestone.
                If Not tsk.Summary And Not tsk.Milestone And tsk.OutlineLevel > 1 And tsk.SuccessorTasks.Count > 0 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 45) < 0 And IncludedOf(45) Then
                    For Each tsk2 In tsk.SuccessorTasks
                        If tsk.OutlineParent.ID <> tsk2.OutlineParent.ID And Not tsk2.Milestone Then
                            LogErrorTask 48, tsk, "!NameID! distant successor dependency to !NameID2! should be to/from an Interim Milestone.", tsk2
                        End If
                    Next
                End If
        
                ' Check that distant dependencies are to an interim milestone.
                If Not tsk.Summary And Not tsk.Milestone And tsk.OutlineLevel > 1 And tsk.PredecessorTasks.Count > 0 And tskFieldExactMatch(tsk, HealthCheckOptionsID, 48) < 0 And IncludedOf(48) Then
                    For Each tsk2 In tsk.PredecessorTasks
                        If tsk.OutlineParent.ID <> tsk2.OutlineParent.ID And Not tsk2.Milestone Then
                            LogErrorTask 48, tsk, "!NameID2! successor dependency to !NameID! should be to/from an Interim Milestone.", tsk2
                        End If
                    Next
                End If
            End If
        Next
        
    End If
    
    
    message = ""
    
    ' #16 is a global test and the goal is 1
    If numOf(16) = 1 Then
        ClearActualResults 16
    End If
    
    ' #18 is a global test and the goal is 1
    If numOf(18) = 1 Then
        ClearActualResults 18
    End If
    
    ' #32 is a global test and the goal is 1
    If numOf(32) = 1 Then
        ClearActualResults 32
    End If
  
        
    maxSev = 0
    For i = 1 To maxTest
        TotalFound = TotalFound + numOf(i)
        If numOf(i) > 0 Then maxSev = Max(maxSev, sevOf(i))
    Next

    If maxSev <= sevWarning Then
        For Each tsk In ActiveProject.tasks
            ' Progress theme only perform these if no other errors
            If tsk.ID >= LowID And tsk.ID <= HighID Then
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 21) And IncludedOf(21) Then
                    If tsk.Start < ReallyStatusDate() And Not IsDate(tsk.ActualStart) Then
                        LogErrorTask 21, tsk, "!NameID!"
                    End If
                End If
            
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, 22) And IncludedOf(22) Then
                    If tsk.Finish < ReallyStatusDate() And Not IsDate(tsk.ActualFinish) Then
                        LogErrorTask 22, tsk, "!NameID!"
                    End If
                End If
        
                testNo = 25
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(testNo) Then
                    If IsDate(tsk.ActualStart) Then
                        If Int(tsk.ActualStart) > ReallyStatusDate() Then
                            LogErrorTask testNo, tsk, "!NameID!"
                        End If
                    End If
                End If
            
                testNo = 26
                If Not tsk.Summary And tskFieldExactMatch(tsk, HealthCheckOptionsID, testNo) < 0 And IncludedOf(testNo) Then
                    If IsDate(tsk.ActualFinish) Then
                        If Int(tsk.ActualFinish) > ReallyStatusDate() Then
                            LogErrorTask testNo, tsk, "!NameID!"
                        End If
                    End If
                End If
            End If
        Next
    End If
    Dim bandFound(100) As Boolean
    
    TotalFound = 0
    For i = 1 To maxTest
        TotalFound = TotalFound + numOf(i)
        If bandOf(i) = 10 Then bandFound(10) = True
        If bandOf(i) = 20 Then bandFound(20) = True
        If bandOf(i) = 30 Then bandFound(30) = True
        If bandOf(i) = 40 Then bandFound(40) = True
        If bandOf(i) = 40 Then bandFound(50) = True
        If bandOf(i) = 60 Then bandFound(60) = True
        If bandOf(i) = 70 Then bandFound(70) = True
        
    Next
    Dim band As Integer
    For band = 10 To 70 Step 10
        If bandFound(band) And band = 10 Then message = message & "<h2>Project</h1>"
        If bandFound(band) And band = 20 Then message = message & "<h2>Task Identity</h2>"
        If bandFound(band) And band = 50 Then message = message & "<h2>Resources</h2>"
        If bandFound(band) And band = 30 Then message = message & "<h2>WBS/PBS</h2>"
        If bandFound(band) And band = 40 Then message = message & "<h2>Network</h2>"
        If bandFound(band) And band = 60 Then message = message & "<h2>Scheduling</h2>"
        If bandFound(band) And band = 70 Then message = message & "<h2>Progress</h2>"
        For i = 1 To maxTest
            If numOf(i) > 0 And bandOf(i) = band Then
                If sevOf(i) = sevWarning Then
                    severitymessage = "<span style=""color: darkorange; font-size:150%;"">?</span>"
                Else
                    severitymessage = "<span style=""color: darkred; font-size:150%;"">&#x2718;</span>"
                End If
                message = message & severitymessage & "<b> " & i & ". " & descOf(i) & "</b>" & htmlCrLf
                message = message & "<div class=""details"">" & details(i) & "</div>"
                If numOf(i) >= maxpertestplus1 Then message = message & "    ..." & htmlCrLf
            End If
        Next
    Next
    
    If TotalFound = 0 Then message = "<span style=""color: darkgreen; font-size:150%;"">&#x2714;</span> All Good!" & htmlCrLf
    
    message = message & "<div style=""margin-top: 3em; padding-left:3em; padding-right:3em; width:100%; border-top:2px solid #707070;border-bottom:2px solid #707070; background:#d0d0d0;""><h2>Quality Checks <small>(with their check numbers)</small></h2>"
    message = message & "<p><small>Checks may be excluded for individual tasks by listing the unwanted check numbers, comma separated, in a """ & healthcheckoptionsFieldName & """ custom Text Column. Checks may be turned off for the whole project by listing the check numbers in this column in the Project Summary Task.</small></p>"

    For band = 10 To 70 Step 10
        If bandFound(band) And band = 10 Then message = message & "<h3>Project</h3>"
        If bandFound(band) And band = 20 Then message = message & "<h3>Task Identity</h3>"
        If bandFound(band) And band = 50 Then message = message & "<h3>Resources</h3>"
        If bandFound(band) And band = 30 Then message = message & "<h3>WBS/PBS</h3>"
        If bandFound(band) And band = 40 Then message = message & "<h3>Network</h3>"
        If bandFound(band) And band = 60 Then message = message & "<h3>Scheduling</h3>"
        If bandFound(band) And band = 70 Then message = message & "<h3>Progress</h3>"
        For i = 1 To maxTest
            If bandOf(i) = band Then
                message = message & i & ". " & descOf(i) & htmlCrLf
                If i = 36 Then
                    message = message & "<div class=""details""><small>"
                    For j = 1 To 15
                        message = message & settings36(j) & htmlCrLf
                    Next
                    message = message & "</small></div>"
                End If
                
                If i = 47 Then
                   message = message & "<div class=""details""><small>"
                   message = message & settings47(1) & htmlCrLf
                   message = message & "</small></div>"
                End If
                
            End If
        Next
    Next
    message = message & "</div>"
    
   
   
    preamble = "<h1>" & ReportName & " </h1> " & "<p><small><b>v</b>" & ver & "<b> For:</b> " & ActiveProject.FullName & htmlCrLf & "<b>Status Date is</b> " & ReallyStatusDate() & " <b>Created at:</b> " & Now
    If PermanentIDFieldID <> 0 Then
        preamble = preamble & htmlCrLf & "<b>Next unused Permanent ID:</b> " & MaxPermID + 1
    End If
    preamble = preamble & "</small></p>"
    If Testing Then
        preamble = preamble & htmlCrLf & "<p><b>TESTING MODE SEE ALSO THE 'ACTUAL AS EXPECTED' COLUMN</b></p> "
    End If
    
    message = preamble & message
    
    Res.Add "message", message
    Res.Add "TotalFound", TotalFound
    Res.Add "maxSev", maxSev
    Res.Add "StartMilestoneID", StartMilestoneID
    Res.Add "FinishMilestoneID", FinishMilestoneID
    
    Set CheckAnalyse = Res
    
End Function


Sub AddDeleteImplicitDependencies(AddorDelete As String, StartMilestoneID As Dictionary, FinishMilestoneID As Dictionary)
    Dim StartTsk As Task
    Dim FinishTsk As Task
    Dim SiblingStartTsk As Task
    Dim SiblingFinishTsk As Task
    Dim tsk As Task
    Dim Sibling As Task
    Set Lookaside = New Dictionary ' Reset the calculation cache
    For Each tsk In ActiveProject.tasks
        If tsk.Summary Then
            Set StartTsk = ActiveProject.tasks(StartMilestoneID(tsk.ID))
            Set FinishTsk = ActiveProject.tasks(FinishMilestoneID(tsk.ID))
            For Each Sibling In tsk.OutlineChildren
                If Sibling.ID <> StartMilestoneID(tsk.ID) And Sibling.ID <> FinishMilestoneID(tsk.ID) Then
                    If Sibling.Summary Then
                        Set SiblingStartTsk = ActiveProject.tasks(StartMilestoneID(Sibling.ID))
                        Set SiblingFinishTsk = ActiveProject.tasks(FinishMilestoneID(Sibling.ID))
                    Else
                        Set SiblingStartTsk = Sibling
                        Set SiblingFinishTsk = Sibling
                    End If
                    If AddorDelete = "Add" Then
                        StartTsk.LinkSuccessors SiblingStartTsk
                        SiblingFinishTsk.LinkSuccessors FinishTsk
                    Else
                        If AddorDelete = "Delete" Then
                            If UboundFilterExactMatch(StartTsk.Successors, SiblingStartTsk.ID) >= 0 Then StartTsk.UnlinkSuccessors SiblingStartTsk
                            If UboundFilterExactMatch(SiblingFinishTsk.Successors, FinishTsk.ID) >= 0 Then SiblingFinishTsk.UnlinkSuccessors FinishTsk
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub DeleteRedundantDependencies()
    Dim thisSuccessor As Dictionary
    Dim successorsLessOne As Dictionary
    Dim cola As Dictionary
    Dim tsk As Task
    Dim tsk2 As Task
    Set Lookaside = New Dictionary ' Reset the calculation cache
    For Each tsk In ActiveProject.tasks
        If Not tsk.Summary And tsk.SuccessorTasks.Count > 0 Then
            For Each tsk2 In tsk.SuccessorTasks
                Set thisSuccessor = New Dictionary
                thisSuccessor.Add Str(tsk2.ID), Str(tsk2.ID)
                Set successorsLessOne = Subtract(tasks_set(tsk.SuccessorTasks), thisSuccessor)
                Set cola = Intersect(successors_set(successorsLessOne), thisSuccessor)
                If cola.Count() <> 0 Then
                    tsk.UnlinkSuccessors tsk2
                End If
                
            Next
        End If
    Next
End Sub

Function ReallyStatusDate()
If ActiveProject.StatusDate = "NA" Then
    ReallyStatusDate = Date
Else
    ReallyStatusDate = ActiveProject.StatusDate
End If
End Function

Function Min(x As Variant, y As Variant) As Variant
Min = y
If x < y Then Min = x
End Function

Function Max(x As Variant, y As Variant) As Variant
Max = y
If x > y Then Max = x
End Function

Function tasks_set(tasks As tasks) As Dictionary
Dim Res As Dictionary
Set Res = New Dictionary
Dim tsk As Task
For Each tsk In tasks
    Res.Add Str(tsk.ID), Str(tsk.ID)
Next
Set tasks_set = Res
End Function

Function SubPlans_set() As Dictionary
Dim lookasidekey As String
lookasidekey = "SubPlans_set"
If Lookaside.Exists(lookasidekey) Then
    Set SubPlans_set = Lookaside.Item(lookasidekey)
    Exit Function
End If
Dim tsk As Task
Dim Res As New Dictionary
For Each tsk In ActiveProject.tasks
    If tsk.OutlineLevel = 1 And tsk.OutlineChildren.Count > 0 Then ' top level with outline children
        Res.Add Str(tsk.ID), Str(tsk.ID)
    End If
Next
Lookaside.Add lookasidekey, Res
Set SubPlans_set = Res
End Function
Private Function wbs_descendents_set(tsk As Task, Optional recursive As Boolean = True) As Dictionary
    ' Network Descendents
    Dim lookasidekey As String
    lookasidekey = "wbsdescendents_set" & Str(tsk.ID) & "#" & Str(recursive)
    If Lookaside.Exists(lookasidekey) Then
        Set wbs_descendents_set = Lookaside.Item(lookasidekey)
        Exit Function
    End If
    Dim Res As New Dictionary
    Dim subres As Dictionary
    Dim x As Variant
    Dim t As Task
    Dim tsf As Task
    For Each t In tsk.OutlineChildren
        If Left(t.Name, 6) <> "Start " And Left(t.Name, 7) <> "Finish " Then 'filter out the Start Finish Milestones
            If Not Res.Exists(Str(t.ID)) Then Res.Add Str(t.ID), Str(t.ID)
            If t.Summary And recursive Then
                Set subres = wbs_descendents_set(t)
                For Each x In subres
                    If Not Res.Exists(x) Then Res.Add x, x
                Next
            End If
        End If
    Next
    Lookaside.Add lookasidekey, Res
    Set wbs_descendents_set = Res
End Function

Function descendents_set(taskids As Dictionary, Optional recursive As Boolean = True) As Dictionary
Dim lookasidekey As String
lookasidekey = "descendents_set" & Join(taskids.Keys(), "#") & "#" & Str(recursive)
If Lookaside.Exists(lookasidekey) Then
    Set descendents_set = Lookaside.Item(lookasidekey)
    Exit Function
End If
Dim Res As New Dictionary
Dim subres As Dictionary
Dim x As Variant

Dim tid As Variant
Dim t As Task
Dim tsf As Task
For Each tid In taskids
    Set t = ActiveProject.tasks(Val(tid))
    If t.Summary Then
        If recursive Then
            Set subres = descendents_set(tasks_set(t.OutlineChildren))
            For Each x In subres
                If Not Res.Exists(x) Then Res.Add x, x
            Next
        Else
            'just add in the start and finish nodes to represent this sibling summary item
            Set subres = descendents_set(tasks_set(t.OutlineChildren), False)
            For Each x In subres
               Set tsf = ActiveProject.tasks(Val(x))
               If (tsf.Name = "Start " & t.Name Or tsf.Name = "Finish " & t.Name) And Not Res.Exists(x) Then Res.Add x, x
            Next
            
        End If
    Else
        If Not Res.Exists(Str(t.ID)) Then Res.Add Str(t.ID), Str(t.ID)
    End If


Next
Lookaside.Add lookasidekey, Res
Set descendents_set = Res
End Function

Function descendents_set_except(taskids As Dictionary, except As String, Optional recursive As Boolean = True) As Dictionary
Dim lookasidekey As String
lookasidekey = "descendents_set_except" & Join(taskids.Keys(), "#") & "#" & except & "#" & Str(recursive)
If Lookaside.Exists(lookasidekey) Then
    Set descendents_set_except = Lookaside.Item(lookasidekey)
    Exit Function
End If

Dim Res As New Dictionary
Dim subres As Dictionary
Dim x As Variant

Dim tid As Variant
Dim t As Task
Dim tsf As Task
For Each tid In taskids
    Set t = ActiveProject.tasks(Val(tid))
    If Trim(Str(t.ID)) <> Trim(except) Then
        If t.Summary Then
            If recursive Then
                Set subres = descendents_set_except(tasks_set(t.OutlineChildren), except)
                For Each x In subres
                    If Not Res.Exists(x) Then Res.Add x, x
                Next
            Else
                'just add in the start and finish nodes to represent this sibling summary item
                Set subres = descendents_set_except(tasks_set(t.OutlineChildren), except, False)
                For Each x In subres
                   Set tsf = ActiveProject.tasks(Val(x))
                   If (tsf.Name = "Start " & t.Name Or tsf.Name = "Finish " & t.Name) And Not Res.Exists(x) Then Res.Add x, x
                Next
                
            End If
        Else
            If Not Res.Exists(Str(t.ID)) Then Res.Add Str(t.ID), Str(t.ID)
        End If
    End If
Next
Lookaside.Add lookasidekey, Res
Set descendents_set_except = Res
End Function
Function successors_set(taskids As Dictionary, Optional recursive As Boolean = True) As Dictionary
Dim lookasidekey As String
lookasidekey = "successors_set" & Join(taskids.Keys(), "#") & "#" & Str(recursive)
If Lookaside.Exists(lookasidekey) Then
    Set successors_set = Lookaside.Item(lookasidekey)
    Exit Function
End If

Dim Res As New Dictionary
Dim subres As Dictionary
Dim x As Variant

Dim tid As Variant
Dim t As Task
For Each tid In taskids
    Set t = ActiveProject.tasks(Val(tid))
    If t.SuccessorTasks.Count > 0 And recursive Then
        Set subres = successors_set(tasks_set(t.SuccessorTasks))
        For Each x In subres
            If Not Res.Exists(x) Then Res.Add x, x
        Next
    End If
    If Not t.Summary And Not Res.Exists(Str(t.ID)) Then Res.Add Str(t.ID), Str(t.ID)
Next
Lookaside.Add lookasidekey, Res
Set successors_set = Res
End Function

Function sibling_set(tsk As Task) As Dictionary
Dim lookasidekey As String
lookasidekey = "sibling_set" & Str(tsk.ID)
If Lookaside.Exists(lookasidekey) Then
    Set sibling_set = Lookaside.Item(lookasidekey)
    Exit Function
End If
Dim Res As New Dictionary
Dim thisTask As New Dictionary
thisTask.Add Str(tsk.ID), Str(tsk.ID)
Set Res = Subtract(tasks_set(tsk.OutlineParent.OutlineChildren), thisTask)
Lookaside.Add lookasidekey, Res
Set sibling_set = Res

End Function

Function start_successor_set(tsk As Task, StartMilestoneID As Dictionary, FinishMilestones As Dictionary) As Dictionary
Dim lookasidekey As String
lookasidekey = "start_successor_set" & Str(tsk.ID)
If Lookaside.Exists(lookasidekey) Then
    Set start_successor_set = Lookaside.Item(lookasidekey)
    Exit Function
End If

Dim Res As New Dictionary
Set Res = sibling_set(tsk)
Dim r As Variant
For Each r In Res.Keys
    If ActiveProject.tasks(Int(r)).Summary Then ' substitute the Start Milestone of the summary item
        Res.Remove r
        Res.Add Str(StartMilestoneID.Item(Int(r))), Str(StartMilestoneID.Item(Int(r)))
    ElseIf FinishMilestones.Exists(Int(r)) Then ' remove the finish milestone sibling
        Res.Remove r
    End If
Next
Lookaside.Add lookasidekey, Res
Set start_successor_set = Res

End Function

Function finish_predecessor_set(tsk As Task, FinishMilestoneID As Dictionary, StartMilestones As Dictionary) As Dictionary
Dim lookasidekey As String
lookasidekey = "finish_predecessor_set" & Str(tsk.ID)
If Lookaside.Exists(lookasidekey) Then
    Set finish_predecessor_set = Lookaside.Item(lookasidekey)
    Exit Function
End If

Dim Res As New Dictionary
Set Res = sibling_set(tsk)
Dim r As Variant
For Each r In Res.Keys
    If ActiveProject.tasks(Int(r)).Summary Then ' substitute the Finish Milestone of the summary item
        Res.Remove r
        Res.Add Str(FinishMilestoneID.Item(Int(r))), Str(FinishMilestoneID.Item(Int(r)))
    ElseIf StartMilestones.Exists(Int(r)) Then ' remove the start milestone sibling.
        Res.Remove r
    End If
Next
Lookaside.Add lookasidekey, Res
Set finish_predecessor_set = Res

End Function

' assumes that the list of 'predecessor tasks' is initially passed in
Function predecessors_set(taskids As Dictionary, Optional recursive As Boolean = True) As Dictionary
Dim lookasidekey As String
lookasidekey = "predecessors_set" & Join(taskids.Keys(), "#") & "#" & Str(recursive)
If Lookaside.Exists(lookasidekey) Then
    Set predecessors_set = Lookaside.Item(lookasidekey)
    Exit Function
End If

Dim Res As New Dictionary
Dim subres As Dictionary
Dim x As Variant

Dim tid As Variant
Dim t As Task
For Each tid In taskids
    Set t = ActiveProject.tasks(Val(tid))
    If t.PredecessorTasks.Count > 0 And recursive Then
        Set subres = predecessors_set(tasks_set(t.PredecessorTasks))
        For Each x In subres
            If Not Res.Exists(x) Then Res.Add x, x
        Next
    End If
    If Not t.Summary And Not Res.Exists(Str(t.ID)) Then Res.Add Str(t.ID), Str(t.ID)
Next
Lookaside.Add lookasidekey, Res
Set predecessors_set = Res
End Function
Function printkeys(col As Dictionary)
    Dim x As Variant
    For Each x In col.Keys

        Debug.Print "'" & x & "'"
    Next
End Function

Function Intersect(col1 As Dictionary, col2 As Dictionary) As Dictionary
Dim item1 As Variant
Dim col3 As Dictionary
Set col3 = New Dictionary
For Each item1 In col1
    If col2.Exists(item1) Then
        col3.Add item1, item1
    End If
Next
Set Intersect = col3
End Function
' col2 is a subset of col1
Function Subset(col1 As Dictionary, col2 As Dictionary) As Boolean
Dim item1 As Variant
Dim col3 As Dictionary
Set col3 = New Dictionary
Subset = True
For Each item1 In col2
    If Not col1.Exists(item1) Then
        Subset = False
        Exit Function
    End If
Next
End Function
Function Subtract(col1 As Dictionary, col2 As Dictionary) As Dictionary
Dim item1 As Variant
Dim col3 As Dictionary
Set col3 = New Dictionary
For Each item1 In col1
    'Debug.Print item1
    If Not col2.Exists(item1) Then
        col3.Add item1, item1
    End If
Next
Set Subtract = col3
End Function

Private Function InvertDictionary(col1 As Dictionary) As Dictionary
    Dim key1 As Variant
    Dim col2 As New Dictionary
    For Each key1 In col1
       col2.Add col1(key1), key1
    Next key1
    Set InvertDictionary = col2
End Function
Function firstword(x As String) As String
If InStr(Trim(x), " ") > 1 Then
    firstword = Mid(Trim(x), 1, InStr(Trim(x), " ") - 1)
Else
    firstword = Trim(x)
End If
End Function

Function UboundFilterExactMatch(astrItems As String, _
                          intSearch As Integer) As Long
                  
   ' This function searches a string array for elements
   ' that exactly match the search string.
   ' http://msdn.microsoft.com/en-us/library/office/aa164525(v=office.10).aspx

   Dim astrFilter()   As String
   Dim astrTemp()       As String
   Dim lngUpper         As Long
   Dim lngLower         As Long
   Dim lngIndex         As Long
   Dim lngCount         As Long
   Dim UBoundResult As Long
   UBoundResult = -1
   
   ' Filter array for search string.
   astrFilter = Split(astrItems, ",")
   
   ' Store upper and lower bounds of resulting array.
   lngUpper = UBound(astrFilter)
   lngLower = LBound(astrFilter)
   
   If lngUpper = -1 Then
        UboundFilterExactMatch = -1
        Exit Function
   End If
   
   ' Resize temporary array to be same size.
   ReDim astrTemp(lngLower To lngUpper)
   
   ' Loop through each element in filtered array.
   For lngIndex = lngLower To lngUpper
      ' Check that element matches search string exactly.
      If Int(astrFilter(lngIndex)) = intSearch Then
         ' Store elements that match exactly in another array.
        UBoundResult = UBoundResult + 1
      End If
   Next lngIndex
   
   ' Return array containing exact matches.
   UboundFilterExactMatch = UBoundResult
End Function

Function tskFieldExactMatch(tsk As Task, FieldID As Long, _
                          intSearch As Integer) As Long
                  
   ' This function searches a string array for elements
   ' that exactly match the search string.
   ' http://msdn.microsoft.com/en-us/library/office/aa164525(v=office.10).aspx

   Dim astrFilter()   As String
   Dim astrTemp()       As String
   Dim lngUpper         As Long
   Dim lngLower         As Long
   Dim lngIndex         As Long
   Dim lngCount         As Long
   Dim UBoundResult As Long
   UBoundResult = -1
   
   If FieldID = 0 Then
    tskFieldExactMatch = UBoundResult
    Exit Function
   End If
   
   ' Filter array for search string.
   astrFilter = Split(tsk.GetField(FieldID), ",")
   
   ' Store upper and lower bounds of resulting array.
   lngUpper = UBound(astrFilter)
   lngLower = LBound(astrFilter)
   
   If lngUpper = -1 Then
        tskFieldExactMatch = -1
        Exit Function
   End If
   
   ' Resize temporary array to be same size.
   ReDim astrTemp(lngLower To lngUpper)
   
   ' Loop through each element in filtered array.
   For lngIndex = lngLower To lngUpper
      ' Check that element matches search string exactly.
      If Int(astrFilter(lngIndex)) = intSearch Then
         ' Store elements that match exactly in another array.
        UBoundResult = UBoundResult + 1
      End If
   Next lngIndex
   
   ' Return array containing exact matches.
   tskFieldExactMatch = UBoundResult
End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0
End Function



