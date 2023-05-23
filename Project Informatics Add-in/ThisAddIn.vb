Imports Microsoft.Office.Interop.MSProject

Public Class ThisAddIn

    Public C1_Tasks As Integer
    Public C2_Tasksover25Days As Integer
    Public C3_Notindays As Integer
    Public C4_PlannedinPast As Integer
    Public C5_NoSuccessor As Integer
    Public C6_NoPredeccessor As Integer
    Public C7_LinkstosummaryTasks As Integer
    Public C8_StartFinish As Integer
    Public C9_StartStart As Integer
    Public C10_FinishFinish
    Public C11_PositiveLag As Integer
    Public C12_NegativeLag As Integer
    Public C13_AsLateAsPossible As Integer
    Public C14_MustStartOn As Integer
    Public C15_MustFinishOn As Integer
    Public C16_StartNoEarlierThan As Integer
    Public C17_StarNoLaterThan As Integer
    Public C18_FinishNoEarlierThan As Integer
    Public C19_FinishNoLaterThan As Integer
    Public C20_LargeTotalSlack As Integer
    Public C21_NegativeSlack As Integer

    Public Sub SetupGrades(ByVal pj As Microsoft.Office.Interop.MSProject.Project)
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj

        Globals.ThisAddIn.C1_Tasks = 0
        Globals.ThisAddIn.C2_Tasksover25Days = 0
        Globals.ThisAddIn.C3_Notindays = 0
        Globals.ThisAddIn.C4_PlannedinPast = 0
        Globals.ThisAddIn.C5_NoSuccessor = 0
        Globals.ThisAddIn.C6_NoPredeccessor = 0
        Globals.ThisAddIn.C7_LinkstosummaryTasks = 0
        Globals.ThisAddIn.C8_StartFinish = 0
        Globals.ThisAddIn.C9_StartStart = 0
        Globals.ThisAddIn.C10_FinishFinish = 0
        Globals.ThisAddIn.C11_PositiveLag = 0
        Globals.ThisAddIn.C12_NegativeLag = 0
        Globals.ThisAddIn.C13_AsLateAsPossible = 0
        Globals.ThisAddIn.C14_MustStartOn = 0
        Globals.ThisAddIn.C15_MustFinishOn = 0
        Globals.ThisAddIn.C16_StartNoEarlierThan = 0
        Globals.ThisAddIn.C17_StarNoLaterThan = 0
        Globals.ThisAddIn.C18_FinishNoEarlierThan = 0
        Globals.ThisAddIn.C19_FinishNoLaterThan = 0
        Globals.ThisAddIn.C20_LargeTotalSlack = 0
        Globals.ThisAddIn.C21_NegativeSlack = 0

        For Each x_task In project.Tasks
            ' 1) C1_Tasks 
            If x_task.OutlineChildren.Count = 0 Then Globals.ThisAddIn.C1_Tasks = Globals.ThisAddIn.C1_Tasks + 1
            ' 2) C2_Tasks over 25 days 
            If x_task.Duration > (project.HoursPerDay * 25 * 60) And x_task.OutlineChildren.Count = 0 Then Globals.ThisAddIn.C2_Tasksover25Days = Globals.ThisAddIn.C2_Tasksover25Days + 1

            '3) C3_Not In days 
            'NEED TO CREATE \

            '4) C4_PlannedinPast 
            If x_task.FinishVariance > 0 And x_task.OutlineChildren.Count = 0 Then Globals.ThisAddIn.C4_PlannedinPast = Globals.ThisAddIn.C4_PlannedinPast + 1
            ' 5) C5_NoSuccessor 
            If x_task.Successors = "" And x_task.OutlineChildren.Count = 0 Then Globals.ThisAddIn.C5_NoSuccessor = Globals.ThisAddIn.C5_NoSuccessor + 1
            '6) C6_NoPredeccesor 
            If x_task.Predecessors = "" And x_task.OutlineChildren.Count = 0 Then Globals.ThisAddIn.C6_NoPredeccessor = Globals.ThisAddIn.C6_NoPredeccessor + 1
            '7) C7_ Links to Summary Task
            If (x_task.TaskDependencies.Count > 0) And x_task.OutlineChildren.Count > 0 Then
                Globals.ThisAddIn.C7_LinkstosummaryTasks = Globals.ThisAddIn.C7_LinkstosummaryTasks + x_task.TaskDependencies.Count
            End If
            '8) C6_ StartFinish + 9)StartStart + 10) FinishFinish 
            For Each Tdep In x_task.TaskDependencies
                '8) Start Finish
                If Tdep.Type = 2 Then Globals.ThisAddIn.C8_StartFinish = Globals.ThisAddIn.C8_StartFinish + 1
                '9) Start Start 
                If Tdep.Type = 3 Then Globals.ThisAddIn.C9_StartStart = Globals.ThisAddIn.C9_StartStart + 1
                '10) Finish FInish 
                If Tdep.Type = 0 Then Globals.ThisAddIn.C10_FinishFinish = Globals.ThisAddIn.C10_FinishFinish + 1
            Next
            '11) Positive Lag + 12) NegativeLag 
            For Each Pred In x_task.Predecessors
                If Pred.ToString.Contains("+") Then Globals.ThisAddIn.C11_PositiveLag = Globals.ThisAddIn.C11_PositiveLag + 1
                If Pred.ToString.Contains("-") Then Globals.ThisAddIn.C12_NegativeLag = Globals.ThisAddIn.C12_NegativeLag + 1
            Next
            '13) AsLateAsPossible 
            If x_task.ConstraintType = 1 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C13_AsLateAsPossible = Globals.ThisAddIn.C13_AsLateAsPossible + 1
            End If
            '14) Must Start On 
            If x_task.ConstraintType = 2 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C14_MustStartOn = Globals.ThisAddIn.C14_MustStartOn + 1
            End If
            '15) MustFinishOn
            If x_task.ConstraintType = 3 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C15_MustFinishOn = Globals.ThisAddIn.C15_MustFinishOn + 1
            End If
            '16) Start No Earlier Than 
            If x_task.ConstraintType = 4 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C16_StartNoEarlierThan = Globals.ThisAddIn.C16_StartNoEarlierThan + 1
            End If
            '17) Start No Later Than 
            If x_task.ConstraintType = 5 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C17_StarNoLaterThan = Globals.ThisAddIn.C17_StarNoLaterThan + 1
            End If
            '18) Finish No Earlier Than 
            If x_task.ConstraintType = 6 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C18_FinishNoEarlierThan = Globals.ThisAddIn.C18_FinishNoEarlierThan + 1
            End If
            '19) Finish No later Than 
            If x_task.ConstraintType = 7 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C19_FinishNoLaterThan = Globals.ThisAddIn.C19_FinishNoLaterThan + 1
            End If
            ' 20) Large total Slack  כרגע "סלק גדול הוא מגדיור כ"סלק" גדול מ40 יום אך בעתיד אולי זה יהיה מוגדר מהמשתמש
            If x_task.TotalSlack > 40 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C20_LargeTotalSlack = Globals.ThisAddIn.C20_LargeTotalSlack + 1
            End If
            '21) Negative Slack 
            If x_task.TotalSlack < 0 And x_task.OutlineChildren.Count = 0 Then
                Globals.ThisAddIn.C21_NegativeSlack = Globals.ThisAddIn.C21_NegativeSlack + 1
            End If
        Next
        ' לקישורים סס, הה, הס הוא סופר פעמיים את הקישור ולכן צריכים לחלק ל2 לקבל מספר הקישורים ו לא מספר הפעולות עם קישור כאלו
        Globals.ThisAddIn.C8_StartFinish = Globals.ThisAddIn.C8_StartFinish / 2
        Globals.ThisAddIn.C9_StartStart = Globals.ThisAddIn.C9_StartStart / 2
        Globals.ThisAddIn.C10_FinishFinish = Globals.ThisAddIn.C10_FinishFinish / 2

    End Sub

    ' יש כל פונקציה בנפרד בחלק הקוד הזה- אולי נרצה להשתמש בו בעתיד 
    Public Function C_LargeTotalSlack(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.TotalSlack > 40 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next

        Return task_count
    End Function

    Public Function C_NegativeSlack(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.TotalSlack < 0 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next

        Return task_count
    End Function

    Public Function C_NoSuccessor(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.Successors = "" And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next

        Return task_count
    End Function

    Public Function C_StartFinish(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        Dim P As Microsoft.Office.Interop.MSProject.TaskDependency
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks

            For Each Tdep In x_task.TaskDependencies
                If Tdep.Type = 2 Then
                    task_count = task_count + 1

                End If
            Next

        Next


        Return task_count / 2
    End Function

    Public Function C_StartStart(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        Dim P As Microsoft.Office.Interop.MSProject.TaskDependency
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks

            For Each Tdep In x_task.TaskDependencies
                If Tdep.Type = 3 Then
                    task_count = task_count + 1

                End If
            Next

        Next
        Return task_count / 2
    End Function

    Public Function C_FinishFinish(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        Dim Tdep As Microsoft.Office.Interop.MSProject.TaskDependency
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks

            For Each Tdep In x_task.TaskDependencies
                If Tdep.Type = 0 Then
                    task_count = task_count + 1

                End If
            Next
        Next
        Return task_count / 2
    End Function

    Public Function C_PositiveLag(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task

        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks

            For Each Pred In x_task.Predecessors
                If Pred.ToString.Contains("+") Then
                    task_count = task_count + 1
                End If
            Next
        Next
        Return task_count
    End Function

    Public Function C_NegativeLag(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task

        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks

            For Each Pred In x_task.Predecessors
                If Pred.ToString.Contains("-") Then
                    task_count = task_count + 1
                End If
            Next
        Next
        Return task_count
    End Function


    Public Function C_AsLateaspossible(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 1 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function
    Public Function C_MustStartOn(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 2 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function

    Public Function C_MustFinishOn(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 3 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function

    Public Function C_StartNoEarlierThan(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 4 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function

    Public Function C_StartNoLaterThan(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 5 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function

    Public Function C_FinishNoEarlierThan(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 6 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function

    Public Function C_FinishNoLaterThan(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.ConstraintType = 7 And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function

    Public Function C_Linkstosummarytasks(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If (x_task.Successors <> "" Or x_task.Predecessors <> "") And x_task.OutlineChildren.Count > 0 Then
                task_count = task_count + x_task.Successors.Count + x_task.Predecessors.Count

            End If
        Next

        Return task_count
    End Function

    Public Function C_NoPredecessor(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.Predecessors = "" And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next

        Return task_count
    End Function
    Public Function Count_tasksover25(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Double = 0
        For Each x_task In project.Tasks
            If x_task.Duration > (project.HoursPerDay * 60 * 25 * 60) And x_task.OutlineChildren.Count = 0 Then
                task_count = task_count + 1
            End If
        Next

        Return task_count
    End Function


    Public Function Count_tasks(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As Integer
        'סופר כל פעולה שאינה ערסל או קבוצה
        Dim project As Microsoft.Office.Interop.MSProject.Project
        Dim x_task As Microsoft.Office.Interop.MSProject.Task
        project = pj
        Dim task_count As Integer = 0
        For Each x_task In project.Tasks

            If x_task.OutlineChildren.Count = 0 Then

                task_count = task_count + 1
            End If
        Next
        Return task_count
    End Function
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub


End Class
