﻿Imports Microsoft.Office.Interop.MSProject

Public Class ThisAddIn


    'Public Function LagType(ByVal pj As Microsoft.Office.Interop.MSProject.Project) As String
    '    Dim TaskDep As String
    '    Dim project As Microsoft.Office.Interop.MSProject.Project
    '    project = pj
    '    TaskDep = project.Tasks(5).TaskDependencies(1).Type
    '    Return TaskDep
    'End Function

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
        Dim P As Microsoft.Office.Interop.MSProject.TaskDependency
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

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

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

End Class
