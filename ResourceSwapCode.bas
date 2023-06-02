Attribute VB_Name = "ResourceSwapCode"
Sub ReplaceResource(ByVal res1 As Resource, ByVal res2 As Resource)
    Dim projApp As Application
    Dim proj As Project
    Dim resAssignment As Assignment
    Dim newAssignment As Assignment
    Dim taskManual As Boolean
     
    ' Set a reference to the active instance of Microsoft Project
    Set projApp = Application
     
    ' Check if a project is open
    If Not projApp Is Nothing Then
        ' Get the active project
        Set proj = projApp.ActiveProject
         
        ' Check if a project is loaded
        If Not proj Is Nothing Then
            ' Iterate through all assignments of res1
            For Each resAssignment In res1.Assignments
                ' Get the parent task of the assignment
                Dim task As task
                Set task = resAssignment.task
                
                If Not task Is Nothing And Not task.Summary Then
                    ' Set the task to be manually scheduled if it's not already
                    ' Necssary because MSP will do wacky things with work and duration as we handover to new resource
                    If task.Manual = False Then
                        taskManual = False
                        task.Manual = True
                    Else
                        taskManual = True
                    End If
                    ' Create a new assignment for res2
                    Set newAssignment = task.Assignments.Add(, res2.ID)
                     
                    ' Copy assignment parameters from res1 to newAssignment
                    newAssignment.Start = resAssignment.Start
                    newAssignment.Finish = resAssignment.Finish
                    newAssignment.Units = resAssignment.Units
                    newAssignment.Work = resAssignment.Work
                    newAssignment.PercentWorkComplete = resAssignment.PercentWorkComplete
                    newAssignment.ActualWork = resAssignment.ActualWork
                    newAssignment.ActualStart = resAssignment.ActualStart
                    newAssignment.ActualFinish = resAssignment.ActualFinish
                    'newAssignment.ActualDuration = resAssignment.ActualDuration
                    newAssignment.RemainingWork = resAssignment.RemainingWork
                    newAssignment.BaselineStart = resAssignment.BaselineStart
                    newAssignment.BaselineFinish = resAssignment.BaselineFinish
                    newAssignment.BaselineWork = resAssignment.BaselineWork
                    'newAssignment.ActualCost = resAssignment.ActualCost
                     
                    ' Delete the assignment for res1
                    resAssignment.Delete
                
                    ' Restore the task scheduling mode to its original setting
                    If taskManual = False Then
                        task.Manual = False
                    End If
                End If
            Next resAssignment
             
            ' Notify the user that the resource has been replaced
            MsgBox "Resource " & res1.Name & " has been replaced with " & res2.Name & " for all tasks."
        Else
            ' Notify the user that a project is not loaded
            MsgBox "No project is loaded."
        End If
    Else
        ' Notify the user that Microsoft Project is not open
        MsgBox "Microsoft Project is not open."
    End If
End Sub


