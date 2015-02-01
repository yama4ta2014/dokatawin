Attribute VB_Name = "Module1"
Option Explicit


Dim conn As New ADODB.Connection
Const fixx As String = "ABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJ"
Const fix2 As String = "àÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZàÍìÒéOélå‹òZéµî™ã„ÅZ"
Function test()

    Set conn = New ADODB.Connection
    
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dokatawin\db1.mdb"
    conn.Open
    On Error GoTo CatchException
    
    conn.BeginTrans
    On Error GoTo CatchExceptionB
    
    loaddata 100
    
    If MsgBox("COMMIT?", vbQuestion + vbOKCancel) = vbOK Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    On Error GoTo CatchException
    
    conn.Close
    Exit Function
    
    
CatchExceptionB:
    conn.RollbackTrans
    
CatchException:
    
    conn.Close
    MsgBox Err.Description

End Function

Function loaddata(cnt As Long)

    Dim r As New ADODB.Recordset
     
    r.Open "M_KOKYAK", conn, adOpenDynamic, adLockOptimistic
    
    
    If r Is Nothing Then
    
    Else
        
        Dim ccc As Long
        ccc = 0
        Do While ccc < cnt
            addItems r, ccc
            ccc = ccc + 1
        Loop
        
    
    End If
    
    r.Close
    
    

End Function

Function addItems(r As ADODB.Recordset, yobi As Long)

    Dim c As Range
    Dim row As Integer
    Dim col As Integer
    
    Dim min As String
    Dim max As String
    Dim fld As String
    Dim typ As String
    Dim frm As String
    
    Dim x As Long
    Dim x2 As String
    
    r.AddNew
    
    row = -1
    For Each c In ActiveWindow.Selection
    
        If (row <> c.row) Then
            
            If (row >= 0) Then
            
        
                Select Case typ
                Case "NUMSEQ"
                
                    x = min + yobi
                    If frm <> "" Then
                        x2 = Format(x, frm)
                        r.Fields(fld).Value = x2
                    Else
                        r.Fields(fld).Value = x
                    End If
                    
                
                Case "NUM"
                    x = Rnd
                    Dim v As Long
                    x = Fix((max - min + 1) * Rnd)
                    If frm <> "" Then
                        x2 = Format(x, frm)
                        r.Fields(fld).Value = x2
                    Else
                        r.Fields(fld).Value = x
                    End If
                    
                                
                
                Case "STRING"
                    x = min + Fix((max - min + 1) * Rnd)
                    r.Fields(fld).Value = Left(fixx, x)
                
                Case "JSTRING"
                    x = min + Fix((max - min + 1) * Rnd)
                    r.Fields(fld).Value = Left(fix2, x)
                
                Case "CODE"
                Case "DATE"
                    Dim dd As Long
                    dd = DateDiff("d", min, max)
                    x = Fix(dd * Rnd)
                    r.Fields(fld).Value = DateAdd("d", x, min)
                
                End Select
            
            End If
                        
            
            col = 0
        End If
        row = c.row
    
        Select Case col
        Case 0
            fld = c.Value
            col = col + 1
        Case 1
            min = c.Value
            col = col + 1
        Case 2
            max = c.Value
            col = col + 1
        Case 3
            typ = c.Value
            col = col + 1
        Case 4
            If Trim(c.Value) <> "" Then
                frm = Mid(c.Value, 3)
            End If
            col = col + 1
        End Select
    
    
    Next
    If (row > 0) Then
        r.Update
    Else
        r.Cancel
    End If
    
    
End Function
