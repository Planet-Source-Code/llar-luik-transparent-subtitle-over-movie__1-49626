Attribute VB_Name = "mdvd_moodul"
Option Explicit

Dim temp As String
Dim ix As Integer
Dim ix2 As Integer
Dim ix3 As Integer
Dim ix4 As Integer
Dim SisestusPikkus As Integer
Dim lngStart As Long
Dim lngEnd As Long
Dim strText As String
Dim i As Integer
Dim ii As Integer
Dim var As ListItem

Global subnr As Integer
Global subtext As String
Global startshow As String
Global stopshow As String

Public Function OpenSub_mDVD(Filename As String)

    On Error GoTo hell
    
    With frm_main
    
        .ListView1.ListItems.Clear
        
        Open Filename For Input As #1
        
            Do While Not EOF(1)
      
                Line Input #1, temp
                ix = 1
        
                SisestusPikkus = Len(temp)
        
                For ii = 1 To SisestusPikkus
                    ix = ix + 1
                    If Mid(temp, ix, 1) = "}" Then
                        ix2 = ix
                        ii = SisestusPikkus - 1
                    End If
                Next ii
        
                lngStart = Mid(temp, 2, ix2 - 2)
                ix3 = ix2
                ix = ix3
            
                For ii = ix3 To SisestusPikkus
                    ix = ix + 1
                    If Mid(temp, ix, 1) = "}" Then
                        ix2 = ix
                    End If
                Next ii
        
                ix4 = ix2
                ix2 = ix2 - ix3
                lngEnd = Mid(temp, ix3 + 2, ix2 - 2)
        
                strText = Mid(temp, ix4 + 1, (SisestusPikkus - ix4))
        
                Set var = .ListView1.ListItems.Add(, "", lngStart, 0, 0)
                    var.SubItems(1) = lngEnd
                    var.SubItems(2) = strText
            
            Loop
      
        Close #1
    
hell:

        Close #1
        
    End With

End Function

Public Sub Read_mDVD_subtitle()
    
    subtext = Replace(frm_main.ListView1.ListItems(subnr).SubItems(2), "|", "" & vbCrLf & "")
    startshow = frm_main.ListView1.ListItems(subnr).Text
    stopshow = frm_main.ListView1.ListItems(subnr).SubItems(1)
    
    subnr = subnr + 1

End Sub
