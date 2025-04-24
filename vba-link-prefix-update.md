Do Alt + F11 to open VBA.  Select "Insert a Module", paste the code and run it with F5.

```
Sub UpdateHyperlinkPrefixes()
    Dim ws As Worksheet
    Dim hl As Hyperlink
    Dim oldPrefix As String
    Dim newPrefix As String
    
    oldPrefix = "https://stacks.loc.gov/item/"
    newPrefix = "https://hdl.loc.gov/loc.sgp/npe."
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each hl In ws.Hyperlinks
            If InStr(hl.Address, oldPrefix) = 1 Then
                hl.Address = Replace(hl.Address, oldPrefix, newPrefix)
            End If
            If hl.TextToDisplay Like "*" & oldPrefix & "*" Then
                hl.TextToDisplay = Replace(hl.TextToDisplay, oldPrefix, newPrefix)
            End If
        Next hl
    Next ws
End Sub
```
