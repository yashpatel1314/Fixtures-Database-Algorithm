Sub Appear()
    Dim name As String
    Dim depart As String
    Dim used As String
    Dim supply As String
    Dim status As String
    Dim fix_num As String
    Dim Destination1 As String
    Dim Destination2 As String
    Dim imageFile As String
    Dim img_name As String
    
    name = Application.WorksheetFunction.VLookup(Sheets("Input").Cells(3, 3), Sheets("Fixtures").Range("B1:I200"), 3, False)
    depart = Application.WorksheetFunction.VLookup(Sheets("Input").Cells(3, 3), Sheets("Fixtures").Range("B1:I200"), 4, False)
    used = Application.WorksheetFunction.VLookup(Sheets("Input").Cells(3, 3), Sheets("Fixtures").Range("B1:I200"), 5, False)
    supply = Application.WorksheetFunction.VLookup(Sheets("Input").Cells(3, 3), Sheets("Fixtures").Range("B1:I200"), 6, False)
    status = Application.WorksheetFunction.VLookup(Sheets("Input").Cells(3, 3), Sheets("Fixtures").Range("B1:I200"), 7, False)
    
    Sheets("Input").Cells(11, 5) = name
    Sheets("Input").Cells(13, 5) = depart
    Sheets("Input").Cells(15, 5) = used
    Sheets("Input").Cells(17, 5) = supply
    Sheets("Input").Cells(19, 5) = status
    
    fix_num = Sheets("Input").Cells(3, 3)
    
    img_name = fix_num & ".HEIC"
   Destination1 = "\\clcrs010\CLCOvensGroup\Manufacturing\Pub\BTP 2020\ME Projects\Jigs and Fixtures\Photo Archive\"
Destination2 = "\Pictures\"
    imageFile = Destination1 & fix_num & Destination2 & img_name

    Dim ws As Worksheet
    Set ws = Sheets("Input")
    
    Dim pic As Object
    Set pic = ws.Shapes.AddPicture(imageFile, msoFalse, msoTrue, 0, 0, -1, -1)
    pic.Top = ws.Range("E26").Top
    pic.Left = ws.Range("E26").Left
    pic.Width = ws.Range("E26").Width
    pic.Height = ws.Range("E26").Height
    
End Sub
