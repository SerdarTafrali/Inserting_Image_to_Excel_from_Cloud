Sub ImageInsertBySerdarTAFRALI()
  'Here's an example VBA macro that you can use to insert images to column A 
  'from links that are constructed using the values in column B. 
  'This macro will use the =IMAGE() function to display the images in the cells of column A.
  
  'Set column width for column A
With Columns("A")
        .ColumnWidth = 18
          
    End With
    
    Dim LastRow As Long
    'The macro takes a part of URL from a spesified column -in this example columb B-
    LastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    'The loop written to iterate through a spesified column -in this example columb B-
    For i = 2 To LastRow
        'Setting images columns height
        Rows(i).RowHeight = 150
        Dim URL As String
        'The Image URL that some part of it taken from a spesified column
        URL = "https://write-cloud-url-here.com/example" & Range("B" & i).Value & "_01.jpg"
        'Pasting Image to a spesified column
        ActiveSheet.Range("A" & i).Formula = "=IMAGE(""" & URL & """,1)"
    Next i

End Sub
