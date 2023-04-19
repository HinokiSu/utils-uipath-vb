 ' Get output datatable formatted string
 ' Reference: https://stackoverflow.com/questions/691363/printing-a-datatable-to-textbox-textfile-in-net
 Public Function GetFormattedDataTableString(dataTable As DataTable) As String
        If dataTable Is Nothing Then Throw New ArgumentNullException("'dataTable' cannot be null.")

        Dim representationWriter As New StringWriter()

        ' First, set the width of every column to the length of its largest element.
        Dim columnWidths = New Integer(dataTable.Columns.Count - 1) {}
        For columnIndex As Integer = 0 To dataTable.Columns.Count - 1
            Dim headerWidth As Integer = dataTable.Columns(columnIndex).ColumnName.Length
            Dim longestElementWidth As Integer = dataTable.AsEnumerable().[Select](Function(row) row(columnIndex).ToString().Length).Max()
            columnWidths(columnIndex) = Math.Max(headerWidth, longestElementWidth)
        Next


        ' Next, write the table
        ' Write a horizontal line.
        representationWriter.Write("+-")
        For columnIndex As Integer = 0 To dataTable.Columns.Count - 1
            For i = 0 To columnWidths(columnIndex) - 1
                representationWriter.Write("-")
            Next
            representationWriter.Write("-+")
            If columnIndex <> dataTable.Columns.Count - 1 Then representationWriter.Write("-")
        Next
        representationWriter.WriteLine(" ")
        ' Print the headers
        representationWriter.Write("| ")
        For columnIndex As Integer = 0 To dataTable.Columns.Count - 1
            Dim header As String = dataTable.Columns(columnIndex).ColumnName
            representationWriter.Write(header)
            For blanks = columnWidths(columnIndex) - header.Length To 1 Step -1
                representationWriter.Write(" ")
            Next
            representationWriter.Write(" | ")
        Next
        representationWriter.WriteLine()
        ' Print another horizontal line.
        representationWriter.Write("+-")
        For columnIndex As Integer = 0 To dataTable.Columns.Count - 1
            For i = 0 To columnWidths(columnIndex) - 1
                representationWriter.Write("-")
            Next
            representationWriter.Write("-+")
            If columnIndex <> dataTable.Columns.Count - 1 Then representationWriter.Write("-")
        Next
        representationWriter.WriteLine(" ")

        ' Print the contents of the table.
        For row As Integer = 0 To dataTable.Rows.Count - 1
            representationWriter.Write("| ")
            For column As Integer = 0 To dataTable.Columns.Count - 1
                representationWriter.Write(dataTable.Rows(row)(column))
                For blanks = columnWidths(column) - dataTable.Rows(row)(column).ToString().Length To 1 Step -1
                    representationWriter.Write(" ")
                Next
                representationWriter.Write(" | ")
            Next
            representationWriter.WriteLine()
        Next

        ' Print a final horizontal line.
        representationWriter.Write("+-")
        For column As Integer = 0 To dataTable.Columns.Count - 1
            For i = 0 To columnWidths(column) - 1
                representationWriter.Write("-")
            Next
            representationWriter.Write("-+")
            If column <> dataTable.Columns.Count - 1 Then representationWriter.Write("-")
        Next
        representationWriter.WriteLine(" ")

        Return representationWriter.ToString()
    End Function

