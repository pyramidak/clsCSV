#Region " CSV "

Class clsCSV

    Public Function GetDataTable(filename As String) As Data.DataTable
        Dim csvData As Data.DataTable = New Data.DataTable()

        Try
            Using csvReader As FileIO.TextFieldParser = New FileIO.TextFieldParser(filename)
                csvData = GetDataTable(csvReader)
            End Using
        Catch ex As Exception
        End Try

        Return csvData
    End Function

    Public Function GetDataTable(stream As IO.MemoryStream) As Data.DataTable
        Dim csvData As Data.DataTable = New Data.DataTable()

        stream.Seek(0, IO.SeekOrigin.Begin)
        Try
            Using csvReader As FileIO.TextFieldParser = New FileIO.TextFieldParser(stream)
                csvData = GetDataTable(csvReader)
            End Using
        Catch ex As Exception
        End Try

        Return csvData
    End Function

    Private Function GetDataTable(csvReader As FileIO.TextFieldParser) As Data.DataTable
        Dim csvData As Data.DataTable = New Data.DataTable()

        csvReader.SetDelimiters(New String() {","})
        csvReader.HasFieldsEnclosedInQuotes = True

        Dim colFields As String() = csvReader.ReadFields()
        For Each column As String In colFields
            Dim datecolumn As Data.DataColumn = New Data.DataColumn(column)
            datecolumn.AllowDBNull = True
            csvData.Columns.Add(datecolumn)
        Next

        csvData.Rows.Add(colFields) 'first row read as data
        While Not csvReader.EndOfData
            Dim fieldData As String() = csvReader.ReadFields()
            csvData.Rows.Add(fieldData)
        End While

        Return csvData
    End Function

End Class

#End Region