Imports System.IO

Public Class Form2

    Dim OpenFileDialog As New OpenFileDialog()
    Dim listFilename As New List(Of String)
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim dt3 As DataTable
    Dim dt4 As DataTable
    Dim listDt As New List(Of DataTable)
    Dim listDt2 As New List(Of DataTable)
    Dim newDT As DataTable
    Dim dt5 As DataTable
    Dim dt6 As DataTable

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        OpenFileDialog.Filter = "CSV File(*.csv)|*.csv"
        OpenFileDialog.Multiselect = True
        If OpenFileDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Dim i As Integer
            Dim Filename As String
            Dim stFilename As String
            Dim stFilePath As String
            For i = 0 To OpenFileDialog.SafeFileNames.Count - 1
                Filename = OpenFileDialog.SafeFileNames(i)
                stFilePath = OpenFileDialog.FileNames(i)
                If i < OpenFileDialog.SafeFileNames.Count - 1 Then
                    stFilename += Filename
                    stFilename += " , "
                Else
                    stFilename += Filename
                End If
                listFilename.Add(stFilePath)
            Next
            Me.txtFilename.Text = stFilename
        Else
            Exit Sub
        End If
    End Sub
    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        listDt = ReadFiles()
        dt4 = WriteFile(listDt)
    End Sub
    Private Function ReadFiles()
        Dim fileExists As Boolean
        Dim dt As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim i As Integer
        For i = 0 To listFilename.Count - 1
            Dim Filename = listFilename(i)
            fileExists = My.Computer.FileSystem.FileExists(Filename)
            If fileExists = False Then
                MessageBox.Show("ไม่พบไฟล์ข้อมูล")
                Return dt
                Exit Function
            End If
            Cursor.Current = Cursors.WaitCursor
            Dim StrWer As StreamReader
            StrWer = File.OpenText(Filename)
            Dim filepath As String = Filename
            dt = CsvToTable(filepath, False)
            Cursor.Current = Cursors.Default
            dt2 = TabletoTable(dt)
            listDt.Add(dt2)
        Next
        Return listDt
    End Function
    Private Function WriteFile(ByVal listDt As List(Of DataTable))
        Dim newDT As New DataTable
        Dim dr As DataRow
        Dim i As Integer
        newDT.Columns.Add("Main Product")
        newDT.Columns.Add("Major Product")
        newDT.Columns.Add("WMS")
        newDT.Columns.Add("Shipping")
        newDT.Columns.Add("Sevices")
        newDT.Columns.Add("Minor Product")
        newDT.Columns.Add("Salesman")
        newDT.Columns.Add("Revenue")
        newDT.Columns.Add("RevenueAdvanceTransport")
        newDT.Columns.Add("RevenueAdvancePortCharge")
        For Each table As DataTable In listDt
            For i = 0 To table.Rows.Count - 1
                dr = newDT.NewRow
                dr("Main Product") = table.Rows(i).Item("Main Product")
                dr("Major Product") = table.Rows(i).Item("Major Product")
                dr("WMS") = table.Rows(i).Item("WMS")
                dr("Shipping") = table.Rows(i).Item("Shipping")
                dr("Sevices") = table.Rows(i).Item("Sevices")
                dr("Minor Product") = table.Rows(i).Item("Minor Product")
                dr("Salesman") = table.Rows(i).Item("Salesman")
                dr("Revenue") = table.Rows(i).Item("Revenue")
                dr("RevenueAdvanceTransport") = table.Rows(i).Item("RevenueAdvanceTransport")
                dr("RevenueAdvancePortCharge") = table.Rows(i).Item("RevenueAdvancePortCharge")
                newDT.Rows.Add(dr)
            Next
        Next
        Dim a As Integer = newDT.Rows.Count
        dt3 = CalToTalRevenue(newDT)
        dt4 = groupbySale(dt3)
        dt5 = CalSumRev(dt4)
        Return dt5
    End Function
    Private Function CsvToTable(ByVal filePathName As String, Optional ByVal hasHeader As Boolean = False) As DataTable
        ' Parses a csv into a datatable.
        Try
            Dim result As New DataTable
            If System.IO.File.Exists(filePathName) Then
                Dim parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(filePathName)
                parser.Delimiters = New String() {","}
                parser.HasFieldsEnclosedInQuotes = True 'use if data may contain delimiters 
                parser.TextFieldType = FileIO.FieldType.Delimited
                parser.TrimWhiteSpace = True
                Dim HeaderFlag As Boolean
                If hasHeader Then HeaderFlag = True
                While Not parser.EndOfData
                    If AddValuesToTable(parser.ReadFields, result, HeaderFlag) Then
                        HeaderFlag = False
                    Else
                        parser.Close()
                        Return Nothing
                    End If
                End While
                parser.Close()
                Return result
            Else : Return Nothing
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return Nothing
        End Try
    End Function
    Private Function AddValuesToTable(ByRef source() As String, ByVal destination As DataTable, Optional ByVal HeaderFlag As Boolean = False) As Boolean
        'Ensures a datatable can hold an array of values and then adds a new row 
        Try
            Dim existing As Integer = destination.Columns.Count
            If HeaderFlag Then
                Resolve_Duplicate_Names(source)
                For i As Integer = 0 To source.Length - existing - 1
                    destination.Columns.Add(source(i).ToString, GetType(String))
                Next i
                Return True
            End If
            For i As Integer = 0 To source.Length - existing - 1
                destination.Columns.Add("Column" & (existing + 1 + i).ToString, GetType(String))
            Next
            destination.Rows.Add(source)
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return False
        End Try
    End Function
    Private Sub Resolve_Duplicate_Names(ByRef source() As String)
        ' Resolves the possibility of duplicated names by appending "Duplicate Name" and a number at the end of any duplicates
        Dim i, n, dnum As Integer
        dnum = 1
        For n = 0 To source.Length - 1
            For i = n + 1 To source.Length - 1
                If source(i) = source(n) Then
                    source(i) = source(i) & "Duplicate Name " & dnum
                    dnum += 1
                End If
            Next
        Next
        Return
    End Sub
    Private Function TabletoTable(ByVal dt As DataTable)
        Dim dt2 As New DataTable
        Dim dr As DataRow
        dt2.Columns.Add("Main Product")
        dt2.Columns.Add("Major Product")
        dt2.Columns.Add("WMS")
        dt2.Columns.Add("Shipping")
        dt2.Columns.Add("Sevices")
        dt2.Columns.Add("Minor Product")
        dt2.Columns.Add("Salesman")
        dt2.Columns.Add("Revenue")
        dt2.Columns.Add("RevenueAdvanceTransport")
        dt2.Columns.Add("RevenueAdvancePortCharge")
        For i = 0 To dt.Rows.Count - 1
            Dim Job As String = dt.Rows(i).Item("Column2") 'Column Job No
            Dim WMSJob As String = dt.Rows(i).Item("Column3") 'Column WMSJob
            If Job.Contains("AE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "AE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PNA") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "AE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("AI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "AI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Sea Freight"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "SE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PNS") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "SE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "SI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("TSE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Land Transport"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "TSE"
                dr("Minor Product") = "Local transport"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("TSI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Land Transport"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "TSI"
                dr("Minor Product") = "Local trabsport"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("TS") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Land Transport"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "TS"
                dr("Minor Product") = "Local transport"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHAE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = "SHAE"
                dr("Sevices") = "AE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHAI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("Shipping") = "SHAI"
                dr("Sevices") = "AI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHSE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Sea Freight"
                dr("WMS") = ""
                dr("Shipping") = "SHSE"
                dr("Sevices") = "SE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHSI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Sea Freight"
                dr("WMS") = ""
                dr("Shipping") = "SHSE"
                dr("Sevices") = "SE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHTE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Land Transport"
                dr("WMS") = ""
                dr("Shipping") = "SHSE"
                dr("Sevices") = "SHE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHTI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Land Transport"
                dr("WMS") = ""
                dr("Shipping") = "SHTI"
                dr("Sevices") = "SHI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHVE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SHVE"
                dr("Sevices") = "SHE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHVI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SHVI"
                dr("Sevices") = "SHI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHPE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SHPE"
                dr("Sevices") = "SHE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHPI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SHPI"
                dr("Sevices") = "SHI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHOE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SHOE"
                dr("Sevices") = "SH"
                dr("Minor Product") = "Free Zone"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SHE"
                dr("Sevices") = "SHE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SH") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = ""
                dr("Shipping") = "SH"
                dr("Sevices") = "SHI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("HDQ") Then
                If WMSJob.Contains("CKT-IN") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Genaral WH"
                    dr("WMS") = "HDQ"
                    dr("Shipping") = ""
                    dr("Sevices") = "WI"
                    dr("Minor Product") = "WH-HDQ"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("CKT-OUT") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Genaral WH"
                    dr("WMS") = "HDQ"
                    dr("Shipping") = ""
                    dr("Sevices") = "WE"
                    dr("Minor Product") = "WH-HDQ"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                Else
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Genaral WH"
                    dr("WMS") = "HDQ"
                    dr("Shipping") = ""
                    dr("Sevices") = "HDQ"
                    dr("Minor Product") = "WH-HDQ"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                End If
            ElseIf Job.Contains("WE") Then
                If WMSJob.Contains("LKB-OUT") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "LKB"
                    dr("Shipping") = ""
                    dr("Sevices") = "WE"
                    dr("Minor Product") = "WH-LKB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("EPN-ONLI") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Genaral WH"
                    dr("WMS") = "HDQ"
                    dr("Shipping") = ""
                    dr("Sevices") = "WE"
                    dr("Minor Product") = "WH-HDQ"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                Else
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WE"
                    dr("Shipping") = ""
                    dr("Sevices") = "WE"
                    dr("Minor Product") = "WH-LKB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                End If
            ElseIf Job.Contains("SVI-WI") Then
                If WMSJob.Contains("SVI-19-0") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "MTL"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVI"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SVI") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WI"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVI"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-107-IN") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "MTL"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVI"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-109-IN") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "EAS"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVI"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-110-IN") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "MTL"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVI"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                Else
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WI"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVI"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                End If
            ElseIf Job.Contains("SVO-WE") Then
                If WMSJob.Contains("SVO-19-0") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "MTL"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVO"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SVO") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WO"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVO"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-107-OUT") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "MTL"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVO"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-109-OUT") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "EAS"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVO"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-110-OUT") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "MTL"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVO"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                Else
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WO"
                    dr("Shipping") = ""
                    dr("Sevices") = "SVO"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                End If
            ElseIf Job.Contains("PE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Packing"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Packing"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Packing"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PL"
                dr("Minor Product") = "Local"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("FE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Fumigate"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "FE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("FI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Fumigate"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "FI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("FL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Fumigate"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "FL"
                dr("Minor Product") = "Local"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PHE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Purchase"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PHE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PHI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Purchase"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PHI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PHL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Purchase"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PHL"
                dr("Minor Product") = "Local"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PSE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "PersonClear"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PSE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PSI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "PersonClear"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PSI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PSL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "PersonClear"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PSL"
                dr("Minor Product") = "Local"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("RE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Rental"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "RE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("RI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Rental"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "RI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("RL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Rental"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "RL"
                dr("Minor Product") = "Local"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("CFE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "ChaneForklift"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "CFE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("CFI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "ChaneForklift"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "CFI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("CFL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "ChaneForklift"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "CFL"
                dr("Minor Product") = "Local"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("FSE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "FeeStateOfReltns"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "FSE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("FSI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "FeeStateOfReltns"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "FSI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("FSL") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "FeeStateOfReltns"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "FSL"
                dr("Minor Product") = "Land"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PME") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Permit"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PME"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PMI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Permit"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PMI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PML") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "Permit"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "PML"
                dr("Minor Product") = "Land"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("INT") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other service"
                dr("Major Product") = "PermitReq"
                dr("WMS") = ""
                dr("Shipping") = ""
                dr("Sevices") = "INT"
                dr("Minor Product") = "Other"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            End If
        Next
        Return dt2
    End Function
    Private Function CalToTalRevenue(ByVal newDT As DataTable)
        Dim dt3 As New DataTable
        Dim dr As DataRow
        Dim i As Integer
        Dim arrayMain() As String = {"Freight", "Customer Broker", "Warehouse", "Other service"}
        Dim arrayMajor() As String = {"Air Freight", "Sea Freight", "Land Transport", "Formality", "Genaral WH", "Free Zone WH", "Packing", "Fumigate", "Purchase", "PersonClear", "Rental", "ChaneForklift", "FeeStateOfReltns", "PermitReq"}
        Dim arrayService() As String = {"AE", "AI", "SE", "SI", "TSE", "TSI", "TS", "SHE", "SHI", "WI", "WE", "HD", "SVI", "SVO", "PE", "PI", "PL", "FE", "FI", "FL", "PHE", "PHI", "PHL", "PSE", "PSI", "PSL", "RE", "RI", "RL", "CFE", "CFI", "CFL", "FSE", "FSI", "FSL", "PME", "PMI", "PML", "IN"}
        dt3.Columns.Add("Main Product")
        dt3.Columns.Add("Major Product")
        dt3.Columns.Add("WMS")
        dt3.Columns.Add("Shipping")
        dt3.Columns.Add("Sevices")
        dt3.Columns.Add("Minor Product")
        dt3.Columns.Add("Salesman")
        dt3.Columns.Add("ToTalRevenue+TS")
        For Each main As String In arrayMain
            For i = 0 To newDT.Rows.Count - 1
                Dim stMain As String = newDT.Rows(i).Item("Main Product")
                If stMain = main Then
                    Dim stMajor As String = newDT.Rows(i).Item("Major Product")
                    For Each major As String In arrayMajor
                        If stMajor = major Then
                            Dim service As String = newDT.Rows(i).Item("Sevices")
                            For Each stService As String In arrayService
                                If service = stService Then
                                    Dim Revenue As Double = newDT.Rows(i).Item("Revenue")
                                    Dim RevenueAdvanceTransport As Double = newDT.Rows(i).Item("RevenueAdvanceTransport")
                                    Dim RevenueAdvancePortCharge As Double = newDT.Rows(i).Item("RevenueAdvancePortCharge")
                                    Dim ToTalRevenue As Double = Revenue + RevenueAdvanceTransport + RevenueAdvancePortCharge
                                    dr = dt3.NewRow
                                    dr("Main Product") = newDT.Rows(i).Item("Main Product")
                                    dr("Major Product") = newDT.Rows(i).Item("Major Product")
                                    dr("WMS") = newDT.Rows(i).Item("WMS")
                                    dr("Shipping") = newDT.Rows(i).Item("Shipping")
                                    dr("Sevices") = newDT.Rows(i).Item("Sevices")
                                    dr("Minor Product") = newDT.Rows(i).Item("Minor Product")
                                    dr("Salesman") = newDT.Rows(i).Item("Salesman")
                                    dr("ToTalRevenue+TS") = ToTalRevenue
                                    dt3.Rows.Add(dr)
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Next
        Dim a As Integer = dt3.Rows.Count
        Return dt3
    End Function
    Private Function groupbySale(ByVal dt3 As DataTable)
        Dim dt4 As New DataTable
        Dim listSale As New List(Of String)
        Dim i As Integer
        Dim j As Integer
        Dim bool As Boolean = False
        Dim dr As DataRow
        dt4.Columns.Add("Main Product")
        dt4.Columns.Add("Major Product")
        dt4.Columns.Add("WMS")
        dt4.Columns.Add("Shipping")
        dt4.Columns.Add("Sevices")
        dt4.Columns.Add("Minor Product")
        dt4.Columns.Add("Salesman")
        dt4.Columns.Add("ToTalRevenue+TS")
        For i = 0 To dt3.Rows.Count - 1
            Dim stSale As String = dt3.Rows(i).Item("Salesman")
            If listSale.Count = 0 Then
                listSale.Add(stSale)
            Else
                For Each sale As String In listSale
                    If sale = stSale Then
                        bool = True
                    End If
                Next
                If bool = True Then
                    bool = False
                Else
                    listSale.Add(stSale)
                End If
            End If
        Next
        For Each sale As String In listSale
            For j = 0 To dt3.Rows.Count - 1
                Dim stSale As String = dt3.Rows(j).Item("Salesman")
                If stSale = sale Then
                    dr = dt4.NewRow
                    dr("Main Product") = dt3.Rows(j).Item("Main Product")
                    dr("Major Product") = dt3.Rows(j).Item("Major Product")
                    dr("WMS") = dt3.Rows(j).Item("WMS")
                    dr("Shipping") = dt3.Rows(j).Item("Shipping")
                    dr("Sevices") = dt3.Rows(j).Item("Sevices")
                    dr("Minor Product") = dt3.Rows(j).Item("Minor Product")
                    dr("Salesman") = dt3.Rows(j).Item("Salesman")
                    dr("ToTalRevenue+TS") = dt3.Rows(j).Item("ToTalRevenue+TS")
                    dt4.Rows.Add(dr)
                End If
            Next
        Next
        Dim a As Integer = dt4.Rows.Count
        Return dt4
    End Function
    Private Function CalSumRev(ByVal dt4 As DataTable)
        Dim dt5 As New DataTable
        Dim listSale As New List(Of String)
        Dim i As Integer
        Dim j As Integer
        Dim dr As DataRow
        Dim bool As Boolean = False
        Dim listTotalRev As New List(Of Double)
        Dim arrayMain() As String = {"Freight", "Customer Broker", "Warehouse", "Other service"}
        Dim arrayMajor() As String = {"Air Freight", "Sea Freight", "Land Transport", "Formality", "Genaral WH", "Free Zone WH", "Packing", "Fumigate", "Purchase", "PersonClear", "Rental", "ChaneForklift", "FeeStateOfReltns", "PermitReq"}
        Dim arrayService() As String = {"AE", "AI", "SE", "SI", "TSE", "TSI", "TS", "SHE", "SHI", "WI", "WE", "HD", "SVI", "SVO", "PE", "PI", "PL", "FE", "FI", "FL", "PHE", "PHI", "PHL", "PSE", "PSI", "PSL", "RE", "RI", "RL", "CFE", "CFI", "CFL", "FSE", "FSI", "FSL", "PME", "PMI", "PML", "IN"}
        dt5.Columns.Add("Main Product")
        dt5.Columns.Add("Major Product")
        dt5.Columns.Add("WMS")
        dt5.Columns.Add("Shipping")
        dt5.Columns.Add("Sevices")
        dt5.Columns.Add("Minor Product")
        'dt5.Columns.Add("Salesman")
        'dt5.Columns.Add("ToTalRevenue+TS")
        For i = 0 To dt3.Rows.Count - 1
            Dim stSale As String = dt4.Rows(i).Item("Salesman")
            If listSale.Count = 0 Then
                listSale.Add(stSale)
            Else
                For Each sale As String In listSale
                    If stSale = "" Then
                        stSale = "rooting"
                    End If
                    If sale = stSale Then
                        bool = True
                    End If
                Next
                If bool = True Then
                    bool = False
                Else
                        listSale.Add(stSale)
                End If
            End If
        Next
        For Each salename As String In listSale
            dt5.Columns.Add(salename)
        Next
        For Each salename As String In listSale
            For Each main As String In arrayMain
                    For Each major As String In arrayMajor
                    For Each service As String In arrayService
                        Dim WMS As String
                        Dim Shipping As String
                        Dim minor As String
                        Dim bool2 As Boolean = False
                        For j = 0 To dt4.Rows.Count - 1
                            Dim stMain As String = dt4.Rows(j).Item("Main Product")
                            Dim stMajor As String = dt4.Rows(j).Item("Major Product")
                            Dim stService As String = dt4.Rows(j).Item("Sevices")
                            Dim stSalename As String = dt4.Rows(j).Item("Salesman")
                            If stSalename = salename Then
                                If stMain = main Then
                                    If stMajor = major Then
                                        If stService = service Then
                                            WMS = dt4.Rows(j).Item("WMS")
                                            Shipping = dt4.Rows(j).Item("Shipping")
                                            minor = dt4.Rows(j).Item("Minor Product")
                                            Dim TotalRev As Double = dt4.Rows(j).Item("ToTalRevenue+TS")
                                            listTotalRev.Add(TotalRev)
                                            bool2 = True
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        Dim k As Integer
                        Dim CalRev As Double
                        For k = 0 To listTotalRev.Count - 1
                            CalRev += listTotalRev(k)
                        Next
                        If bool2 = True Then
                            dr = dt5.NewRow
                            dr("Main Product") = main
                            dr("Major Product") = major
                            dr("WMS") = WMS
                            dr("Shipping") = Shipping
                            dr("Sevices") = service
                            dr("Minor Product") = minor
                            dr(salename) = CalRev
                            dt5.Rows.Add(dr)
                            listTotalRev.Clear()
                            bool2 = False
                        End If
                    Next
                Next
            Next
        Next
        Dim a As Integer = dt5.Rows.Count
        Return dt5
    End Function
    Private Function groupData(ByVal dt5 As DataTable)
        Dim dt6 As New DataTable
        Dim listSale As New List(Of String)
        Dim i As Integer
        Dim j As Integer
        Dim arrayDr() As DataRow
        Dim bool As Boolean = False
        Dim listTotalRev As New List(Of Double)
        Dim list As New List(Of String)
        Dim arrayMain() As String = {"Freight", "Customer Broker", "Warehouse", "Other service"}
        Dim arrayMajor() As String = {"Air Freight", "Sea Freight", "Land Transport", "Formality", "Genaral WH", "Free Zone WH", "Packing", "Fumigate", "Purchase", "PersonClear", "Rental", "ChaneForklift", "FeeStateOfReltns", "PermitReq"}
        Dim arrayService() As String = {"AE", "AI", "SE", "SI", "TSE", "TSI", "TS", "SHE", "SHI", "WI", "WE", "HD", "SVI", "SVO", "PE", "PI", "PL", "FE", "FI", "FL", "PHE", "PHI", "PHL", "PSE", "PSI", "PSL", "RE", "RI", "RL", "CFE", "CFI", "CFL", "FSE", "FSI", "FSL", "PME", "PMI", "PML", "IN"}
        For i = 0 To dt3.Rows.Count - 1
            Dim stSale As String = dt4.Rows(i).Item("Salesman")
            If listSale.Count = 0 Then
                listSale.Add(stSale)
            Else
                For Each sale As String In listSale
                    If sale = stSale Then
                        bool = True
                    End If
                Next
                If bool = True Then
                    bool = False
                Else
                    listSale.Add(stSale)
                End If
            End If
        Next
        dt5.Columns.Add("Main Product")
        dt5.Columns.Add("Major Product")
        dt5.Columns.Add("WMS")
        dt5.Columns.Add("Shipping")
        dt5.Columns.Add("Sevices")
        dt5.Columns.Add("Minor Product")
        For Each salename As String In listSale
            dt5.Columns.Add(salename)
        Next
        For Each main As String In arrayMain
            For Each major As String In arrayMajor
                For Each service As String In arrayService
                    For j = 0 To dt5.Rows.Count - 1
                        Dim stMain As String = dt5.Rows(j).Item("Main Product")
                        Dim stMajor As String = dt5.Rows(j).Item("Major Product")
                        Dim stService As String = dt5.Rows(j).Item("Services")
                        If stMain = main Then
                            If stMajor = major Then
                                If stService = service Then
                                    arrayDr = dt5.Select(j)
                                End If
                            End If
                        End If
                    Next
                Next
            Next
        Next
        Return dt6
    End Function
End Class