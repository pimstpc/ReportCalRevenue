Imports System.IO
Imports System.Net.Mail

Public Class Form1
    Dim OpenFileDialog As New OpenFileDialog()
    Dim status As Boolean
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim dt3 As DataTable
    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        OpenFileDialog.Filter = "CSV File(*.csv)|*.csv"
        OpenFileDialog.Multiselect = True
        If OpenFileDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

            Dim fi As New FileInfo(OpenFileDialog.FileName)
            Me.txtfile.Text = OpenFileDialog.FileName
        Else
            Exit Sub
        End If
    End Sub
    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        Dim dt As New DataTable
        dt = ReadFile()
        status = WriteFile(dt)
        If (status = True) Then
            MessageBox.Show("Complete")
        Else
            MessageBox.Show("Not complete")
        End If
    End Sub
    Private Function ReadFile()
        Dim fileExists As Boolean
        Dim dt As New DataTable()

        fileExists = My.Computer.FileSystem.FileExists(txtfile.Text.Trim)

        'check file is csv only

        If fileExists = False Then
            MessageBox.Show("ไม่พบไฟล์ข้อมูล")
            Return dt
            Exit Function
        End If

        Cursor.Current = Cursors.WaitCursor

        Dim StrWer As StreamReader

        StrWer = File.OpenText(txtfile.Text.Trim)

        Dim filepath As String = txtfile.Text.Trim

        dt = CsvToTable(filepath, False)

        Cursor.Current = Cursors.Default

        Return dt
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
    Private Function WriteFile(ByVal dt As DataTable)
        dt2 = genTable(dt)
        dt3 = Cal(dt, dt2)
        Dim Month As String = Format(Now(), "MM")
        Dim Year As String = Format(Now(), "yy")
        Dim filepath As String = "D:\AE JOB INCOMPLETE " + " " + Month + "-" + Year + ".csv"
        status = TableToCSV(dt2, filepath)
        If (status = False) Then
            Return False
            Exit Function
        End If
        status = sendMail(filepath)
        If (status = False) Then
            Return False
            Exit Function
        End If
        Return True
    End Function
    Private Function TableToCSV(ByVal sourceTable As DataTable, ByVal filePathName As String, Optional ByVal HasHeader As Boolean = True) As Boolean
        'Writes a datatable back into a csv 
        Try
            Dim sb As New System.Text.StringBuilder
            If HasHeader Then
                Dim nameArray(200) As Object
                Dim i As Integer = 0
                For Each col As DataColumn In sourceTable.Columns
                    nameArray(i) = CType(col.ColumnName, Object)
                    i += 1
                Next col
                ReDim Preserve nameArray(i - 1)
                sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(nameArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o))))
            End If
            For Each dr As DataRow In sourceTable.Rows
                sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(dr.ItemArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o.ToString))))
            Next
            System.IO.File.WriteAllText(filePathName, sb.ToString)
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return False
        End Try
    End Function
    Private Function sendMail(ByVal filepath As String)
        Try

            Dim myMail = New MailMessage()
            myMail.From = New MailAddress("Auto report<aa@gmail.com>")

            Dim name As String = filepath.Replace(".csv", "")
            name = name.Replace("D:\", "")
            myMail.Subject = "Auto report : " + name
            myMail.To.Add(New MailAddress("bb@gmail.com<bb@gmail.com>"))
            myMail.IsBodyHtml = True
            myMail.BodyEncoding = System.Text.Encoding.UTF8
            myMail.Body = "This is auto report: " + name + ". Please see attached file"

            myMail.Attachments.Add(New Attachment(filepath))

            Dim smtpClient = New SmtpClient()
            smtpClient.Send(myMail)

            smtpClient.Dispose()
            myMail.Dispose()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function genTable(ByVal dt As DataTable)
        Dim i As Integer
        Dim dt2 As New DataTable
        Dim dr As DataRow
        dt2.Columns.Add("Main Product")
        dt2.Columns.Add("Major Product")
        dt2.Columns.Add("WMS")
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
                dr("SERVICES") = "AI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("AE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Air Freight"
                dr("WMS") = ""
                dr("SERVICES") = "AE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Sea Freight"
                dr("WMS") = ""
                dr("SERVICES") = "SI"
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
                dr("SERVICES") = "SE"
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
                dr("SERVICES") = "AE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("PNS") Then
                dr = dt2.NewRow
                dr("Main Product") = "Freight"
                dr("Major Product") = "Sea Freight"
                dr("WMS") = ""
                dr("SERVICES") = "SE"
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
                dr("WMS") = "SHAI"
                dr("SERVICES") = "AI"
                dr("Minor Product") = "Import"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHAE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Air Freight"
                dr("WMS") = "SHAE"
                dr("SERVICES") = "AE"
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
                dr("WMS") = "SHSI"
                dr("SERVICES") = "SI"
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
                dr("WMS") = "SHSE"
                dr("SERVICES") = "SE"
                dr("Minor Product") = "Export"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHOE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = "SHOE"
                dr("SERVICES") = "SH"
                dr("Minor Product") = "Free Zone"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHPE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = "SHPE"
                dr("SERVICES") = "SH"
                dr("Minor Product") = ""
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHTI") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = "SHTI"
                dr("SERVICES") = "SH"
                dr("Minor Product") = "Local Transport"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("SHTE") Then
                dr = dt2.NewRow
                dr("Main Product") = "Customer Broker"
                dr("Major Product") = "Formality"
                dr("WMS") = "SHTE"
                dr("SERVICES") = "SH"
                dr("Minor Product") = "Local Transport"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("HDQ") Then
                If WMSJob.Contains("CKT-IN") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "General WH"
                    dr("WMS") = "HDQ"
                    dr("SERVICES") = "WI"
                    dr("Minor Product") = "WH-HDQ"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("CKT-OUT") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "General WH"
                    dr("WMS") = "HDQ"
                    dr("SERVICES") = "WE"
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
                    dr("WMS") = "HD"
                    dr("SERVICES") = "HD"
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
                    dr("SERVICES") = "WE"
                    dr("Minor Product") = "WH-LKB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("EPN-ONLI") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "HDQ"
                    dr("SERVICES") = "WE"
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
                    dr("SERVICES") = "WE"
                    dr("Minor Product") = "WH-LKB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                End If
            ElseIf Job.Contains("WI") Then
                If WMSJob.Contains("LKB-IN") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "LKB"
                    dr("SERVICES") = "WI"
                    dr("Minor Product") = "WH-LKB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("EPN-ONLI") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "HDQ"
                    dr("SERVICES") = "WI"
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
                    dr("WMS") = "WI"
                    dr("SERVICES") = "WI"
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
                    dr("WMS") = "WI-MTL"
                    dr("SERVICES") = "SV"
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
                    dr("SERVICES") = "SV"
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
                    dr("WMS") = "WI-107"
                    dr("SERVICES") = "SV"
                    dr("Minor Product") = "SBIA-107"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-109") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WI-109"
                    dr("SERVICES") = "SV"
                    dr("Minor Product") = "SBIA-109"
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
                    dr("SERVICES") = "SV"
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
                    dr("WMS") = "WO-MTL"
                    dr("SERVICES") = "SV"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-107") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WE-107"
                    dr("SERVICES") = "SV"
                    dr("Minor Product") = "SBIA-107"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                ElseIf WMSJob.Contains("SBIA-109") Then
                    dr = dt2.NewRow
                    dr("Main Product") = "Warehouse"
                    dr("Major Product") = "Free Zone WH"
                    dr("WMS") = "WE-109"
                    dr("SERVICES") = "SV"
                    dr("Minor Product") = "SBIA-109"
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
                    dr("SERVICES") = "SV"
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
                    dr("WMS") = "WE"
                    dr("SERVICES") = "SV"
                    dr("Minor Product") = "SBIA/SBIA-BKK HUB"
                    dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                    dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                    dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                    dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                    dt2.Rows.Add(dr)
                End If
            ElseIf Job.Contains("TS") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other Service"
                dr("Major Product") = "Transportation"
                dr("WMS") = "TS"
                dr("SERVICES") = "TS"
                dr("Minor Product") = "HDQ"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            ElseIf Job.Contains("INT") Then
                dr = dt2.NewRow
                dr("Main Product") = "Other Service"
                dr("Major Product") = "Misc1"
                dr("WMS") = "IN"
                dr("SERVICES") = "IN"
                dr("Minor Product") = "OTHER"
                dr("Salesman") = dt.Rows(i).Item("Column9") 'Column Slaesman
                dr("Revenue") = dt.Rows(i).Item("Column40") 'Column Revenue
                dr("RevenueAdvanceTransport") = dt.Rows(i).Item("Column68") 'Column Revenue Advance Transport
                dr("RevenueAdvancePortCharge") = dt.Rows(i).Item("Column70") 'Column Revenue Advance Port Charge
                dt2.Rows.Add(dr)
            End If
        Next
        Return dt2
    End Function
    Private Function Cal(ByVal dt As DataTable, ByVal dt2 As DataTable)
        Dim dt3 As New DataTable
        Dim listSalesname As New List(Of String)
        Dim i As Integer
        Dim bool As Boolean = False
        Dim j As Integer
        Dim listTotalRev As New List(Of Double)
        For i = 0 To dt.Rows.Count
            Dim stSalesname As String = dt.Rows(i).Item("Column9") 'Column Salesman
            If listSalesname.Count = 0 Then
                listSalesname.Add(stSalesname)
            Else
                For Each salesname As String In listSalesname
                    If salesname = stSalesname Then
                        bool = True
                    End If
                    If bool = False Then
                        listSalesname.Add(stSalesname)
                    Else
                        bool = False
                    End If
                Next
            End If
        Next
        For Each salesname As String In listSalesname
            For j = 0 To dt2.Rows.Count - 1
                Dim stSalesname As String = dt2.Rows(j).Item("Salesman")
                If stSalesname = salesname Then
                    Dim Rev As Double = dt2.Rows(j).Item("Revenue")
                    Dim RevTrans As Double = dt2.Rows(j).Item("RevenueAdvanceTransport")
                    Dim RevAdvance As Double = dt2.Rows(j).Item("RevenueAdvancePortCharge")
                    Dim SumRev As Double = Rev + RevTrans + RevAdvance
                    listTotalRev.Add(SumRev)
                End If
            Next
        Next
        Return dt3
    End Function

    Private Function CalToTalRevenue(ByVal dt2 As DataTable)
        Dim dt3 As New DataTable
        Dim listSalesname As New List(Of String)
        Dim i As Integer
        Dim bool As Boolean = False
        Dim j As Integer
        Dim listTotalRev As New List(Of Double)
        Dim dr As DataRow
        dt2.Columns.Add("Main Product")
        dt2.Columns.Add("Major Product")
        dt2.Columns.Add("WMS")
        dt2.Columns.Add("Sevices")
        dt2.Columns.Add("Minor Product")
        dt2.Columns.Add("Salesman")
        dt2.Columns.Add("ToTalRevenue+TS")
        For i = 0 To dt.Rows.Count
            Dim stSalesname As String = dt.Rows(i).Item("Column9") 'Column Salesman
            If listSalesname.Count = 0 Then
                listSalesname.Add(stSalesname)
            Else
                For Each salesname As String In listSalesname
                    If salesname = stSalesname Then
                        bool = True
                    End If
                    If bool = False Then
                        listSalesname.Add(stSalesname)
                    Else
                        bool = False
                    End If
                Next
            End If
        Next
        For Each salesname As String In listSalesname
            For j = 0 To dt2.Rows.Count - 1
                Dim stSalesname As String = dt2.Rows(j).Item("Salesman")
                If stSalesname = salesname Then
                    Dim Rev As Double = dt2.Rows(j).Item("Revenue")
                    Dim RevTrans As Double = dt2.Rows(j).Item("RevenueAdvanceTransport")
                    Dim RevAdvance As Double = dt2.Rows(j).Item("RevenueAdvancePortCharge")
                    Dim SumRev As Double = Rev + RevTrans + RevAdvance
                    'listTotalRev.Add(SumRev)
                    dr = dt3.NewRow
                    dr("Main Product") = dt2.Rows(j).Item("Main Product")
                    dr("Major Product") = dt2.Rows(j).Item("Major Product")
                    dr("WMS") = dt2.Rows(j).Item("WMS")
                    dr("SERVICES") = dt2.Rows(j).Item("SERVICES")
                    dr("Minor Product") = dt2.Rows(j).Item("Minor Product")
                    dr("Salesman") = stSalesname
                    dr("ToTalRevenue+TS") = SumRev
                    dt3.Rows.Add(dr)
                End If
            Next
        Next
        Return dt3
    End Function
    Private Function GroupData(ByVal dt3 As DataTable)
        Dim dt4 As New DataTable
        Dim i As Integer
        For i = 0 To dt3.Rows.Count - 1
            If dt3.Rows(i).Item("Main Product") = "Freight" Then
                If dt3.Rows(i).Item("Sevices") = "AE" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "PNA" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "AI" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "SE" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "PNS" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "SI" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "TSE" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "TSI" Then
                ElseIf dt3.Rows(i).Item("Sevices") = "TS" Then
                End If
            End If
        Next
        Return dt4
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form2.ShowDialog()
    End Sub
End Class
