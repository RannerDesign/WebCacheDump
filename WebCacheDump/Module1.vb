Imports Microsoft.Office.Interop

Module Module1

    Sub Main()
        '   WebCacheDump
        '   09.10.2017 RB
        '   Dump content of WebCacheV01.dat into Excel workbook suitable for further analysis
        '
        Dim debugMode As Boolean = True
        Dim logLineStart As String = "++++" + vbTab + "WebCacheDump: "
        Console.WriteLine(logLineStart + "Start WebCacheDump")

        Dim ese As New ESE_Engine
        Dim cacheFile As String
        Dim cacheDir As String
        Dim iret As Integer

        '   Prepare Excel workbook for output of dump results with first sheet for the database tables
        Dim xlApp As New Excel.Application
        xlApp.Visible = True
        Dim xlBook As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlSheet1 As Excel.Worksheet = xlBook.Sheets(1)
        xlSheet1.Name = "Tables"
        xlSheet1.Cells(1, 1).value = "ID"
        xlSheet1.Cells(1, 2).value = "Table"
        xlSheet1.Cells(1, 3).value = "Records"
        xlSheet1.Cells(1, 4).value = "Tab"

        '   Location of WebCache database depending on test or productive environment
        If debugMode Then
            cacheDir = Environ$("LOCALAPPDATA") + "\Microsoft\Windows\WebCacheTest"
            cacheFile = cacheDir + "\WebCacheV01.dat"
        Else
            cacheDir = Environ$("LOCALAPPDATA") + "\Microsoft\Windows\WebCache"
            cacheFile = cacheDir + "\WebCacheV01.dat"

            '   Stop WebCache task
            Console.WriteLine(logLineStart + "Stop WebCache task")
            iret = ESE_Tools.RunCmd("schtasks", "/end /tn ""Microsoft\Windows\Wininet\CacheTask""", "")

            '   Check if database is locked
            Dim LockListe As List(Of Process)
            LockListe = LockTools.FindLockers(cacheFile)

            '   Eventually kill other processes
            If LockListe.Count > 0 Then
                Console.WriteLine(logLineStart + "The following processes need to be terminated in order to access the database:")
            End If
            Dim PP As Process
            For Each PP In LockListe
                Console.WriteLine(logLineStart + "Process to be killed: " + PP.ProcessName + " (" + CStr(PP.Id) + ")")
                Process.GetProcessById(PP.Id).Kill()
            Next
        End If

        '   Recover database
        Console.WriteLine(logLineStart + "start database recovery ...")
        iret = ESE_Tools.RunCmd("esentutl", "/r V01 /d", cacheDir)
        Console.WriteLine(logLineStart + "start database compression ...")
        iret = ESE_Tools.RunCmd("esentutl", "/d WebCacheV01.dat", cacheDir)

        '   Open database
        iret = ese.DBopen(cacheFile)
        Dim errorFree As Boolean = True

        '   Prepare list of tables
        Dim dtSysObj As DataTable
        Dim dr As DataRow
        Dim ContainerID As Long = 0
        Dim iRowTables As Long = 1
        Console.WriteLine(logLineStart + "prepare list of tables from MSysObjects ...")
        Try
            dtSysObj = ese.GetTable("MSysObjects", False)
            For Each dr In dtSysObj.Select("[Type]=1")
                iRowTables = iRowTables + 1
                xlSheet1.Cells(iRowTables, 1).value = dr("Id")
                xlSheet1.Cells(iRowTables, 2).value = dr("Name")
                xlSheet1.Cells(iRowTables, 3).value = ese.RecordCount(dr("Name"))
                If dr("Name") = "Containers" Then ContainerID = dr("Id")
            Next
        Catch ex As Exception
            errorFree = False
            MsgBox("Error reading MSysObjects:" + vbCrLf + ex.Message)
        End Try

        '   Prepare Container data
        Dim dtContainer As DataTable
        Dim xlSheet2 As Excel.Worksheet = xlBook.Sheets.Add(, xlBook.Worksheets(xlBook.Worksheets.Count))
        Console.WriteLine(logLineStart + "prepare Container data from table Containers ...")
        '   Container columns
        xlSheet2.Name = "Containers"
        Dim jC As Long = 0
        For Each dr In dtSysObj.Select("[Type]=2 AND [ObjidTable]=" + CStr(ContainerID))
            jC = jC + 1
            xlSheet2.Cells(1, jC).value = dr("ColtypOrPgnoFDP")
            xlSheet2.Cells(2, jC).value = dr("Name")
        Next
        '   Container contents
        Dim iRowContainer As Long = 2
        Dim jX As Long
        Try
            dtContainer = ese.GetTable("Containers", False)
            For Each dr In dtContainer.Select()
                iRowContainer = iRowContainer + 1
                For jX = 1 To jC
                    xlSheet2.Cells(iRowContainer, jX).value = dr(CStr(xlSheet2.Cells(2, jX).value))
                Next
            Next
        Catch ex As Exception
            errorFree = False
            MsgBox("Error reading Containers:" + vbCrLf + ex.Message)
        End Try

        '   Loop thru all tables and shuffle contents to appropriate tabs of the Excel sheet
        Dim iTable As Long = 2
        Dim TableName As String
        Dim TableID As Long
        Dim TableRows As Long
        Dim TableSheet As String
        Dim TableParts() As String
        Dim PartNumber As String
        Dim xlSheet3 As Excel.Worksheet
        Dim UsedRow3 As Long
        Dim jY As Long
        Dim columsMismatch As Boolean

        Do While iTable <= iRowTables And errorFree
            TableName = xlSheet1.Cells(iTable, 2).value
            TableID = xlSheet1.Cells(iTable, 1).value
            TableRows = xlSheet1.Cells(iTable, 3).value
            If TableName.Substring(0, 4) <> "MSys" And TableName <> "Containers" And xlSheet1.Cells(iTable, 3).value > 0 Then
                Console.WriteLine(logLineStart + "load table " + TableName + " with " + CStr(TableRows) + " rows")
                TableParts = TableName.Split("_")
                If TableParts.Length = 1 Then
                    '   Table without sequence number will be stored in own sheet
                    PartNumber = ""
                    TableSheet = TableName
                Else
                    '   Table with sequence number will be stored together in one sheet (without number)
                    PartNumber = TableParts(1)
                    TableSheet = TableParts(0)
                    If TableName.Substring(0, 10) = "Container_" Then
                        '   Table series within Container will be stored in one sheet (without number) depending on container content
                        For Each dr In dtContainer.Select("[ContainerID]=" + PartNumber)
                            TableSheet = dr("Name")
                        Next
                        '   History containers will be packed together
                        If TableSheet.Substring(0, 6) = "MSHist" Then
                            TableSheet = TableSheet.Substring(0, 6)
                        End If
                        TableSheet = "C_" + TableSheet
                    End If
                End If
                '   Check if sheet already exists
                If XLTableExists(xlBook, TableSheet) Then
                    '   If so, position to end of existing tab
                    xlSheet3 = xlBook.Sheets(TableSheet)
                    UsedRow3 = xlSheet3.UsedRange.Rows.Count
                    '   Check if columns are same
                    columsMismatch = False
                    jY = 1
                    For Each dr In dtSysObj.Select("[Type]=2 AND [ObjidTable]=" + CStr(TableID))
                        jY = jY + 1
                        If xlSheet3.Cells(1, jY).value <> dr("ColtypOrPgnoFDP") Then columsMismatch = True
                        If xlSheet3.Cells(2, jY).value <> dr("Name") Then columsMismatch = True
                    Next
                    If columsMismatch Then MsgBox("Warning:" + vbCrLf + "Table " + TableName + " structure not matching sheet " + TableSheet)
                Else
                    '   If not, create new tab
                    xlSheet3 = xlBook.Sheets.Add(, xlBook.Worksheets(xlBook.Worksheets.Count))
                    xlSheet3.Name = TableSheet
                    '   Fill first 2 rows
                    jY = 1
                    xlSheet3.Cells(2, jY).value = "Nr"
                    For Each dr In dtSysObj.Select("[Type]=2 AND [ObjidTable]=" + CStr(TableID))
                        jY = jY + 1
                        xlSheet3.Cells(1, jY).value = dr("ColtypOrPgnoFDP")
                        xlSheet3.Cells(2, jY).value = dr("Name")
                    Next
                    UsedRow3 = 2
                End If
                '   Load data from ESE database table to Excel tab
                Dim dtDataTable As DataTable
                Try
                    dtDataTable = ese.GetTable(TableName, False)
                    For Each dr In dtDataTable.Select()
                        UsedRow3 = UsedRow3 + 1
                        xlSheet3.Cells(UsedRow3, 1).value = PartNumber
                        For jX = 2 To xlSheet3.UsedRange.Columns.Count
                            xlSheet3.Cells(UsedRow3, jX).value = dr(CStr(xlSheet3.Cells(2, jX).value))
                        Next
                    Next
                Catch ex As Exception
                    MsgBox("Error reading " + TableName + ":" + vbCrLf + ex.Message)
                End Try
                xlSheet1.Cells(iTable, 4).value = TableSheet
            End If
            iTable = iTable + 1
        Loop
        '   Close database
        Console.WriteLine(logLineStart + "all tables loaded")
        iret = ese.DBclose(cacheFile)
        '   Start WebCache task again if in productive environment
        If Not debugMode Then
            Console.WriteLine(logLineStart + "Start WebCache task")
            iret = ESE_Tools.RunCmd("schtasks", "/run /tn ""Microsoft\Windows\Wininet\CacheTask""", "")
        End If
        '   Beautify the resulting Excel workbook
        Dim nRows As Long
        Dim nCols As Long
        For Each xlSheet3 In xlBook.Sheets()
            With xlSheet3.UsedRange
                nRows = .Rows.Count
                nCols = .Columns.Count
            End With
            Console.WriteLine(logLineStart + "beautifying sheet " + xlSheet3.Name + " with " + CStr(nRows) + " rows")
            If xlSheet3.Name = "Tables" Then
                '   First row center and bold for sheet "Tables"
                With xlSheet3.Rows(1)
                    .Font.Bold = True
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                End With
                '   Freeze pane with first row
                Try
                    xlSheet3.Activate()
                    With xlApp.ActiveWindow
                        .SplitColumn = 0
                        .SplitRow = 1
                        .FreezePanes = True
                    End With
                Catch ex As Exception
                    Console.WriteLine(logLineStart + "error freezing panes > " + ex.Message)
                End Try
            Else
                '   Second row center and bold for sheets besides "Tables"
                With xlSheet3.Rows(2)
                    .Font.Bold = True
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                End With
                xlSheet3.Rows(1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                '   Format Columns
                Try
                    For jC = 1 To nCols
                        If xlSheet3.Cells(1, jC).value > 13 Then xlSheet3.Columns(jC).NumberFormat = "0"
                    Next
                Catch ex As Exception
                    Console.WriteLine(logLineStart + "error changing NumberFormat > " + ex.Message)
                End Try
                '   Freeze pane with first 2 rows
                Try
                    xlSheet3.Activate()
                    With xlApp.ActiveWindow
                        .SplitColumn = 0
                        .SplitRow = 2
                        .FreezePanes = True
                    End With
                Catch ex As Exception
                    Console.WriteLine(logLineStart + "error freezing panes > " + ex.Message)
                End Try
            End If
            '   Adjust all column widths
            xlSheet3.Columns.EntireColumn.AutoFit()
        Next

        MsgBox("End WebCacheDump")
    End Sub

    Private Function XLTableExists(workbook As Excel.Workbook, tablename As String) As Boolean
        For Each t In workbook.Sheets
            If t.name = tablename Then
                XLTableExists = True
                Exit Function
            End If
        Next
        XLTableExists = False

    End Function

End Module
