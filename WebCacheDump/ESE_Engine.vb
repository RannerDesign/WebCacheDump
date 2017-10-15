'   ESE_Engine.vb
'   Origin:
'   http://www.emmet-gray.com/Articles/ESE.html
'   License:
'   "The source code posted here is made available free of charge as `public domain` software.
'   There Is no licensing requirement. That means you can Do anything you want With this software, to include making money from it!
'   However, that also means that there Is no implied warranty And I'm under no obligation to provide you support."
'
Imports System.Collections.Specialized
Imports System.Runtime.InteropServices

Public Class ESE_Engine
    '   Class-wide variables
    Dim InstanceID, SessionID As IntPtr
    Dim DatabaseID As Integer
    Dim DatabaseFile As String

    ' Note: Beginning with Vista, there is an (optional) Unicode version of
    ' the ESE database. But I have't seen one yet, so I didn't implement it here.


    'JET_ERR JET_API JetGetDatabaseFileInfo(
    '  const tchar* szDatabaseName,
    '  void* pvResult,          ' hard wired for Int32
    '  unsigned long cbMax,
    '  unsigned long InfoLevel
    ');
    Declare Ansi Function JetGetDatabaseFileInfo Lib "esent.dll" ( _
        ByVal szDatabaseName As String, _
        ByRef pvResult As Integer, _
        ByVal cbMax As Integer, _
        ByVal InfoLevel As Integer
    ) As Integer

    'JET_ERR JET_API JetCreateInstance(
    '  JET_INSTANCE* pinstance,
    '  const tchar* szInstanceName
    ');
    Declare Ansi Function JetCreateInstance Lib "esent.dll" ( _
        ByRef instance As IntPtr, _
        ByVal szInstanceName As String _
    ) As Integer

    'JET_ERR JET_API JetInit(
    '  JET_INSTANCE* pinstance
    ');
    Declare Function JetInit Lib "esent.dll" ( _
        ByRef instance As IntPtr _
    ) As Integer

    'JET_ERR JET_API JetBeginSession(
    '  JET_INSTANCE instance,
    '  JET_SESID* psesid,
    '  JET_PCSTR szUserName,    ' Not used
    '  JET_PCSTR szPassword     ' Not used
    ');
    Declare Ansi Function JetBeginSession Lib "esent.dll" ( _
        ByVal instance As IntPtr, _
        ByRef psesid As IntPtr, _
        ByVal szUserName As String, _
        ByVal szPassword As String _
    ) As Integer

    'JET_ERR JET_API JetAttachDatabase(
    '  JET_SESID sesid,
    '  const tchar* szFilename, ' ANSI only
    '  JET_GRBIT grbit
    ');
    Declare Ansi Function JetAttachDatabase Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal szFilename As String, _
        ByVal grpit As Integer _
    ) As Integer

    'JET_ERR JET_API JetOpenDatabase(
    '  JET_SESID sesid,
    '  const tchar* szFilename, ' ANSI Only
    '  const tchar* szConnect,  ' Not Used
    '  JET_DBID* pdbid,
    '  JET_GRBIT grbit
    ');
    Declare Ansi Function JetOpenDatabase Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal szFilename As String, _
        ByVal szConnect As String, _
        ByRef pdbid As Integer, _
        ByVal grpit As Integer _
    ) As Integer

    'JET_ERR JET_API JetOpenTable(
    '  JET_SESID sesid,
    '  JET_DBID dbid,
    '  const tchar* szTableName,    ' ANSI Only
    '  const void* pvParameters,
    '  unsigned long cbParameters,
    '  JET_GRBIT grbit,
    '  JET_TABLEID* ptableid
    ');
    Declare Ansi Function JetOpenTable Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal dbid As Integer, _
        ByVal szTableName As String, _
        ByVal pvParameters As Integer, _
        ByVal cbParameters As Integer, _
        ByVal grbit As Integer, _
        ByRef ptableid As IntPtr _
    ) As Integer

    'JET_ERR JET_API JetGetTableColumnInfo(
    '  JET_SESID sesid,
    '  JET_TABLEID tableid,
    '  const tchar* szColumnName,   ' ANSI Only
    '  void* pvResult,
    '  unsigned long cbMax,
    '  unsigned long InfoLevel
    ');
    Declare Ansi Function JetGetTableColumnInfo Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal tableid As IntPtr, _
        ByVal szColumnName As String, _
        ByVal pvResult As IntPtr, _
        ByVal cbMax As Integer, _
        ByVal InfoLevel As Integer _
    ) As Integer

    'JET_ERR JET_API JetRetrieveColumn(
    '  JET_SESID sesid,
    '  JET_TABLEID tableid,
    '  JET_COLUMNID columnid,
    '  void* pvData,
    '  unsigned long cbData,
    '  unsigned long* pcbActual,
    '  JET_GRBIT grbit,
    '  JET_RETINFO* pretinfo
    ');
    Declare Function JetRetrieveColumn Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal tableid As IntPtr, _
        ByVal columnid As Integer, _
        ByVal pvData As IntPtr, _
        ByVal cbData As Integer, _
        ByRef pcbActual As Integer, _
        ByVal grbit As Integer, _
        ByRef pretinfo As JET_RETINFO _
    ) As Integer

    'JET_ERR JET_API JetMove(
    '  JET_SESID sesid,
    '  JET_TABLEID tableid,
    '  long cRow,
    '  JET_GRBIT grbit
    ');
    Declare Function JetMove Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal tableid As IntPtr, _
        ByVal cRow As Integer, _
        ByVal grbit As Integer _
    ) As Integer

    'JET_ERR JET_API JetCloseTable(
    '  JET_SESID sesid,
    '  JET_TABLEID tableid
    ');
    Declare Function JetCloseTable Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal tableid As IntPtr _
    ) As Integer

    'JET_ERR JET_API JetCloseDatabase(
    '  JET_SESID sesid,
    '  JET_DBID dbid,
    '  JET_GRBIT grbit
    ');
    Declare Function JetCloseDatabase Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal dbid As Integer, _
        ByVal grpit As Integer _
    ) As Integer

    'JET_ERR JET_API JetDetachDatabase(
    '  JET_SESID sesid,
    '  const tchar* szFilename
    ');
    Declare Ansi Function JetDetachDatabase Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal szFilename As String _
    ) As Integer

    'JET_ERR JET_API JetEndSession(
    '  JET_SESID sesid,
    '  JET_GRBIT grbit
    ');
    Declare Function JetEndSession Lib "esent.dll" ( _
        ByVal sesid As IntPtr, _
        ByVal grbit As Integer _
    ) As Integer

    'JET_ERR JET_API JetTerm(
    '  JET_INSTANCE instance
    ');
    Declare Function JetTerm Lib "esent.dll" ( _
        ByVal instance As IntPtr _
    ) As Integer

    'JET_ERR JET_API JetGetSystemParameter(
    '  JET_INSTANCE instance,
    '  JET_SESID sesid,
    '  unsigned long paramid,
    '  JET_API_PTR* plParam,    ' really is Int32
    '  JET_PSTR szParam,        ' ANSI Only
    '  unsigned long cbMax
    ');
    Declare Ansi Function JetGetSystemParameter Lib "esent.dll" ( _
        ByVal instance As IntPtr, _
        ByVal sesid As IntPtr, _
        ByVal paramid As Integer, _
        ByRef plParam As Integer, _
        ByVal szParam As String, _
        ByVal cbMax As Integer _
    ) As Integer

    'JET_ERR JET_API JetSetSystemParameter(
    '  JET_INSTANCE* pinstance,
    '  JET_SESID sesid,
    '  unsigned long paramid,
    '  JET_API_PTR lParam,      ' really is Int32
    '  JET_PCSTR szParam
    ');
    Declare Ansi Function JetSetSystemParameter Lib "esent.dll" ( _
        ByRef pinstance As IntPtr, _
        ByVal sesid As IntPtr, _
        ByVal paramid As Integer, _
        ByVal lParam As Integer, _
        ByVal szParam As String _
    ) As Integer

    'typedef struct {
    '  unsigned long cbStruct;
    '  JET_TABLEID tableid;
    '  unsigned long cRecord;
    '  JET_COLUMNID columnidPresentationOrder;
    '  JET_COLUMNID columnidcolumnname;
    '  JET_COLUMNID columnidcolumnid;
    '  JET_COLUMNID columnidcoltyp;
    '  JET_COLUMNID columnidCountry;
    '  JET_COLUMNID columnidLangid;
    '  JET_COLUMNID columnidCp;
    '  JET_COLUMNID columnidCollate;
    '  JET_COLUMNID columnidcbMax;
    '  JET_COLUMNID columnidgrbit;
    '  JET_COLUMNID columnidDefault;
    '  JET_COLUMNID columnidBaseTableName;
    '  JET_COLUMNID columnidBaseColumnName;
    '  JET_COLUMNID columnidDefinitionName;
    '} JET_COLUMNLIST;
    ' Note: is not packed (ie Pack:=4)
    <StructLayout(LayoutKind.Sequential)> _
    Structure JET_COLUMNLIST
        Dim cbStruct As Integer
        Dim tableid As IntPtr
        Dim cRecord As Integer
        Dim columnidPresentationOrder As Integer
        Dim columnidcolumnname As Integer
        Dim columnidcolumnid As Integer
        Dim columnidcoltyp As Integer
        Dim columnidCountry As Integer
        Dim columnidLangid As Integer
        Dim columnidCp As Integer
        Dim columnidCollate As Integer
        Dim columnidcbMax As Integer
        Dim columnidgrbit As Integer
        Dim columnidDefault As Integer
        Dim columnidBaseTableName As Integer
        Dim columnidBaseColumnName As Integer
        Dim columnidDefinitionName As Integer
    End Structure

    'typedef struct {
    '  unsigned long cbStruct;
    '  unsigned long ibLongValue;
    '  unsigned long itagSequence;
    '  JET_COLUMNID columnidNextTagged;
    '} JET_RETINFO;
    <StructLayout(LayoutKind.Sequential)> _
    Structure JET_RETINFO
        Dim cbStruct As Integer
        Dim ibLongValue As Integer
        Dim itagSequence As Integer
        Dim columnidNextTagged As Integer
    End Structure

    ' bunch of constants from essent.h
    Const JET_bitDbReadOnly As Integer = 1
    Const JET_bitTableReadOnly As Integer = 4
    Const JET_ColInfoListSortColumnid As Integer = 7

    Const JET_paramErrorToString As Integer = 70
    Const JET_paramTempPath As Integer = 1
    Const JET_paramLogFilePath As Integer = 2
    Const JET_paramSystemPath As Integer = 0
    Const JET_paramAccessDeniedRetryPeriod As Integer = 53
    Const JET_paramDatabasePageSize As Integer = 64

    Const JET_cbNameMost As Integer = 64
    Const JET_wrnColumnNull As Integer = 1004
    Const JET_wrnBufferTruncated As Integer = 1006

    Const JET_MoveFirst As Integer = &H80000000
    Const JET_MoveNext As Integer = 1

    Const JET_coltypNil As Integer = 0
    Const JET_coltypBit As Integer = 1
    Const JET_coltypUnsignedByte As Integer = 2
    Const JET_coltypShort As Integer = 3
    Const JET_coltypLong As Integer = 4
    Const JET_coltypCurrency As Integer = 5
    Const JET_coltypIEEESingle As Integer = 6
    Const JET_coltypIEEEDouble As Integer = 7
    Const JET_coltypDateTime As Integer = 8
    Const JET_coltypBinary As Integer = 9
    Const JET_coltypText As Integer = 10
    Const JET_coltypLongBinary As Integer = 11
    Const JET_coltypLongText As Integer = 12
    Const JET_coltypSLV As Integer = 13
    Const JET_coltypUnsignedLong As Integer = 14 ' New for Vista/2008
    Const JET_coltypLongLong As Integer = 15
    Const JET_coltypGUID As Integer = 16
    Const JET_coltypUnsignedShort As Integer = 17
    Const JET_coltypMax As Integer = 18

    Const JET_DbInfoPageSize As Integer = 17

    Const BUF_SIZE As Integer = 8192

    Public Function DBopen(ByVal ESEdbFile As String)
        '   Prepare ESE database for multiple uses
        '   Create instance, init, attach and open database
        Dim page, ret As Integer

        '   Saved DatabaseFile for this class instance
        DatabaseFile = ESEdbFile

        '   Get the size of the Database Page
        ret = JetGetDatabaseFileInfo(DatabaseFile, page, 4, JET_DbInfoPageSize)
        If ret <> 0 Then Throw New ApplicationException("JetGetDatabaseFileInfo: " & JetErrorMessage(ret))
        '   Set the page size (Note: this is a global setting, so the InstanceID is ignored)
        ret = JetSetSystemParameter(InstanceID, IntPtr.Zero, JET_paramDatabasePageSize, page, Nothing)
        If ret <> 0 Then Throw New ApplicationException("JetSetSystemParameter: " & JetErrorMessage(ret))

        '   DBopen:
        '   JetCreateInstance()
        '       JetInit()
        '           JetBeginSession()
        '               JetAttachDatabase()
        '                   JetOpenDatabase()
        '   Multiple operations like JetOpenTable, JetCloseTable
        '   DBclose:
        '               JetCloseDatabase()
        '           JetDetachDatabase()
        '       JetEndSession()
        '   JetTerm() ' also destroys the Instance

        '   Create an instance
        ret = JetCreateInstance(InstanceID, "SEStest")
        If ret <> 0 Then Throw New ApplicationException("JetCreateInstance: " & JetErrorMessage(ret))

        '   Let's create a log directory if one doesn't already exist
        If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(DatabaseFile) & "\Logs") Then
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(DatabaseFile) & "\Logs")
        End If
        '   Set a few parameters... the location of the Log/Temp/Check directories
        ret = JetSetSystemParameter(InstanceID, IntPtr.Zero, JET_paramTempPath, Nothing, System.IO.Path.GetDirectoryName(DatabaseFile) & "\Logs\")
            If ret <> 0 Then Throw New ApplicationException("JetSetSystemParameter: " & JetErrorMessage(ret))
            ret = JetSetSystemParameter(InstanceID, IntPtr.Zero, JET_paramLogFilePath, Nothing, System.IO.Path.GetDirectoryName(DatabaseFile) & "\Logs\")
            If ret <> 0 Then Throw New ApplicationException("JetSetSystemParameter: " & JetErrorMessage(ret))
            ret = JetSetSystemParameter(InstanceID, IntPtr.Zero, JET_paramSystemPath, Nothing, System.IO.Path.GetDirectoryName(DatabaseFile) & "\Logs\")
            If ret <> 0 Then Throw New ApplicationException("JetSetSystemParameter: " & JetErrorMessage(ret))
            ret = JetSetSystemParameter(InstanceID, IntPtr.Zero, JET_paramAccessDeniedRetryPeriod, 1000, Nothing)
        If ret <> 0 Then            Throw New ApplicationException("JetSetSystemParameter: " & JetErrorMessage(ret))
        
        '   Initialize the instance
        ret = JetInit(InstanceID)
        If ret <> 0 Then Throw New ApplicationException("JetInit: " & JetErrorMessage(ret))

        '   Start the session
        ret = JetBeginSession(InstanceID, SessionID, Nothing, Nothing)
        If ret <> 0 Then Throw New ApplicationException("JetBeginSession: " & JetErrorMessage(ret))

        '   Attach the database file
        ret = JetAttachDatabase(SessionID, DatabaseFile, 0)
        If ret <> 0 Then Throw New ApplicationException("JetAttachDatabase: " & JetErrorMessage(ret))

        '   Open the database
        ret = JetOpenDatabase(SessionID, DatabaseFile, Nothing, DatabaseID, 0)
        If ret <> 0 Then Throw New ApplicationException("JetOpenDatabase: " & JetErrorMessage(ret))

        Return 0

    End Function

    Public Function DBclose(DatabaseFile)
        '   Close ESE database, detach, end session and destroy instance
        Dim ret As Integer
        '   Close the database
        ret = JetCloseDatabase(SessionID, DatabaseID, 0)
        If ret <> 0 Then Throw New ApplicationException("JetCloseDatabase: " & JetErrorMessage(ret))

        '   Detach the database file
        ret = JetDetachDatabase(SessionID, DatabaseFile)
        If ret <> 0 Then Throw New ApplicationException("JetDetachDatabase: " & JetErrorMessage(ret))

        '   End the session
        ret = JetEndSession(SessionID, 0)
        If ret <> 0 Then Throw New ApplicationException("JetEndSession: " & JetErrorMessage(ret))

        '   Terminate the instance
        ret = JetTerm(InstanceID)
        If ret <> 0 Then Throw New ApplicationException("JetTerm: " & JetErrorMessage(ret))

        Return 0

    End Function

    '
    '   Count records in table given
    '   As there is no table attribute like count, this has to be counted by reading in the whole table
    '
    Public Function RecordCount(ByVal TableName As String) As Integer
        Dim table_id As IntPtr
        Dim Rcount, ret As Integer

        '   open the table
        ret = JetOpenTable(SessionID, DatabaseID, TableName, 0, 0, JET_bitTableReadOnly, table_id)
        If ret <> 0 Then Throw New ApplicationException("JetOpenTable: " & JetErrorMessage(ret))
        '   Position to first row
        ret = JetMove(SessionID, table_id, JET_MoveFirst, 0)
        Rcount = 0
        '   Loop through table
        Do While ret = 0
            Rcount = Rcount + 1
            ret = JetMove(SessionID, table_id, JET_MoveNext, 0)
        Loop
        '   close the table
        ret = JetCloseTable(SessionID, table_id)
        If ret <> 0 Then Throw New ApplicationException("JetCloseTable: " & JetErrorMessage(ret))

        Return Rcount
    End Function

    '
    ' Open an table in an ESE database and return all rows/columns in
    ' an ADO.Net DataTable
    '
    Public Function GetTable(ByVal TableName As String, ByVal RetrieveBinaryFields As Boolean) As DataTable
        Dim dt As New DataTable
        Dim dc As DataColumn
        Dim dr As DataRow
        Dim iPtr, table_id As IntPtr
        Dim ret, req_size As Integer
        Dim col_list As JET_COLUMNLIST
        Dim ret_info As JET_RETINFO

        ' open the table
        ret = JetOpenTable(SessionID, DatabaseID, TableName, 0, 0, JET_bitTableReadOnly, table_id)
        If ret <> 0 Then            Throw New ApplicationException("JetOpenTable: " & JetErrorMessage(ret))
        
        '
        ' This is kinda strange... to get the schema of a table, you use the
        ' JetGetTableColumnInfo function which returns a listing of column IDs
        ' that map to parts of the schema.  The actual schema itself is returned
        ' in a temporary table.  So, now that you know the column IDs to the temp
        ' table, you can read the records of the temp table to get the schema
        ' of the real table.  Whew!
        '
        col_list = Nothing
        iPtr = Marshal.AllocHGlobal(Marshal.SizeOf(col_list))
        ret = JetGetTableColumnInfo(SessionID, table_id, Nothing, iPtr, Marshal.SizeOf(col_list), JET_ColInfoListSortColumnid)
        If ret <> 0 Then            Throw New ApplicationException("JetGetTableColumnInfo: " & JetErrorMessage(ret))
        
        ' do some pointer magic to populate the structure
        col_list = CType(Marshal.PtrToStructure(iPtr, GetType(JET_COLUMNLIST)), JET_COLUMNLIST)

        ' clean up
        Marshal.FreeHGlobal(iPtr)

        ' Set the cursor at the begining of the temp table.  Probably not nesseccary,
        ' but what the heck...
        ret = JetMove(SessionID, col_list.tableid, JET_MoveFirst, 0)
        If ret <> 0 Then Throw New ApplicationException("JetMove: " & JetErrorMessage(ret))

        ' allocate a buffer for the colum names, et al
        iPtr = Marshal.AllocHGlobal(JET_cbNameMost)
        ret_info.cbStruct = Marshal.SizeOf(ret_info)

        ' If there are multiple values, we get only the first one
        ret_info.itagSequence = 1

        '
        ' Loop thru each record in the temp table to build our ADO.Net DataTable
        '
        dt = New DataTable
        dt.TableName = TableName
        Do While ret = 0
            Dim col_name As String
            Dim col_id, col_type As Integer

            ' column name
            ret = JetRetrieveColumn(SessionID, col_list.tableid, col_list.columnidcolumnname, iPtr, JET_cbNameMost, req_size, 0, ret_info)
            If ret <> 0 Then                Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        col_name = Marshal.PtrToStringAnsi(iPtr, req_size)

            ' column data type
            ret = JetRetrieveColumn(SessionID, col_list.tableid, col_list.columnidcoltyp, iPtr, 4, req_size, 0, ret_info)
            If ret <> 0 Then                Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        col_type = Marshal.ReadInt32(iPtr)

            ' column ID
            ret = JetRetrieveColumn(SessionID, col_list.tableid, col_list.columnidcolumnid, iPtr, 4, req_size, 0, ret_info)
            If ret <> 0 Then                Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        col_id = Marshal.ReadInt32(iPtr)

            ' build the ADO.Net DataColumn
            Select Case col_type
                Case JET_coltypBit
                    dc = New DataColumn(col_name, GetType(System.Boolean))
                Case JET_coltypUnsignedByte
                    dc = New DataColumn(col_name, GetType(System.Byte))
                Case JET_coltypShort
                    dc = New DataColumn(col_name, GetType(System.Int16))
                Case JET_coltypLong
                    dc = New DataColumn(col_name, GetType(System.Int32))
                Case JET_coltypCurrency
                    dc = New DataColumn(col_name, GetType(System.Decimal))
                Case JET_coltypIEEESingle
                    dc = New DataColumn(col_name, GetType(System.Single))
                Case JET_coltypIEEEDouble
                    dc = New DataColumn(col_name, GetType(System.Double))
                Case JET_coltypDateTime
                    dc = New DataColumn(col_name, GetType(System.DateTime))
                Case JET_coltypBinary, JET_coltypGUID
                    dc = New DataColumn(col_name, GetType(System.Byte()))
                Case JET_coltypText
                    dc = New DataColumn(col_name, GetType(System.String))
                    ' record the Code Page so we can distinguish between ANSI
                    ' and Unicode strings later on in the program
                    ret = JetRetrieveColumn(SessionID, col_list.tableid, col_list.columnidCp, iPtr, 4, req_size, 0, ret_info)
                    If ret <> 0 Then                        Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                                        dc.ExtendedProperties.Add("CodePage", Marshal.ReadInt32(iPtr))
                Case JET_coltypLongBinary
                    dc = New DataColumn(col_name, GetType(System.Byte()))
                Case JET_coltypLongText
                    dc = New DataColumn(col_name, GetType(System.String))
                    ' Code Page
                    ret = JetRetrieveColumn(SessionID, col_list.tableid, col_list.columnidCp, iPtr, 4, req_size, 0, ret_info)
                    If ret <> 0 Then                        Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                                        dc.ExtendedProperties.Add("CodePage", Marshal.ReadInt32(iPtr))
                Case JET_coltypUnsignedLong
                    dc = New DataColumn(col_name, GetType(System.UInt32))
                Case JET_coltypLongLong
                    dc = New DataColumn(col_name, GetType(System.Int64))
                Case JET_coltypUnsignedShort
                    dc = New DataColumn(col_name, GetType(System.UInt16))
                Case Else
                    Throw New ApplicationException("Unsupported column type " & col_type)
            End Select
            ' use the DataColumn's ExtendedProperites to store the Column ID value
            dc.ExtendedProperties.Add("ColumnID", col_id)
            dt.Columns.Add(dc)

            ret = JetMove(SessionID, col_list.tableid, JET_MoveNext, 0)
        Loop

        ' Free the memory we allocated and destroy the temporary table that got
        ' automatically created by the JetGetTableColumnInfo command.
        Marshal.FreeHGlobal(iPtr)
        JetCloseTable(SessionID, col_list.tableid)

        '
        ' OK, now we know the column names and datatypes of the table, we can *finally* read
        ' the records and copy the data into our ADO.Net DataTable
        '

        ret = JetMove(SessionID, table_id, JET_MoveFirst, 0)

        ' Allocate a sufficiently large buffer to handle most needs.  If needed, we allocate
        ' another buffer for larger requirements
        iPtr = Marshal.AllocHGlobal(BUF_SIZE)
        Do While ret = 0
            dr = dt.NewRow
            For Each dc In dt.Columns
                Select Case dc.DataType.ToString
                    Case "System.Boolean"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 1, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        ' Hummm... the documentation is apparently wrong (or ignored
                        ' in practice).  The JET_coltypBit *can* be null
                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = CBool(Marshal.ReadByte(iPtr))
                        Else
                            dr(dc.ColumnName) = False
                        End If
                    Case "System.Byte"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 1, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = Marshal.ReadByte(iPtr)
                        End If
                    Case "System.Int16", "System.UInt16"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 2, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = Marshal.ReadInt16(iPtr)
                        End If
                    Case "System.Int32", "System.UInt32"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 4, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = Marshal.ReadInt32(iPtr)
                        End If
                    Case "System.Int64"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 8, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = Marshal.ReadInt64(iPtr)
                        End If
                    Case "System.Decimal"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 8, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = CType(Marshal.PtrToStructure(iPtr, GetType(Decimal)), Decimal)
                        End If
                    Case "System.Single"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 4, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = CType(Marshal.PtrToStructure(iPtr, GetType(Single)), Single)
                        End If
                    Case "System.Double"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 8, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        If ret <> JET_wrnColumnNull Then
                            dr(dc.ColumnName) = CType(Marshal.PtrToStructure(iPtr, GetType(Double)), Double)
                        End If
                    Case "System.DateTime"
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), iPtr, 8, req_size, 0, ret_info)
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        ' The documentation is wrong... the date format is based upon
                        ' the year 1900 (an Office Automation Date)
                        If ret <> JET_wrnColumnNull Then
                            Dim d As Double
                            d = CType(Marshal.PtrToStructure(iPtr, GetType(Double)), Double)
                            dr(dc.ColumnName) = Date.FromOADate(d).ToLocalTime
                        End If
                    Case "System.Byte[]"
                        ' My application doesn't need any binary fields, so I've made this an
                        ' option to speed things along
                        If RetrieveBinaryFields Then
                            Dim p As IntPtr
                            p = iPtr
                            ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), p, BUF_SIZE, req_size, 0, ret_info)
                            If ret = JET_wrnBufferTruncated Then
                                ' OK, our general purpose buffer isn't large enough
                                p = Marshal.AllocHGlobal(req_size)
                                ' do it again
                                ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), p, req_size, req_size, 0, ret_info)
                            End If
                            If ret <> 0 And ret <> JET_wrnColumnNull Then
                                If p <> iPtr Then
                                    Marshal.FreeHGlobal(p)
                                End If
                                Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                            End If

                            If ret <> JET_wrnColumnNull Then
                                ' copy to a byte array
                                Dim buf(req_size - 1) As Byte
                                Marshal.Copy(p, buf, 0, req_size)
                                dr(dc.ColumnName) = buf
                            End If
                            If p <> iPtr Then
                                Marshal.FreeHGlobal(p)
                            End If
                        End If
                    Case "System.String"
                        Dim p As IntPtr
                        p = iPtr
                        ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), p, BUF_SIZE, req_size, 0, ret_info)
                        If ret = JET_wrnBufferTruncated Then
                            ' our general purpose buffer isn't large enough
                            p = Marshal.AllocHGlobal(req_size)
                            ' do it again
                            ret = JetRetrieveColumn(SessionID, table_id, CInt(dc.ExtendedProperties("ColumnID")), p, req_size, req_size, 0, ret_info)
                        End If
                        If ret <> 0 And ret <> JET_wrnColumnNull Then
                            If p <> iPtr Then
                                Marshal.FreeHGlobal(p)
                            End If
                            Throw New ApplicationException("JetRetrieveColumn: " & JetErrorMessage(ret))
                        End If

                        ' if not null
                        If ret <> JET_wrnColumnNull Then
                            ' Decode the strings based upon ANSI vs Unicode Code Page
                            ' Yes, I know... this isn't very Locale aware.  I guess the
                            ' "Globalization Police" will have to arrest me!
                            If CInt(dc.ExtendedProperties("CodePage")) = 1200 Then
                                dr(dc.ColumnName) = Marshal.PtrToStringUni(iPtr, CInt((req_size / 2) - 1))
                            Else
                                dr(dc.ColumnName) = Marshal.PtrToStringAnsi(iPtr, req_size)
                            End If
                        End If
                        If p <> iPtr Then
                            Marshal.FreeHGlobal(p)
                        End If
                    Case Else
                        Throw New ApplicationException("Unsupported Data type")
                End Select
            Next
            dt.Rows.Add(dr)
            ret = JetMove(SessionID, table_id, JET_MoveNext, 0)
        Loop

        ' close the table
        ret = JetCloseTable(SessionID, table_id)
        If ret <> 0 Then            Throw New ApplicationException("JetCloseTable: " & JetErrorMessage(ret))
        

        Return dt
    End Function

    ' convert a JET_ERR code into a text message
    Private Function JetErrorMessage(ByVal error_code As Integer) As String
        Dim msg As String
        Dim ret As Integer

        msg = Space(256)
        ret = JetGetSystemParameter(IntPtr.Zero, IntPtr.Zero, JET_paramErrorToString, error_code, msg, 256)
        If ret <> 0 Then            Throw New ApplicationException("Can't get error message for code " & error_code)
                Return msg
    End Function

    ' just for testing... not used in the project
    Private Function ByteArrayToString(ByVal ba As Byte()) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        sb.Append("0x")
        For i = 0 To ba.Length - 1
            sb.AppendFormat("{0:X2}", ba(i))
        Next

        Return sb.ToString
    End Function
End Class
