
Imports System.Data.Common
Imports System.Threading
Imports System.Text


'--------------------------------------------------------------------------------------------------------------------
'       ConnectionPool class
'
' usage for a non-parameterized ConnectionPool:
'    dummy = New ConnectionPool(Of NpgsqlConnection)("BankIO", "Server=localhost;Port=5432;User Id=rma;Password=rma;Database=rma_bankio;")
' This call registers a connection pool named 'BankIn'. The pooled connections use the given connection string.
'
' usage for a parameterized ConnectionPool:
' If there are many similar databases on the same server, it makes sense to have one 'configurable' ConnectionPool
' for all connections. A good example is SQL-Ledger; every client gets a set of (identical) tables in his own
' database. Here is the connection string:
'    dummy = New ConnectionPool(Of NpgsqlConnection)("SQLLedger", "Server=localhost;Port=5432;User Id=sql-ledger;Password=De99elria;Database={0}", True)
' Note the additional parameter True, and the "Database={0}" part in the string. When you get a connection from this pool,
' you have to specify a connParameter which replaces the {0} in the connection string.


Public Class ConnectionPool(Of connType As {DbConnection, New})

    Private name As String
    Private connStr As String
    Private configurable As Boolean
    Private pool As New Dictionary(Of String, List(Of connType))
    Private initialThreadID As Integer

    Sub New(ByVal poolname As String, ByVal connStr As String, Optional ByRef configurable As Boolean = False)
        Me.name = poolname
        Me.connStr = connStr
        Me.configurable = configurable

        ' store thread ID
        initialThreadID = Thread.CurrentThread.ManagedThreadId
    End Sub


    Public Function GetConnection(Optional ByVal connParameter As String = Nothing) As connType
        If configurable And connParameter Is Nothing Then
            Throw New ApplicationException("DBPool: Must specify connParameter for a configurable ConnectionPool.")
        ElseIf Not configurable Then
            If connParameter IsNot Nothing Then
                Throw New ApplicationException("DBPool: Can't specify connParameter for a non-configurable ConnectionPool")
            Else
                connParameter = ""
            End If
        End If
        connParameter = connParameter.Trim.ToLower()

        Dim conn As connType = Nothing
        Dim nonPooled = (initialThreadID <> Thread.CurrentThread.ManagedThreadId)

        If Not nonPooled Then
            If Not pool.ContainsKey(connParameter) Then
                ' create pool list, if it doesn't exist
                pool(connParameter) = New List(Of connType)
            Else
                ' pool exists - look for a available connection
                For Each testconn In pool(connParameter)
                    If (testconn.State = ConnectionState.Closed Or testconn.State = ConnectionState.Broken) Then
                        conn = testconn
                        Exit For
                    End If
                Next
            End If
        End If

        If conn Is Nothing Then
            ' no or no available connection - create a new one and store in pool
            conn = New connType
            If configurable Then
                conn.ConnectionString = String.Format(connStr, connParameter)
            Else
                conn.ConnectionString = connStr
            End If

            If Not nonPooled Then
                pool(connParameter).Add(conn)
            End If
        End If

        ' attempt to open the connection.. prepare for errors
        Try
            conn.Open()
        Catch ex As Exception
            Dim msg As String = "DBPool.GetConnection(Pool = " & name
            If configurable Then
                msg &= "." & connParameter
            End If
            msg &= ")" & vbLf & vbLf & String.Format("Die Datenbank '{0}' ist nicht erreichbar.", name)
            Throw New ApplicationException(msg)
        End Try
        Return conn
    End Function

    Public Sub EmptyPool(Optional ByVal connParameter As String = Nothing)
        If configurable And connParameter Is Nothing Then
            Throw New ApplicationException("DBPool: Must specify connParameter for a configurable ConnectionPool.")
        ElseIf Not configurable Then
            If connParameter IsNot Nothing Then
                Throw New ApplicationException("DBPool: Can't specify connParameter for a non-configurable ConnectionPool")
            Else
                connParameter = ""
            End If
        End If
        connParameter = connParameter.Trim.ToLower()

        If pool.ContainsKey(connParameter) Then
            ' pool exists - look for a available connection
            For i = pool(connParameter).Count - 1 To 0 Step -1
                Dim testconn = pool(connParameter)(i)
                If (testconn.State = ConnectionState.Closed Or testconn.State = ConnectionState.Broken) Then
                    pool(connParameter).RemoveAt(i)
                End If
            Next
        End If
    End Sub


    Public Shared Function GetCommand(ByRef conn As DbConnection, ByVal sql As String)
        Dim cmd As DbCommand = conn.CreateCommand
        cmd.CommandText = sql
        Return cmd
    End Function



    '--------------------------------------------------------------------------------------------------------------------
    '       Command executor
    '
    ' use this for INSERT, DELETE, UPDATE

    Public Function SQLExec(ByVal sql As String, Optional ByVal connParameter As String = Nothing) As Integer
        ' executes the given command(s) using the given connection or a connection of the specified pool
        ' & returns the number of affected rows.
        Dim conn = GetConnection(connParameter)
        Dim cmd = GetCommand(conn, sql)

        Dim affected As Integer = 0
        Try
            affected = cmd.ExecuteNonQuery()
        Finally
            conn.Close()
        End Try

        Return affected
    End Function



    '--------------------------------------------------------------------------------------------------------------------
    '       Data getters
    '

    Public Function SQL2O(ByVal sql As String, Optional ByVal connParameter As String = Nothing) As Object
        ' gets the requested data using the given connection or a connection of the specified pool
        ' & returns a single Object. use this for count(*) and such
        Dim conn = GetConnection(connParameter)
        Dim cmd = GetCommand(conn, sql)

        Dim result As Object = Nothing
        Try
            Dim dr = cmd.ExecuteReader()

            ' get data rows
            If dr.Read() Then
                result = dr.GetValue(0)
            End If
            dr.Close()

        Catch ex As Exception
        Finally
            conn.Close()
        End Try

        Return result
    End Function


    Public Function SQL2LO(ByVal sql As String, Optional ByVal connParameter As String = Nothing) As List(Of Object)
        ' gets the requested data using the given connection or a connection of the specified pool
        ' & returns a List of Object.
        ' If the query returns more than one column, only the first one is returned; the rest is ignored.
        Dim conn = GetConnection(connParameter)
        Dim cmd = GetCommand(conn, sql)

        Dim resultSet As New List(Of Object)
        Try
            Dim dr = cmd.ExecuteReader()

            ' get data rows
            While (dr.Read())
                Dim rec As Object = New Object
                rec = dr.GetValue(0)
                resultSet.Add(rec)
            End While
            dr.Close()

        Catch ex As Exception
        Finally
            conn.Close()
        End Try

        Return resultSet
    End Function


    Public Function SQL2LAO(ByVal sql As String, Optional ByVal connParameter As String = Nothing) As List(Of Object())
        ' gets the requested data using the given connection or a connection of the specified pool
        ' & returns a List of Object()
        Dim conn = GetConnection(connParameter)
        Dim cmd = GetCommand(conn, sql)

        Dim resultSet As New List(Of Object())
        Try
            Try
                Dim dr = cmd.ExecuteReader()
                Dim items As Integer = dr.FieldCount

                ' get data rows
                While (dr.Read())
                    Dim rec As Object() = New Object(items - 1) {}
                    dr.GetValues(rec)
                    resultSet.Add(rec)
                End While
                dr.Close()
            Finally
                conn.Close()
            End Try

        Catch ex As Exception
        End Try

        Return resultSet
    End Function



    Public Function SQL2LD(ByVal sql As String, Optional ByVal connParameter As String = Nothing) As List(Of RawDataRecord)
        ' gets the requested data using the given connection or a connection of the specified pool
        ' & returns a List of RawDataRecord
        Dim conn = GetConnection(connParameter)
        Dim cmd = GetCommand(conn, sql)
        Dim resultList As New List(Of RawDataRecord)
        Try
            Try
                Dim dr = cmd.ExecuteReader()
                Dim items As Integer = dr.FieldCount

                ' get column names
                Dim cNames As String() = New String(items - 1) {}
                For i As Integer = 0 To items - 1
                    cNames(i) = dr.GetName(i)
                Next

                ' get data rows
                Dim rec As Object() = New Object(items - 1) {}
                While (dr.Read())
                    Try
                        dr.GetValues(rec)

                        ' .. and fill into Directory
                        Dim rdr As New RawDataRecord
                        For i As Integer = 0 To items - 1
                            rdr(cNames(i)) = rec(i)
                        Next

                        resultList.Add(rdr)
                    Catch ex As Exception
                    End Try
                End While
                dr.Close()
            Finally
                conn.Close()
            End Try

        Catch ex As Exception
        End Try

        Return resultList
    End Function

End Class


' This is what SQL2LD(...) returns
Public Class RawDataRecord
    Inherits Dictionary(Of String, Object)

    '--------------------------------------------------------------------------------------------------------------------
    '       Data translators
    '
    ' these members are used to fill DB data into VB data objects. Empty fields,
    ' special contents and formatting can be processed/corrected here

    Public Function AsString(ByVal key As String, Optional ByVal trim As Boolean = True, Optional ByVal nilAsEmpty As Boolean = False) As String
        Dim value = Me(key)
        If IsDBNull(value) Or value Is Nothing Then
            If nilAsEmpty Then
                Return ""
            Else
                Return Nothing
            End If
        ElseIf Not TypeOf value Is String Then
            value = value.ToString()
        End If

        If trim Then
            Return value.Trim()
        End If
        Return value
    End Function

    Public Function AsChar(ByVal key As String) As Char
        Dim value = Me(key)
        If TypeOf value Is Char Then
            Return value
        ElseIf TypeOf value Is String AndAlso value.ToString.Length > 0 Then
            Return value.ToString(0)
        End If
        Return ""
    End Function

    Public Function AsBoolean(ByVal key As String) As Boolean
        Dim value = Me(key)
        If TypeOf value Is Boolean Then
            Return value
        ElseIf IsDBNull(value) Or value Is Nothing Then
            Return False
        ElseIf value Like "*1*" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AsBooleanNIL(ByVal key As String) As Boolean?
        Dim value = Me(key)
        If TypeOf value Is Boolean Then
            Return value
        ElseIf IsDBNull(value) Or value Is Nothing Then
            Return Nothing
        ElseIf value Like "*1*" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AsInteger(ByVal key As String, Optional ByVal nullAsMinInt As Boolean = False) As Integer
        Dim value = Me(key)
        If IsDBNull(value) Or value Is Nothing Then
            Return If(nullAsMinInt, Integer.MinValue, 0)
        End If
        Return value
    End Function

    Public Function AsIntegerNIL(ByVal key As String, Optional ByVal nullAsMinInt As Boolean = False) As Integer?
        Dim value = Me(key)
        If IsDBNull(value) Or value Is Nothing Then
            Return Nothing
        End If
        Return value
    End Function

    Public Function AsDouble(ByVal key As String, Optional ByVal nilAsZero As Boolean = True) As Double
        Dim value = Me(key)
        If IsDBNull(value) OrElse value Is Nothing Then
            If nilAsZero Then
                Return 0
            Else
                Return Nothing
            End If
        End If

        Dim thisDouble As Double
        Try
            thisDouble = value
        Catch ex As Exception
            If nilAsZero Then
                Return 0
            Else
                Return Nothing
            End If
        End Try
        Return thisDouble
    End Function

    Public Function AsDoubleNIL(ByVal key As String, Optional ByVal nilAsZero As Boolean = True) As Double?
        Dim value = Me(key)
        If IsDBNull(value) OrElse value Is Nothing Then
            Return Nothing
        End If

        Dim thisDouble As Double
        Try
            thisDouble = value
        Catch ex As Exception
            Return Nothing
        End Try
        Return thisDouble
    End Function

    Public Function AsDate(ByVal key As String) As Date
        Dim value = Me(key)
        If IsDBNull(value) Or value Is Nothing Then
            Return Nothing
        End If

        Dim thisDate As Date
        Try
            thisDate = value
        Catch ex As Exception
            Return Nothing
        End Try
        Return thisDate
    End Function

    Public Function AsDateNIL(ByVal key As String) As Date?
        Dim value = Me(key)
        If IsDBNull(value) Or value Is Nothing Then
            Return Nothing
        End If

        Dim thisDate As Date?
        Try
            thisDate = value
        Catch ex As Exception
            Return Nothing
        End Try
        Return thisDate
    End Function

    Public Function AsXML2Dictionary(ByVal key As String) As DictionaryWithDefault(Of String, Object)
        ' parses the header XML field & the transaction XML field and returns all properties in a new dictionary.
        Dim xml = Me(key)
        Dim value As New DictionaryWithDefault(Of String, Object)

        Try
            Dim xmlElem = XElement.Parse(xml)
            For Each attr In xmlElem.Attributes
                value.Add(attr.Name.ToString, attr.Value)
            Next
        Catch ex As Exception
        End Try

        Return value
    End Function

    Public Overrides Function ToString() As String
        Dim result = New StringBuilder()
        For Each key In Keys.OrderBy(Function(_k) _k)
            result.Append(key & Me(key))
        Next
        Return result.ToString
    End Function
End Class


