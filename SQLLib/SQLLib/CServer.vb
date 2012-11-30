'Copyright (C) 2012
'Steve Calandra
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
'documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
'the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and 
'to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
'THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS 
'OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR 
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Imports MySql.Data.MySqlClient
Imports Microsoft.SqlServer.Server

Public Enum DBTYPE
    MYSQL = 0
    SQLSERVER
End Enum

Public Class CServer
    Public serverName As String
    Private _connection As Object
    Private _dataReader As Object
    Private _connectionString As String
    Private _open As Boolean = False
    Private _server As String = ""
    Private _dataBase As String = ""
    Private _type As DBTYPE = 0
    Public showErrors As Boolean = True

    Private _mySQL As List(Of Type)
    Private _sqlServer As List(Of Type)

    Public Sub New(ByVal connectionString As String, Optional ByVal db As DBTYPE = SQLLib.DBTYPE.MYSQL)
        _type = db
        _connectionString = connectionString

        If (_type = SQLLib.DBTYPE.MYSQL) Then
            Try
                _connection = New MySqlConnection(_connectionString)
                _connection.Open()
                _open = True
                _server = serverName
                '_dataBase = DBName
            Catch ex As MySqlException
                MsgBox(ex.Message & " " & serverName & " was not found.")
                _connection.Dispose()
                _connection = Nothing
                Throw New AggregateException(ex.Message & " " & serverName & " was not found.")
            End Try
        Else
            Try
                _connection = New SqlClient.SqlConnection(_connectionString)
                _connection.Open()
                _open = True
                _server = serverName
                '_dataBase = DBName
            Catch ex As SqlClient.SqlException
                MsgBox(ex.Message)
                _connection.Dispose()
                _connection = Nothing
                Throw New AggregateException(ex.Message)
            End Try
        End If


    End Sub

    Public ReadOnly Property dbType() As DBTYPE
        Get
            Return _type
        End Get
    End Property

    Public ReadOnly Property server() As String
        Get
            Return _server
        End Get
    End Property

    Public ReadOnly Property dataBase() As String
        Get
            Return _dataBase
        End Get
    End Property

    Public Sub close()
        Dim sql As String = ""

        _connection.Close()
        _connection.Dispose()

        _open = False

    End Sub

    Public ReadOnly Property isOpen() As Boolean
        Get
            Return _open
        End Get
    End Property

    Public Function callStoredProc(ByVal procedure As String, ByVal ParamArray params() As SQLLib.sqlServerParam)
        If (_type <> SQLLib.DBTYPE.SQLSERVER) Then
            Throw New AggregateException("This library does not yet support MySQL stored procedures.")
        End If

        Try
            _dataReader.Close()
            _dataReader = Nothing
        Catch ex As Exception
        End Try

        Dim _cmd As New SqlClient.SqlCommand()
        _cmd.Connection = _connection
        _cmd.CommandType = CommandType.StoredProcedure
        _cmd.CommandText = procedure

        For Each param As SQLLib.sqlServerParam In params
            _cmd.Parameters.Add(param.wrap())
        Next

        Try
            _dataReader = _cmd.ExecuteReader
        Catch ex As Exception
            If (showErrors) Then
                MsgBox(ex.Message)
            End If

            Return False
        End Try
        Return True

    End Function

    Public Function query(ByVal queryString As String) As Boolean
        Dim _cmd As Object = Nothing

        If (Trim(queryString) = "") Then
            MsgBox("Query string empty!")
            Return False
        End If

        Try
            _dataReader.Close()
            _dataReader = Nothing
        Catch ex As Exception
        End Try

        If (_type = dbType.MYSQL) Then
            _cmd = New MySqlCommand(queryString, _connection)
        Else
            _cmd = New SqlClient.SqlCommand(queryString, _connection)
        End If

        Try
            _dataReader = _cmd.ExecuteReader
        Catch ex As Exception
            If (showErrors) Then
                MsgBox(ex.Message)
            End If

            Return False
        End Try
        Return True
    End Function

    Public Function queryAndRead(ByVal queryString As String) As Boolean
        If (query(queryString)) Then
            Return read()
        Else
            Return False
        End If
    End Function

    Public Function read() As Boolean
        Return _dataReader.Read
    End Function

    Public Function getReaderVal(ByVal col As Integer) As String
        Return _dataReader(col).ToString()
    End Function

    Public Function getReaderVal(ByVal col As String) As String
        Return _dataReader(col).ToString()
    End Function

End Class
