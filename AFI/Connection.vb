Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class sqlControl
    ' Public SQLcon As New SqlConnection With {.ConnectionString = "Server=SQL-THB4;Database=AFI;Trusted_Connection=Yes;"}
    '    Public SQLcon As New SqlConnection With {.ConnectionString = "Server=SQL-THB4;Database=AFI;User=sa;Pwd=xxx;"}
    Public sqlCmd As SqlCommand
    Public SQLDA As SqlDataAdapter
    Public SQLDataset As DataSet

    Public SQLcon As SqlConnection ' connect to server and save connection

    Public ServerName As String = "SQL-THB4"
    Public ServerLogin As String = "sa"
    Public ServerPass As String = "1234"
    Public ServerDBName As String = "AFI"

    Public PathApp As String = "C:\IEIP Project\AFI\AFI\AFI\"

    Public Function ConnectSQL()
        'connect to SQL server
        'SQLcon = New SqlConnection("data source=" & ServerName & ";initial catalog=" & ServerDBName & ";user id = " & ServerLogin & ";password=" & ServerPass & ";")
        SQLcon = New SqlConnection("data source=" & ServerName & ";initial catalog=" & ServerDBName & ";Trusted_Connection=Yes;")
        SQLcon.Open()
        Return SQLcon

    End Function
    Public Function HasConnection() As Boolean
        Try
            ConnectSQL()

         '   SQLcon.Open()

            SQLcon.Close()
            Return True

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Sub RunQuery(ByVal Query As String, ByVal Tname As String)
        Try
            SQLcon.Open()
            sqlCmd = New SqlCommand(Query, SQLcon)
            'Load SQL records for datagrid
            SQLDA = New SqlDataAdapter(sqlCmd)
            SQLDataset = New DataSet
            SQLDA.Fill(SQLDataset, Tname)

            'Dim R As SqlDataReader = sqlCmd.ExecuteReader
            'While R.Read
            '    MsgBox(R.GetName(0) & ": " & R(0))
            '    End While

                SQLcon.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            If SQLcon.State = ConnectionState.Open Then
                SQLcon.Close()
            End If
        End Try
    End Sub
    Public Sub FindQuery(ByVal Query As String, ByVal AID As String)
        Try
            SQLcon.Open()
            sqlCmd = New SqlCommand(Query, SQLcon)
            SQLDA = New SqlDataAdapter(sqlCmd)
            SQLDataset = New DataSet
            SQLDA.Fill(SQLDataset, "AID")

            SQLcon.Close()
            Dim R As Integer = SQLDataset.Tables("AID").Rows.Count
            If R > 0 Then
                ' User already exists
                MsgBox("ID Already Exist!", MsgBoxStyle.Exclamation, "Add New User!")
                frmCRF.FillData()

            Else
                ' User does not exist, add them
                Dim result As Integer = MessageBox.Show("Data is not exist. Will you add a new record?.", "Add New", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    frmCRF.txtAFIID.Text = ""
                    MessageBox.Show("No data added.")
                ElseIf result = DialogResult.Yes Then
                    ' SQLcon.Close()
                    AddNewData(AID)
                    frmCRF.resetAllControls(frmCRF)
                    MsgBox("Records Successfully Added!", MsgBoxStyle.Information, "Add New Customer!")
                End If
            End If


            'SQLcon.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            If SQLcon.State = ConnectionState.Open Then
                SQLcon.Close()
            End If
        End Try
    End Sub

    Public Sub AddNewData(ByVal AFIID As String)

        Try
            Dim strInsert As String = "Insert into tblCase (AFIID) Values ('" & AFIID & "')"

            MsgBox(strInsert)

            SQLcon.Open()
            sqlCmd = New SqlCommand(strInsert, SQLcon)
            sqlCmd.ExecuteNonQuery()

            SQLcon.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function DataUpdate(ByVal Command As String) As Integer
        Try
            SQLcon.Open()
            sqlCmd = New SqlCommand(Command, SQLcon)
            Dim ChangeCount As Integer = sqlCmd.ExecuteNonQuery
            SQLcon.Close()
            MsgBox("Updated successfully")
            Return ChangeCount
        Catch ex As Exception
            MsgBox(ex.Message)
            SQLcon.Close()
        End Try
        Return 0
    End Function






End Class

