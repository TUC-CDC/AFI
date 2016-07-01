Imports System.Data
Imports System.Data.SqlClient
'Imports CrystalDecisions.CrystalReports.Engine
Module Module1
    Public loginid As Integer ' save current user
    Public loginname As String
    Public loginlevel As String


    Public conn As SqlConnection ' connect to server and save connection
    Public cmd As SqlCommand ' variable send query to database
    Public dr As SqlDataReader 'read data and express as array for select


    Public ServerName As String = "SQL-THB4"
    Public ServerLogin As String = "sa"
    Public ServerPass As String = "1234"
    Public ServerDBName As String = "AFI"

    Public PathApp As String = "C:\IEIP Project\AFI\AFI\AFI\"
    Public Function ConnectSQL()
        'connect to SQL server
        'conn = New SqlConnection("data source=" & ServerName & ";initial catalog=" & ServerDBName & ";user id = " & ServerLogin & ";password=" & ServerPass & ";")
        conn = New SqlConnection("data source=" & ServerName & ";initial catalog=" & ServerDBName & ";Trusted_Connection=Yes;")
        conn.Open()
        Return conn

    End Function
    Public Function ChangeToThaiBaht(ByVal moneyValue As String) As String
        Dim digit() As String = {"เอ็ด", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน", "สิบล้าน", "ร้อยล้าน", "พันล้าน"}
        Dim tmp() As String = moneyValue.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
        Dim intNumber = tmp(0).Replace(",", "")
        Dim decNumber As String
        If (tmp.Length > 1) Then
            decNumber = tmp(1)
        Else
            decNumber = ""
        End If

        Dim result As String = ""
        Dim testChar As String = ""
        Dim splitChar As String = ""

        Dim wordLength As Integer = intNumber.Length - 1
        For intPosition As Integer = 0 To wordLength
            splitChar = intNumber.Substring(intPosition, 1)
            Select Case splitChar
                Case "0" : testChar = ""
                Case "1" : testChar = "หนึ่ง"
                Case "2" : testChar = "สอง"
                Case "3" : testChar = "สาม"
                Case "4" : testChar = "สี่"
                Case "5" : testChar = "ห้า"
                Case "6" : testChar = "หก"
                Case "7" : testChar = "เจ็ด"
                Case "8" : testChar = "แปด"
                Case "9" : testChar = "เก้า"
            End Select

            Select Case (wordLength - intPosition)
                Case 0
                    If (splitChar = "1") Then
                        If (wordLength = 0) Then
                            result = result & testChar
                        Else
                            result = result & digit(wordLength - intPosition)
                        End If
                    Else
                        result = result & testChar
                    End If
                Case 1
                    Select Case splitChar
                        Case "0"
                        Case "1"
                            result = result & digit(wordLength - intPosition)
                        Case "2"
                            testChar = "ยี่"
                            result = result & testChar & digit(wordLength - intPosition)
                        Case Else
                            result = result & testChar & digit(wordLength - intPosition)
                    End Select
                Case Else
                    Select Case splitChar
                        Case "0"
                        Case Else
                            result = result & testChar & digit(wordLength - intPosition)
                    End Select
            End Select

        Next
        result = result & "บาท"

        Dim stang As String = ""
        If (decNumber.Length > 0) Then
            wordLength = decNumber.Length - 1
            For intPosition As Integer = 0 To wordLength
                splitChar = decNumber.Substring(intPosition, 1)
                Select Case splitChar
                    Case "0" : testChar = ""
                    Case "1" : testChar = "หนึ่ง"
                    Case "2" : testChar = "สอง"
                    Case "3" : testChar = "สาม"
                    Case "4" : testChar = "สี่"
                    Case "5" : testChar = "ห้า"
                    Case "6" : testChar = "หก"
                    Case "7" : testChar = "เจ็ด"
                    Case "8" : testChar = "แปด"
                    Case "9" : testChar = "เก้า"
                End Select

                Select Case (wordLength - intPosition)
                    Case 0
                        If (splitChar = "1") Then
                            If (wordLength = 0) Then
                                stang = stang & testChar
                            Else
                                stang = stang & digit(wordLength - intPosition)
                            End If
                        Else
                            stang = stang & testChar
                        End If
                    Case 1
                        Select Case splitChar
                            Case "0"
                            Case "1"
                                stang = stang & digit(wordLength - intPosition)
                            Case "2"
                                testChar = "ยี่"
                                stang = stang & testChar & digit(wordLength - intPosition)
                            Case Else
                                stang = stang & testChar & digit(wordLength - intPosition)
                        End Select
                    Case Else
                        Select Case splitChar
                            Case "0"
                            Case Else
                                stang = stang & testChar & digit(wordLength - intPosition)
                        End Select
                End Select

            Next
        End If

        If (stang.Trim().Length > 0) Then
            stang = stang & "สตางค์"
        Else
            stang = "ถ้วน"
        End If
        result = result & stang
        Return result

    End Function

    'Public Function ConnectReport(ByVal rpt1 As ReportDocument) As ReportDocument

    '    '-----แก้ไขการกรอกรหัสผ่านทุกครั้ง-----
    '    Dim tblLogon As CrystalDecisions.CrystalReports.Engine.Table
    '    Dim rptLogon As CrystalDecisions.Shared.TableLogOnInfo
    '    For Each tblLogon In rpt1.Database.Tables
    '        rptLogon = tblLogon.LogOnInfo
    '        With rptLogon.ConnectionInfo
    '            .ServerName = ServerName
    '            .UserID = ServerLogin
    '            .Password = ServerPass
    '            .DatabaseName = ServerDBName
    '        End With
    '        tblLogon.ApplyLogOnInfo(rptLogon)
    '    Next tblLogon
    '    '--------------------------------------------
    '    Return rpt1


    'End Function
End Module
