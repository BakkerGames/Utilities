' -------------------------------
' --- ModMain.vb - 08/03/2016 ---
' -------------------------------

' ----------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text

Module ModMain

    Public Sub Main()

        Dim SourcePath As String = "C:\Arena_Scripts\Tables Arena\"
        Dim TargetPath As String = "C:\Arena_Scripts\DO_Arena\"
        'Dim SourcePath As String = "C:\Arena_Scripts\Tables Advantage\"
        'Dim TargetPath As String = "C:\Arena_Scripts\DO_Advantage\"
        'Dim SourcePath As String = "C:\Arena_Scripts\Tables IDRIS\"
        'Dim TargetPath As String = "C:\Arena_Scripts\DO_IDRIS\"

        For Each Filename As String In Directory.GetFiles(SourcePath, "*.table.sql")
            Dim SourceInfo As New FileInfo(Filename)
            Dim DomainObjName As String
            DomainObjName = SourceInfo.Name.Replace(".", "_")
            DomainObjName = DomainObjName.Substring(0, DomainObjName.Length - 10) ' Remove .table.sql
            If DomainObjName.IndexOf("dbo_", StringComparison.OrdinalIgnoreCase) = 0 Then
                DomainObjName = DomainObjName.Substring(4)
            End If
            If DomainObjName.IndexOf("###") >= 0 Then
                Continue For
            End If
            If DomainObjName.IndexOf("temp_", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Continue For
            End If
            If DomainObjName.IndexOf("_hist", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Continue For
            End If
            Console.WriteLine(Filename) ' ### for testing ###
            Dim Lines() As String = File.ReadAllLines(Filename)
            Dim InSpecs As Boolean = False
            Dim Result As New StringBuilder
            For Each CurrLine As String In Lines
                CurrLine = CurrLine.Trim
                If CurrLine.IndexOf("CREATE TABLE", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    InSpecs = True
                    With Result
                        .AppendLine($"Public Class DO_{DomainObjName}")
                    End With
                ElseIf InSpecs AndAlso CurrLine.StartsWith("[") Then
                    Dim FieldName As String = CurrLine.Substring(1, CurrLine.IndexOf("]") - 1)
                    If FieldName.IndexOf(" ") >= 0 Then
                        Filename = FieldName.Replace(" "c, "_"c)
                    End If
                    If String.Compare(FieldName, "PACKED_DATA", True) = 0 Then
                        Continue For
                    End If
                    If String.Compare(FieldName, "ROWVERSION", True) = 0 Then
                        FieldName = "VersionID"
                    End If
                    Dim DataType As String = CurrLine.Substring(CurrLine.IndexOf("]") + 1).Trim.ToUpper
                    If DataType.StartsWith("[") Then
                        DataType = DataType.Substring(1)
                    End If
                    DataType = DataType.Substring(0, DataType.IndexOf("]"))
                    Dim NullableFlag As Boolean = (CurrLine.IndexOf("NOT NULL", StringComparison.OrdinalIgnoreCase) < 0)
                    Dim VBDataType As String = ""
                    Select Case DataType
                        Case "INT"
                            VBDataType = "Integer" ' Int32
                        Case "SMALLINT"
                            VBDataType = "Short" ' Int16
                        Case "TINYINT"
                            VBDataType = "Byte"
                        Case "CHAR", "VARCHAR", "NCHAR", "NVARCHAR", "SYSNAME"
                            VBDataType = "String"
                        Case "DECIMAL", "NUMERIC", "MONEY", "SMALLMONEY", "FLOAT", "REAL"
                            VBDataType = "Decimal"
                        Case "DATE", "DATETIME"
                            VBDataType = "DateTime"
                        Case "TIMESTAMP"
                            VBDataType = "Int64" ' Long, but better to be exact
                        Case "BIT"
                            VBDataType = "Boolean"
                        Case "UNIQUEIDENTIFIER"
                            VBDataType = "String"
                        Case "BINARY"
                            If FieldName = "VersionID" Then
                                VBDataType = "Int64" ' Long, but better to be exact
                            Else
                                VBDataType = "Byte()"
                            End If
                        Case Else
                            If DataType.StartsWith("AS (") Then ' Calculated field
                                Continue For
                            End If
                            VBDataType = $"### {DataType} ###"
                    End Select
                    If NullableFlag AndAlso VBDataType <> "String" Then
                        VBDataType += "?"
                    End If
                    Result.Append("    Public Property ")
                    Result.Append(FieldName)
                    Result.Append(" As ")
                    Result.Append(VBDataType)
                    Result.AppendLine()
                Else
                    InSpecs = False
                End If
            Next
            If Result.Length > 0 Then
                With Result
                    .AppendLine("End Class")
                End With
                ' --- See if there are any differences ---
                Dim DOFileName As String = $"{TargetPath}DO_{DomainObjName}.vb"
                Dim NewFileContents As String = Result.ToString
                If File.Exists(DOFileName) Then
                    Dim OrigFileContents As String = File.ReadAllText(DOFileName)
                    If NewFileContents <> OrigFileContents Then
                        File.WriteAllText(DOFileName, NewFileContents)
                    End If
                Else
                    File.WriteAllText(DOFileName, NewFileContents)
                End If
            End If
        Next
    End Sub

End Module
