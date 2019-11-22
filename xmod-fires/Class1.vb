' **** SIMONCODER SOFTWARE ****
' You may freely change, use and distribute this code under the following conditions:
' 1. You may NOT charge money for the use of this software or any software that uses this code.
' 2. You must keep this copyright information throughout the code.
' **** © 2018-2019 Scott Reed **** mreed1972@gmail.com ****

Imports System.Windows.Forms

Public Class Class1

    ''' <summary>
    ''' Check directory locations for log file
    ''' </summary>
    ''' <param name="DirName">base directory location (ie: c:\logfile)</param>
    ''' <returns>true or false</returns>
    Public Function CheckLogDirectories(DirName As String) As Boolean
        Try
            Dim f As Boolean
            f = My.Computer.FileSystem.DirectoryExists(DirName)
            If f = False Then
                My.Computer.FileSystem.CreateDirectory(DirName)
                My.Computer.FileSystem.WriteAllText(DirName & "\elog.txt", "", False)
                My.Computer.FileSystem.WriteAllText(DirName & "\slog.txt", "", False)
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Check log files
    ''' </summary>
    ''' <param name="DirName">base directory location (ie: c:\logfile)</param>
    ''' <returns>true or false</returns>
    Public Function CheckLogFiles(DirName As String) As Boolean
        Try

            Dim e, s As Boolean
            e = My.Computer.FileSystem.FileExists(DirName & "\elog.txt")
            s = My.Computer.FileSystem.FileExists(DirName & "\slog.txt")
            If e = False Then
                My.Computer.FileSystem.WriteAllText(DirName & "\elog.txt", "", False)
            End If
            If s = False Then
                My.Computer.FileSystem.WriteAllText(DirName & "\slog.txt", "", False)
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    ''' <summary>
    ''' Generate random string of Characters
    ''' </summary>
    ''' <param name="length">INTEGER: length of characters to generate</param>
    ''' <returns>Generate random string of Characters</returns>
    Public Function grs(ByRef length As Integer) As String
        Randomize()
        Dim ac As String
        ac = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim i As Integer
        For i = 1 To length
            grs = grs & Mid(ac, Int(Rnd() * Len(ac) + 1), 1)
        Next
    End Function

    ''' <summary>
    ''' Total values from a datagrid
    ''' </summary>
    ''' <param name="xDG">Name of Datagrid</param>
    ''' <param name="xCelIndx">Cell Index value (Integer)</param>
    ''' <returns>the total value of all the results in a datagrid column.</returns>
    Public Function TDGV(xDG As DataGridView, xCelIndx As Integer)
        Dim xValue As Decimal
        For Each row As DataGridViewRow In xDG.Rows
            xValue += row.Cells(xCelIndx).Value
        Next

        Return String.Format("{0:n0}", Math.Round(xValue, 2))
    End Function

    ''' <summary>
    ''' Writes to ERROR log file.
    ''' </summary>
    ''' <param name="code">Unique ID</param>
    ''' <param name="msg">Message</param>
    ''' <param name="loc">Directory and File location (ex: "c:\TEST\test.txt")</param>
    ''' <returns>T or F</returns>
    Public Function ELog(code As String, msg As String, loc As String) As Boolean
        Try
            Dim dt As DateTime = Date.Now
            Dim Final As String
            Final = vbCrLf & "===== " & code & " =====" & vbCrLf & msg
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(loc, True)
            file.WriteLine(Final)
            file.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Grab a random word from a text file.
    ''' </summary>
    ''' <param name="FileLocation">Location of the text file</param>
    ''' <returns>random word from a text file.</returns>
    Public Function getword(FileLocation As String)
        Randomize()

        Dim sr As System.IO.StreamReader
        Dim ri As Integer = 0
        Dim wa As New ArrayList

        If System.IO.File.Exists(FileLocation) = True Then 'x4 is location of text file
            sr = New IO.StreamReader(FileLocation)

            Do While sr.Peek > -1
                wa.Add(sr.ReadLine)
            Loop
            ri = CInt((wa.Count - 1) * Rnd())
            Return wa(ri) 'The random word, use as needed.
            sr.Close()
        Else
            sr.Close()
        End If
    End Function

    ''' <summary>
    ''' State Cause Codes converted to Federal Cause Codes
    ''' </summary>
    ''' <param name="StateCode">INTEGER: State Cause Code </param>
    ''' <returns>INTEGER: Federal Cause Code </returns>
    Public Function CauseCodeCrossWalk(StateCode As Integer) As Integer
        Select Case StateCode
            Case 1
                Return 7
            Case 2
                Return 5
            Case 3
                Return 3
            Case 4
                Return 6
            Case 5
                Return 4
            Case 6
                Return 2
            Case 7
                Return 8
            Case 8
                Return 1
            Case 9
                Return 9
            Case Else
                Return 0
        End Select
    End Function

    ''' <summary>
    ''' Gets a KEY VALUE from the registry (HKEY_CURRENT_USER) \ subkey \ keysub \ value
    ''' </summary>
    ''' <param name="keySub"></param>
    ''' <param name="keyValue"></param>
    ''' <returns>KEY VALUE</returns>
    Public Function GetMyKey(keySub As String, keyValue As String)
        Dim readValue As String
        readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & keySub, keyValue, Nothing)
        Return readValue
    End Function

    ''' <summary>
    ''' Sets a KEY VALUE for the registry (HKEY_CURRENT_USER) \ subkey \ keysub \ value
    ''' </summary>
    ''' <param name="subKey"></param>
    ''' <param name="keySub"></param>
    ''' <param name="keyValue"></param>
    ''' <returns></returns>
    Public Function SetMyKey(subKey As String, keySub As String, keyValue As String)
        My.Computer.Registry.CurrentUser.CreateSubKey(subKey)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\" & subKey, keySub, keyValue)
    End Function

    ''' <summary>
    ''' Calculate Available Fuels and Total Tons.
    ''' </summary>
    ''' <param name="FType">STRING: Fuel Type</param>
    ''' <param name="FLoad">STRING: Fuel Load</param>
    ''' <param name="BSize">INTEGER: Burn Size</param>
    ''' <returns>INTEGER:  Total Tons</returns>
    Public Function GetTotalTons(ByVal FType As String, ByVal FLoad As String, BSize As Integer)
        Dim FN As Integer

        Select Case FType
            Case "Shortleaf Pine with Oak"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 3.0
                    Case Is = "Moderate"
                        FN = BSize * 4.0
                    Case Is = "Heavy"
                        FN = BSize * 4.4
                    Case Else
                        FN = 0
                End Select
            Case "Shortleaf Pine Regeneration"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 2.6
                    Case Is = "Moderate"
                        FN = BSize * 3.8
                    Case Is = "Heavy"
                        FN = BSize * 5.1
                    Case Else
                        FN = 0
                End Select
            Case "Loblolly Pine with Oak"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 6.4
                    Case Is = "Moderate"
                        FN = BSize * 6.8
                    Case Is = "Heavy"
                        FN = BSize * 7.9
                    Case Else
                        FN = 0
                End Select
            Case "Loblolly Pine Regeneration"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 4.4
                    Case Is = "Moderate"
                        FN = BSize * 7.6
                    Case Is = "Heavy"
                        FN = BSize * 8.5
                    Case Else
                        FN = 0
                End Select
            Case "Hardwood Leaf Litter"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 0.8
                    Case Is = "Moderate"
                        FN = BSize * 1.5
                    Case Is = "Heavy"
                        FN = BSize * 2.5
                    Case Else
                        FN = 0
                End Select
            Case "Grass or Brush"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 2.0
                    Case Is = "Moderate"
                        FN = BSize * 3.0
                    Case Is = "Heavy"
                        FN = BSize * 5.0
                    Case Else
                        FN = 0
                End Select
            Case "Dispersed Slash"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 4.0
                    Case Is = "Moderate"
                        FN = BSize * 6.0
                    Case Is = "Heavy"
                        FN = BSize * 8.0
                    Case Else
                        FN = 0
                End Select
            Case "Piled Debris"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 5.0
                    Case Is = "Moderate"
                        FN = BSize * 7.5
                    Case Is = "Heavy"
                        FN = BSize * 10.0
                    Case Else
                        FN = 0
                End Select
            Case "Shortleaf Loblolly with Grass"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 1.5
                    Case Is = "Moderate"
                        FN = BSize * 3.8
                    Case Is = "Heavy"
                        FN = BSize * 5.9
                    Case Else
                        FN = 0
                End Select
            Case "Corn"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 3.1
                    Case Is = "Moderate"
                        FN = BSize * 4.7
                    Case Is = "Heavy"
                        FN = BSize * 6.2
                    Case Else
                        FN = 0
                End Select
            Case "Cotton"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 0.8
                    Case Is = "Moderate"
                        FN = BSize * 1.1
                    Case Is = "Heavy"
                        FN = BSize * 1.5
                    Case Else
                        FN = 0
                End Select
            Case "Rice"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 2.5
                    Case Is = "Moderate"
                        FN = BSize * 3.7
                    Case Is = "Heavy"
                        FN = BSize * 4.9
                    Case Else
                        FN = 0
                End Select
            Case "Soybean"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 2.9
                    Case Is = "Moderate"
                        FN = BSize * 4.3
                    Case Is = "Heavy"
                        FN = BSize * 5.7
                    Case Else
                        FN = 0
                End Select
            Case "Wheat"
                Select Case FLoad
                    Case Is = "Low"
                        FN = BSize * 0.9
                    Case Is = "Moderate"
                        FN = BSize * 1.4
                    Case Is = "Heavy"
                        FN = BSize * 1.9
                    Case Else
                        FN = 0
                End Select
            Case Else
                Exit Select
        End Select

        Return FN

    End Function

    ''' <summary>
    ''' Calculate Available Fuels ONLY
    ''' </summary>
    ''' <param name="cTypx">STRING: Fuel Type</param>
    ''' <param name="cLoad">STRING: Fuel Load</param>
    ''' <returns>INTEGER: Available Fuels Value</returns>
    Public Function GetAVFUELS(ByVal cTypx As String, ByVal cLoad As String)
        Select Case cTypx
            Case "Shortleaf Pine with Oak"
                Select Case cLoad
                    Case Is = "Low"
                        Return 3.0
                    Case Is = "Moderate"
                        Return 4.0
                    Case Is = "Heavy"
                        Return 4.4
                    Case Else
                        Exit Select
                End Select
            Case "Shortleaf Pine Regeneration"
                Select Case cLoad
                    Case Is = "Low"
                        Return 2.6
                    Case Is = "Moderate"
                        Return 3.8
                    Case Is = "Heavy"
                        Return 5.1
                    Case Else
                        Exit Select
                End Select
            Case "Loblolly Pine with Oak"
                Select Case cLoad
                    Case Is = "Low"
                        Return 6.4
                    Case Is = "Moderate"
                        Return 6.8
                    Case Is = "Heavy"
                        Return 7.9
                    Case Else
                        Exit Select
                End Select
            Case "Loblolly Pine Regeneration"
                Select Case cLoad
                    Case Is = "Low"
                        Return 4.4
                    Case Is = "Moderate"
                        Return 7.6
                    Case Is = "Heavy"
                        Return 8.5
                    Case Else
                        Exit Select
                End Select
            Case "Hardwood Leaf Litter"
                Select Case cLoad
                    Case Is = "Low"
                        Return 0.8
                    Case Is = "Moderate"
                        Return 1.5
                    Case Is = "Heavy"
                        Return 2.5
                    Case Else
                        Exit Select
                End Select
            Case "Grass or Brush"
                Select Case cLoad
                    Case Is = "Low"
                        Return 2.0
                    Case Is = "Moderate"
                        Return 3.0
                    Case Is = "Heavy"
                        Return 5.0
                    Case Else
                        Exit Select
                End Select
            Case "Dispersed Slash"
                Select Case cLoad
                    Case Is = "Low"
                        Return 4.0
                    Case Is = "Moderate"
                        Return 6.0
                    Case Is = "Heavy"
                        Return 8.0
                    Case Else
                        Exit Select
                End Select
            Case "Piled Debris"
                Select Case cLoad
                    Case Is = "Low"
                        Return 5.0
                    Case Is = "Moderate"
                        Return 7.5
                    Case Is = "Heavy"
                        Return 10.0
                    Case Else
                        Exit Select
                End Select
            Case "Shortleaf Loblolly with Grass"
                Select Case cLoad
                    Case Is = "Low"
                        Return 1.5
                    Case Is = "Moderate"
                        Return 3.8
                    Case Is = "Heavy"
                        Return 5.9
                    Case Else
                        Exit Select
                End Select
            Case "Corn"
                Select Case cLoad
                    Case Is = "Low"
                        Return 3.1
                    Case Is = "Moderate"
                        Return 4.7
                    Case Is = "Heavy"
                        Return 6.2
                    Case Else
                        Exit Select
                End Select
            Case "Cotton"
                Select Case cLoad
                    Case Is = "Low"
                        Return 0.8
                    Case Is = "Moderate"
                        Return 1.1
                    Case Is = "Heavy"
                        Return 1.5
                    Case Else
                        Exit Select
                End Select
            Case "Rice"
                Select Case cLoad
                    Case Is = "Low"
                        Return 2.5
                    Case Is = "Moderate"
                        Return 3.7
                    Case Is = "Heavy"
                        Return 4.9
                    Case Else
                        Exit Select
                End Select
            Case "Soybean"
                Select Case cLoad
                    Case Is = "Low"
                        Return 2.9
                    Case Is = "Moderate"
                        Return 4.3
                    Case Is = "Heavy"
                        Return 5.7
                    Case Else
                        Exit Select
                End Select
            Case "Wheat"
                Select Case cLoad
                    Case Is = "Low"
                        Return 0.9
                    Case Is = "Moderate"
                        Return 1.4
                    Case Is = "Heavy"
                        Return 1.9
                    Case Else
                        Exit Select
                End Select
            Case Else
                Exit Select
        End Select
    End Function

    ''' <summary>
    ''' Smoke Calculation Main Function (20 Mile Distance Fixed)
    ''' </summary>
    ''' <param name="xCatDay">INTEGER: Category Day (must be between 1 - 5)</param>
    ''' <param name="xDistance">DOUBLE: Distance (in miles) to the nearest smoke sensitive target.</param>
    ''' <returns>DOUBLE:  compare this with the total tons to determine if the burn will exceed guidelines.</returns>
    Public Function SmokeCalcFunction(ByRef xCatDay As Integer, ByRef xDistance As Double)
        Select Case xCatDay
            Case 1
                Return 0
            Case 2
                Select Case xDistance
                    Case 0 To 0.19
                        Return 0
                    Case 0.2 To 4.9
                        Return 488
                    Case 5 To 9.9
                        Return 1000
                    Case 10 To 19.9
                        Return 1840
                    Case > 19.9
                        Return 2880
                    Case Else
                        Exit Select
                End Select
            Case 3
                Select Case xDistance
                    Case 0 To 0.19
                        Return 0
                    Case 0.2 To 4.9
                        Return 560
                    Case 5 To 9.9
                        Return 1200
                    Case 10 To 19.9
                        Return 2240
                    Case > 19.9
                        Return 3280
                    Case Else
                        Exit Select
                End Select
            Case 4
                Select Case xDistance
                    Case 0 To 0.19
                        Return 0
                    Case 0.2 To 4.9
                        Return 720
                    Case 5 To 9.9
                        Return 1840
                    Case 10 To 19.9
                        Return 4200
                    Case > 19.9
                        Return 6400
                    Case Else
                        Exit Select
                End Select
            Case 5
                Select Case xDistance
                    Case 0 To 0.19
                        Return 0
                    Case 0.2 To 4.9
                        Return 1280
                    Case 5 To 9.9
                        Return 3200
                    Case 10 To 19.9
                        Return 7200
                    Case > 19.9
                        Return 11600
                    Case Else
                        Exit Select
                End Select
            Case Else
                Exit Select
        End Select
    End Function

    ''' <summary>
    ''' Cardinal Direction Convertor
    ''' </summary>
    ''' <param name="deg">DOUBLE: degree to convert</param>
    ''' <returns>STRING: cardinal direction</returns>
    Function cDir(ByRef deg As Double) As String

        Select Case deg
            Case 348.75 To 360
                Return "N"
            Case 11.25 To 33.75
                Return "NNE"
            Case 33.75 To 56.25
                Return "NE"
            Case 56.25 To 78.75
                Return "ENE"
            Case 78.75 To 101.25
                Return "E"
            Case 101.25 To 123.75
                Return "ESE"
            Case 123.75 To 146.25
                Return "SE"
            Case 146.25 To 168.75
                Return "SSE"
            Case 168.75 To 191.25
                Return "S"
            Case 191.25 To 213.75
                Return "SSW"
            Case 213.75 To 236.25
                Return "SW"
            Case 236.25 To 258.75
                Return "WSW"
            Case 258.75 To 281.25
                Return "W"
            Case 281.25 To 303.75
                Return "WNW"
            Case 303.75 To 326.25
                Return "NW"
            Case 326.25 To 348.75
                Return "NNW"
            Case 0 To 11.25
                Return "N"
        End Select
    End Function

End Class