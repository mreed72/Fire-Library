' **** SIMONCODER SOFTWARE ****
' You may freely change, use and distribute this code under the following conditions:
' 1. You may NOT charge money for the use of this software or any software that uses this code.
' 2. You must keep this copyright information throughout the code.
' **** © 2018 Scott Reed **** mreed1972@gmail.com ****



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
    ''' Writes to log file.
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
    ''' Converts county name to AFC district
    ''' </summary>
    ''' <param name="StringLocation">Location of the string to search for county name(ex. textbox.text contains "Cleburne")</param>
    ''' <returns>INTEGER that represents the AFC district 1-8  if zero(0) then error</returns>
    Public Function CountyToDistrict(StringLocation As String)

        Dim cnt As String = StringLocation

        If cnt.Contains("Ashley") Or cnt.Contains("Bradley") Or cnt.Contains("Calhoun") Or cnt.Contains("Chicot") Or cnt.Contains("Cleveland") Or cnt.Contains("Desha") Or cnt.Contains("Drew") Or cnt.Contains("Jefferson") Or cnt.Contains("Lincoln") Then
            Return 1
        ElseIf cnt.Contains("Howard") Or cnt.Contains("Little River") Or cnt.Contains("Montgomery") Or cnt.Contains("Pike") Or cnt.Contains("Polk") Or cnt.Contains("Scott") Or cnt.Contains("Sevier") Or cnt.Contains("Yell") Then
            Return 2
        ElseIf cnt.Contains("Arkansas") Or cnt.Contains("Clay") Or cnt.Contains("Craighead") Or cnt.Contains("Crittenden") Or cnt.Contains("Cross") Or cnt.Contains("Greene") Or cnt.Contains("Jackson") Or cnt.Contains("Lee") Or cnt.Contains("Lonoke") Or cnt.Contains("Mississippi") Or cnt.Contains("Monroe") Or cnt.Contains("Phillips") Or cnt.Contains("Poinsett") Or cnt.Contains("Prairie") Or cnt.Contains("St Francis") Or cnt.Contains("Woodruff") Then
            Return 1
        ElseIf cnt.Contains("Columbia") Or cnt.Contains("Hempstead") Or cnt.Contains("Lafayette") Or cnt.Contains("Miller") Or cnt.Contains("Nevada") Or cnt.Contains("Ouachita") Or cnt.Contains("Union") Then
            Return 4
        ElseIf cnt.Contains("Clark") Or cnt.Contains("Dallas") Or cnt.Contains("Garland") Or cnt.Contains("Grant") Or cnt.Contains("Hot Spring") Or cnt.Contains("Saline") Then
            Return 5
        ElseIf cnt.Contains("Benton") Or cnt.Contains("Boone") Or cnt.Contains("Carroll") Or cnt.Contains("Crawford") Or cnt.Contains("Franklin") Or cnt.Contains("Johnson") Or cnt.Contains("Logan") Or cnt.Contains("Madison") Or cnt.Contains("Newton") Or cnt.Contains("Pope") Or cnt.Contains("Sebastian") Or cnt.Contains("Washington") Then
            Return 6
        ElseIf cnt.Contains("Cleburne") Or cnt.Contains("Conway") Or cnt.Contains("Faulkner") Or cnt.Contains("Perry") Or cnt.Contains("Pulaski") Or cnt.Contains("Van Buren") Or cnt.Contains("White") Then
            Return 7
        ElseIf cnt.Contains("Baxter") Or cnt.Contains("Fulton") Or cnt.Contains("Independence") Or cnt.Contains("Izard") Or cnt.Contains("Lawrence") Or cnt.Contains("Marion") Or cnt.Contains("Randolph") Or cnt.Contains("Searcy") Or cnt.Contains("Sharp") Or cnt.Contains("Stone") Then
            Return 8
        Else
            Return 0
        End If


    End Function

    ''' <summary>
    ''' State Cause Codes converted to Federal Cause Codes
    ''' </summary>
    ''' <param name="StateCode">State Cause Code INTEGER</param>
    ''' <returns>Federal Cause Code INTEGER</returns>
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

End Class



Public Class FuelCalc

    ''' <summary>
    ''' Main Smoke Guidelines Calculation Code
    ''' </summary>
    ''' <param name="xCatDay">INTEGER: 1-5, Category Day</param>
    ''' <param name="xDistance">DOUBLE: Distance to Target</param>
    ''' <returns>INTEGER: represents the total tons allowed for an airshed.</returns>
    Public Function smpCalc(ByRef xCatDay As Integer, ByRef xDistance As Double)
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
                    Case > 20
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
                    Case > 20
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
                    Case > 20
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
                    Case > 20
                        Return 11600
                    Case Else
                        Exit Select
                End Select
            Case Else
                Return 0
        End Select
    End Function

    ''' <summary>
    ''' Available Fuels
    ''' </summary>
    ''' <param name="cTypx">Fuel Type</param>
    ''' <param name="cLoad">Fuel Load</param>
    ''' <returns>DOUBLE: value represents the available fuels for a burn.</returns>
    Public Function GetAvFuels(ByVal cTypx As String, ByVal cLoad As String)
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
    ''' Low Visibility Occurence Risk Index
    ''' </summary>
    ''' <param name="xRelativeHumidity">Relative Humidity</param>
    ''' <param name="xDispersionIndex">Dispersion Index</param>
    ''' <returns>INTEGER: 1-10 that depicts the LVORI value.</returns>
    Public Function LVORI(xRelativeHumidity As Integer, xDispersionIndex As Integer)
        Select Case xRelativeHumidity
            Case 0 To 55
                Select Case xDispersionIndex
                    Case 1 To 30
                        Return 2
                    Case Is > 30
                        Return 1
                    Case Else
                        Exit Select
                End Select
            Case 56 To 59
                Select Case xDispersionIndex
                    Case 1 To 8
                        Return 3
                    Case 9 To 30
                        Return 2
                    Case Is > 31
                        Return 1
                    Case Else
                        Exit Select
                End Select
            Case 60 To 64
                Select Case xDispersionIndex
                    Case 1 To 10
                        Return 3
                    Case 11 To 30
                        Return 2
                    Case Is > 31
                        Return 1
                    Case Else
                        Exit Select
                End Select
            Case 65 To 69
                Select Case xDispersionIndex
                    Case 1
                        Return 4
                    Case 2 To 40
                        Return 2
                    Case Is > 41
                        Return 1
                    Case Else
                        Exit Select
                End Select
            Case 70 To 74
                Select Case xDispersionIndex
                    Case 1
                        Return 4
                    Case Is > 2
                        Return 3
                    Case Else
                        Exit Select
                End Select
            Case 75 To 79
                Select Case xDispersionIndex
                    Case 1 To 16
                        Return 4
                    Case Is > 17
                        Return 3
                    Case Else
                        Exit Select
                End Select
            Case 80 To 82
                Select Case xDispersionIndex
                    Case 1
                        Return 6
                    Case 2 To 4
                        Return 5
                    Case 5 To 16
                        Return 4
                    Case Is > 17
                        Return 3
                    Case Else
                        Exit Select
                End Select
            Case 83 To 85
                Select Case xDispersionIndex
                    Case 1
                        Return 6
                    Case 2 To 6
                        Return 5
                    Case Is > 7
                        Return 4
                    Case Else
                        Exit Select
                End Select
            Case 86 To 88
                Select Case xDispersionIndex
                    Case 1 To 4
                        Return 6
                    Case 5 To 12
                        Return 5
                    Case Is > 13
                        Return 4
                    Case Else
                        Exit Select
                End Select
            Case 89 To 91
                Select Case xDispersionIndex
                    Case 1 To 2
                        Return 7
                    Case 3 To 6
                        Return 6
                    Case 7 To 16
                        Return 5
                    Case Is > 17
                        Return 4
                    Case Else
                        Exit Select
                End Select
            Case 92 To 94
                Select Case xDispersionIndex
                    Case 1
                        Return 8
                    Case 2
                        Return 7
                    Case 3 To 10
                        Return 6
                    Case 11 To 25
                        Return 5
                    Case Is > 26
                        Return 4
                    Case Else
                        Exit Select
                End Select
            Case 95 To 97
                Select Case xDispersionIndex
                    Case 1
                        Return 9
                    Case 2 To 4
                        Return 8
                    Case 5 To 6
                        Return 7
                    Case 7 To 12
                        Return 6
                    Case 13 To 25
                        Return 5
                    Case Is > 26
                        Return 4
                    Case Else
                        Exit Select
                End Select
            Case Is > 97
                Select Case xDispersionIndex
                    Case 1 To 2
                        Return 10
                    Case 3 To 6
                        Return 9
                    Case 7 To 10
                        Return 8
                    Case 11 To 12
                        Return 7
                    Case 13 To 25
                        Return 5
                    Case Is > 26
                        Return 4
                    Case Else
                        Exit Select
                End Select
            Case Else
                Return 0
        End Select
    End Function

End Class


