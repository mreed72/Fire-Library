' **** SIMONCODER SOFTWARE ****
' You may freely change, use and distribute this code under the following conditions:
' 1. You may NOT charge money for the use of this software or any software that uses this code.
' 2. You must keep this copyright information throughout the code.
' **** © 2018 Scott Reed **** mreed1972@gmail.com ****



Imports System.Windows.Forms

Public Class Class1

    ''' <summary>
    ''' Check Log File Locations
    ''' </summary>
    ''' <param name="DirName">STRING: the directory location.</param>
    ''' <returns>NA</returns>
    Public Function ChkLocations(DirName As String)
        Dim f As Boolean
        f = My.Computer.FileSystem.DirectoryExists(DirName)
        If f = False Then
            My.Computer.FileSystem.CreateDirectory(DirName)
            My.Computer.FileSystem.WriteAllText(DirName & "\elog.txt", "", False)
            My.Computer.FileSystem.WriteAllText(DirName & "\slog.txt", "", False)
        End If

        Dim e, s As Boolean
        e = My.Computer.FileSystem.FileExists(DirName & "\elog.txt")
        s = My.Computer.FileSystem.FileExists(DirName & "\slog.txt")
        If e = False Then
            My.Computer.FileSystem.WriteAllText(DirName & "\elog.txt", "", False)
        End If
        If s = False Then
            My.Computer.FileSystem.WriteAllText(DirName & "\slog.txt", "", False)
        End If
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
    ''' <param name="x4">Location of the text file</param>
    ''' <returns>random word from a text file.</returns>
    Public Function getword(x4 As String)
        Randomize()

        Dim sr As System.IO.StreamReader
        Dim ri As Integer = 0
        Dim wa As New ArrayList

        If System.IO.File.Exists(x4) = True Then 'x4 is location of text file
            sr = New IO.StreamReader(x4)

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
    ''' <param name="cnty">County Name (without county. ex:  ASHLEY   /   not ASHLEY COUNTY)</param>
    ''' <returns>INTEGER that represents the AFC district 1-8  if zero(0) then error</returns>
    Public Function CountyToDistrict(cnty As String)
        Select Case cnty
            Case "Ashley" Or "Bradley" Or "Calhoun" Or "Cleveland" Or "Drew" Or "Lincoln" Or "Desha" Or "Chicot"
                Return 1
            Case "Calhoun" Or "Union" Or "Columbia" Or "Ouachita" Or "Nevada" Or "Lafayette" Or "Hempstead" Or "Miller"
                Return 4
            Case "Little River" Or "Howard" Or "Sevier" Or "Pike" Or "Polk" Or "Montgomery" Or "Scott" Or "Yell"
                Return 2
            Case "Benton" Or "Carroll" Or "Boone" Or "Washington" Or "Madison" Or "Newton" Or "Crawford" Or "Franklin" Or "Johnson" Or "Pope" Or "Sebastian" Or "Logan"
                Return 6
            Case "Pulaski" Or "Perry" Or "Faulkner" Or "White" Or "Conway" Or "Van Buren" Or "Cleburne"
                Return 7
            Case "Baxter" Or "Marion" Or "Searcy" Or "Stone" Or "Fulton" Or "Izard" Or "Sharp" Or "Randolph" Or "Lawrence" Or "Independence"
                Return 8
            Case "Clay" Or "Greene" Or "Craighead" Or "Mississippi" Or "Jackson" Or "Poinsett" Or "Crittenden" Or "Cross" Or "Woodruff" Or "St Francis" Or "Lee" Or "Phillips" Or "Monroe" Or "Arkansas" Or "Prairie" Or "Lonoke"
                Return 3
            Case Else
                Return 0
        End Select
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
                End Select
            Case Else
                Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
                End Select
            Case 70 To 74
                Select Case xDispersionIndex
                    Case 1
                        Return 4
                    Case Is > 2
                        Return 3
                    Case Else
                        Return 0
                End Select
            Case 75 To 79
                Select Case xDispersionIndex
                    Case 1 To 16
                        Return 4
                    Case Is > 17
                        Return 3
                    Case Else
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
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
                        Return 0
                End Select
            Case Else
                Return 0
        End Select
    End Function

End Class


Public Class BETA
    Public Function cypt()

        Dim a1, a2 As Integer
        Dim va, vb, vc, vd, ve As Integer
        Do Until Val(va) + Val(vb) + Val(vc) + Val(vd) + Val(ve) = 10
            va = CInt(Int((9 * Rnd()) + 1))
            vb = CInt(Int((9 * Rnd()) + 1))
            vc = CInt(Int((9 * Rnd()) + 1))
            vd = CInt(Int((9 * Rnd()) + 1))
            ve = CInt(Int((9 * Rnd()) + 1))

            a1 = va & vb & vc & vd & ve
        Loop

        Dim v1, v2, v3, v4, v5 As Integer
        Do Until Val(v1) + Val(v2) + Val(v3) + Val(v4) + Val(v5) = 15
            v1 = CInt(Int((9 * Rnd()) + 1))
            v2 = CInt(Int((9 * Rnd()) + 1))
            v3 = CInt(Int((9 * Rnd()) + 1))
            v4 = CInt(Int((9 * Rnd()) + 1))
            v5 = CInt(Int((9 * Rnd()) + 1))

            a2 = v1 & v2 & v3 & v4 & v5
        Loop

        Return a1 & "-" & a2 & "-" & grs(5)
    End Function

    Public Function grs(ByRef length As Integer) As String
        Randomize()
        Dim ac As String
        ac = "123456789"
        Dim i As Integer
        For i = 1 To length
            grs = grs & Mid(ac, Int(Rnd() * Len(ac) + 1), 1)
        Next
    End Function


    Public Function Dcypt(rkey As String)
        Dim total As Integer
        Dim num() As String
        num = Split(rkey, "-")
        total = 0
        For Each n In num
            total = total + Val(num)
        Next
        Return total
    End Function







End Class
