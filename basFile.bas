Attribute VB_Name = "basFile"
'    C2VB  Converts C style definitions to VB
'    Copyright (C) 2000  Kimon Andreou (kimon@mindless.com)
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit
Option Base 0

'Split the instructions within a text file
'This function works in EXACTLY the same way as the function
'"SplitInstructions" only that instead of working with an input string
'it works with an input file
'Check the comments in that function
Public Function ProcessFile(fname As String) As Variant
Dim Instructions As Integer
Dim InInstr As Boolean
Dim InSection As Boolean
Dim SectionStarts As Integer
Dim dummy As String * 1
Dim Instrc() As String
Dim hFileIn As Integer
Dim InPreProcessor As Boolean
Dim InCComment As Boolean
Dim CheckingCComment As Boolean
Dim CheckingEndCComment As Boolean

ReDim Instrc(1)
hFileIn = FreeFile

Open fname For Binary Access Read Lock Write As hFileIn

While Not EOF(hFileIn)
        Get hFileIn, , dummy
        
        If dummy <> Chr$(0) Then
        
        If Len(Instrc(Instructions)) = 0 Then
            If (dummy = vbCr) Or (dummy = vbLf) Then dummy = ""
        End If
        Instrc(Instructions) = Instrc(Instructions) & dummy
        Select Case dummy
            Case "/":
                If CheckingCComment Then
                    CheckingCComment = False
                Else
                    If CheckingEndCComment Then
                        InCComment = False
                        CheckingEndCComment = False
                    Else
                        CheckingCComment = True
                    End If
                End If
            Case "*":
                If CheckingCComment Then
                    CheckingCComment = False
                    InCComment = True
                Else
                    If InCComment Then
                        CheckingEndCComment = True
                    End If
                End If
            Case "#":
                If Not InCComment Then InPreProcessor = True
            Case vbLf:
                If Not InCComment Then
                    If InPreProcessor Then
                        InInstr = False
                        Instructions = Instructions + 1
                        ReDim Preserve Instrc(Instructions)
                        InPreProcessor = False
                        Instrc(Instructions) = ""
                    End If
                End If
            Case ";":
                If Not InCComment Then
                    If (Not InSection) And InInstr Then
                        InInstr = False
                        Instructions = Instructions + 1
                        ReDim Preserve Instrc(Instructions)
                        Instrc(Instructions) = ""
                    End If
                End If
            Case "{":
                If Not InCComment Then
                    SectionStarts = SectionStarts + 1
                    InSection = True
                    InInstr = True
                End If
            Case "}":
                If Not InCComment Then
                    SectionStarts = SectionStarts - 1
                    If SectionStarts = 0 Then
                        InInstr = True
                        InSection = False
                    End If
                End If
            Case Else:
                If Not InCComment Then InInstr = True
                CheckingCComment = False
                CheckingEndCComment = False
        End Select
        
        End If
Wend
Close hFileIn

ProcessFile = Instrc
End Function

'Write the translated instructions to a file
'Arguments
'fname:         The filename
'Instructions:  An array of strings containg the VB instructions
Public Sub OutputToFile(fname As String, Instructions() As String)
Dim cnt As Integer
Dim hOut As Integer

'Get the next free file handle
hOut = FreeFile

'Open the file
Open fname For Output Access Write Lock Read Write As hOut

'Print the header to the file
Print #hOut, "Attribute VB_Name=""" & GetFileName(fname) & """"
Print #hOut, "'This file was created by the C2VB convertor, ver " & App.Major & _
    "." & App.Minor & "." & App.Revision
Print #hOut, "'On " & Format(Now(), "dddd, mmm d yyyy") & vbCrLf & vbCrLf
Print #hOut, "Option Explicit"
Print #hOut, "Option Base 0" & vbCrLf

'Write the instructions
For cnt = 0 To UBound(Instructions)
    Print #hOut, Instructions(cnt)
Next cnt

'Close the file
Close hOut

End Sub

Private Function GetFileName(FullName As String) As String
Dim parts() As String
Dim cnt As Integer
Dim dummy As String

parts = Split(FullName, "\")
dummy = parts(UBound(parts))
parts = Split(dummy, ".")

dummy = ""
For cnt = 0 To UBound(parts) - 1
    dummy = dummy & IIf(dummy = "", "", "_") & parts(cnt)
Next cnt
GetFileName = dummy
End Function
