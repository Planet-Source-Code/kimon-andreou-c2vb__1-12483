Attribute VB_Name = "basTranslatePreprocessor"
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

'Translate a preprocessor instruction
Public Function ProcessPreProcessor(Cstring As String) As String
Dim parts() As String
Dim dummy As String

'Split into tokens
parts = GetToken(Trim(Cstring), 0, dummy)

'Find out what kind of preprocessor instruction we're dealing with
Select Case parts(0)
    Case "#ifdef":
        parts(0) = "#If"
        ProcessPreProcessor = Join(parts, " ") & " Then"
        
    Case "#endif":
        parts(0) = "#End If"
        ProcessPreProcessor = Join(parts, " ") & vbCrLf
        
    Case "#else":
        parts(0) = "#Else"
        ProcessPreProcessor = Join(parts, " ")
        
    Case "#elseif":
        parts(0) = "#ElseIf"
        ProcessPreProcessor = Join(parts, " ") & " Then"
        
    Case "#ifndef":
        parts(0) = "#If Not"
        ProcessPreProcessor = Join(parts, " ") & " Then"
        
    Case "#define":
        ProcessPreProcessor = ProcessDefine(Trim(Cstring)) & vbCrLf
        
    Case Else:
        ProcessPreProcessor = vbCrLf & vbCrLf & "#If False Then" & vbCrLf & _
            "' Cannot handle the following" & vbCrLf
        ProcessPreProcessor = ProcessPreProcessor & Trim(Cstring) & vbCrLf & _
            "#End If" & vbCrLf
End Select

End Function

'Process a "#define" instruction
'These are tricky, since they can be constant definitions or flags or macros
Private Function ProcessDefine(strDefinition As String) As String
Dim parts() As String
Dim dummy As String

'Find out if we're dealing with a macro or not
If InStr(strDefinition, "(") <> 0 Then
    ProcessDefine = ProcessDefineAsFunction(strDefinition)
Else
    parts = GetToken(strDefinition, 0, dummy)
    parts(0) = IIf(IsPublic, "Public ", "Private ") & "Const"
    parts(1) = parts(1) & " ="
    
    'If the value is a hex number, convert it to VB style notation
    If InStr(parts(UBound(parts)), "0x") <> 0 Then
        parts(UBound(parts)) = Chex2VBhex(Trim(parts(UBound(parts))))
    End If
    
    'If the value is an octal number, convert it to VB style notation
    If InStr(parts(UBound(parts)), "0o") <> 0 Then
        parts(UBound(parts)) = Coct2VBoct(Trim(parts(UBound(parts))))
    End If
    
    dummy = Trim(Join(parts, " "))
    
    'See if we're dealing with a flag
    parts = Split(dummy, "=")
    
    If parts(UBound(parts)) = "" Then   'Is this part of an #ifndef?
        parts(UBound(parts)) = "True"
        dummy = Join(parts, " = ")
        parts = Split(dummy, "Const")
        parts(0) = "#"
        dummy = Join(parts, "Const")
    Else
        'Put it all together
        dummy = Join(parts, " = ")
    End If
    
    ProcessDefine = dummy
End If
End Function

'This function will someday be able to process macro definitions
'Not yet though....
Private Function ProcessDefineAsFunction(strDefine As String) As String
Dim dummy As String

dummy = "#If False Then" & vbCrLf
dummy = dummy & "'The following is a macro defined in C/C++." & vbCrLf
dummy = dummy & strDefine & vbCrLf & "#End If" & vbCrLf
ProcessDefineAsFunction = dummy

End Function
