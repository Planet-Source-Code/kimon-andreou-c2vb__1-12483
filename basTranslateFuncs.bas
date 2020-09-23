Attribute VB_Name = "basTranslateFuncs"
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

'Translate a C function
'
'This is the only function that has an error handler.
'This is because, this function handles unknown instructions also.
Function ProcessFunction(CFunction As String) As String
On Error GoTo HandleError

Dim Source As String
Dim dummy As String
Dim pos As Integer
Dim Naming As String
Dim paramList As String
Dim NamePart() As String
Dim Params() As String
Dim FuncName As String
Dim FuncType As String
Dim cnt As Integer
Dim TypePart() As String
Dim ArgTypes() As String
Dim ArgNames() As String
Dim strVal1 As String
Dim strVal2 As String
'Flags
Dim IsProcedure As Boolean
Dim IsByVal As Boolean
Dim IsArray As Boolean



IsProcedure = False
Source = Trim(CFunction)                        'Trim spaces
Source = AddSpaces(Source)                      'Add in-between spaces

If Len(Source) <= 1 Then    'If empty, exit
    ProcessFunction = ""
    Exit Function
End If

'Get the function name and type definition
pos = InStr(1, Source, "(")
Naming = Left(Source, pos)
Naming = Trim(Naming)
paramList = Right(Source, Len(Source) - pos)
paramList = Trim(paramList)
NamePart = GetToken(Naming, 1, dummy, "(")
Naming = Join(NamePart, " ")
Naming = Trim(Naming)
NamePart = GetToken(Naming, 1, dummy, vbCr)
Naming = Trim(Join(NamePart, " "))
NamePart = GetToken(Naming, 1, dummy, vbLf)
Naming = Trim(Join(NamePart, " "))
NamePart = GetToken(Naming, 1, dummy)

'Get the argument list
Params = GetToken(paramList, 1, dummy, "(")
paramList = Trim(Join(Params, " "))
Params = GetToken(paramList, 1, dummy, ";")
paramList = Trim(Join(Params, " "))
Params = GetToken(paramList, 1, dummy, ")")
paramList = Trim(Join(Params, " "))
Params = GetToken(paramList, 1, dummy, vbCr)
paramList = Trim(Join(Params, " "))
Params = GetToken(paramList, 1, dummy, vbLf)
paramList = Trim(Join(Params, " "))
Params = GetToken(paramList, 1, dummy, ",")

'Set the function's name
FuncName = NamePart(UBound(NamePart))

ReDim TypePart(UBound(NamePart) - 1)
For cnt = 0 To UBound(TypePart)
    TypePart(cnt) = NamePart(cnt)
Next cnt

'Determine if we are dealing with a function that returns a pointer/array
If InStr(FuncName, "*") <> 0 Then IsByVal = False
If InStr(FuncName, "**") <> 0 Then IsArray = True
If InStr(FuncName, "[") <> 0 Then
    IsByVal = False
    IsArray = True
End If

'Get rid of the special characters and replace them with spaces
NamePart = GetToken(FuncName, 1, dummy, "*")
FuncName = Trim(Join(NamePart, " "))
NamePart = GetToken(FuncName, 1, dummy, "[")
FuncName = Trim(Join(NamePart, " "))
NamePart = GetToken(FuncName, 1, dummy, "]")
FuncName = Trim(Join(NamePart, " "))

'Determine if the function is really a procedure
FuncType = GetType(TypePart, IsProcedure)

'Set the data type of the function (if any)
FuncType = IIf(((Not IsByVal) Or IsArray) And (FuncType = "Byte"), "String", _
    IIf(IsByVal Or IsArray, "Any", FuncType))

'Resize the arrays that hold the argument names and types
ReDim ArgNames(UBound(Params))
ReDim ArgTypes(UBound(Params))

'Loop through the parameter list
For cnt = 0 To UBound(Params)
    'Get the datatypes and names
    ProcessArg Trim(Params(cnt)), strVal1, strVal2, , Trim(Str(cnt))
    ArgNames(cnt) = strVal1
    ArgTypes(cnt) = strVal2
Next cnt

'Set the declaration of the function
ProcessFunction = IIf(IsPublic, "Public ", "Private ") & "Declare " & _
    IIf(IsProcedure, "Sub ", "Function ") & FuncName & " Lib """ & LibraryName & """" & " ( "

'Add the arguments
ProcessFunction = ProcessFunction & ArgNames(0) & ArgTypes(0)
If UBound(ArgNames) > 0 Then
    For cnt = 1 To UBound(ArgNames)
        ProcessFunction = ProcessFunction & ", " & ArgNames(cnt) & ArgTypes(cnt)
    Next cnt
End If

'Finish the declaration of the function
ProcessFunction = ProcessFunction & " ) " & IIf(IsProcedure, "", " As " & FuncType)
Exit Function

'If we have an error
HandleError:
    ProcessFunction = vbCrLf & "#If False Then" & vbCrLf
    ProcessFunction = ProcessFunction & "'I couldn't handle the following:" & vbCrLf
    ProcessFunction = ProcessFunction & CFunction & vbCrLf
    ProcessFunction = ProcessFunction & "#End If" & vbCrLf
End Function
