'This script converts exponents into superscripts
'Copyright (C) 2022 Yusuke Takahashi
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <https://www.gnu.org/licenses/>.

Attribute VB_Name = "Module1"
Option Explicit

Sub exp2superscript()
    Dim target_range As Range
    Dim default_range As Range
    Set default_range = Selection
    On Error Resume Next
    Set target_range = Application.InputBox(Prompt:="Select target cells (to be overwritten)", Title:="Input for exp2superscript", Default:=default_range.Address, Type:=8)
    If Err <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    target_range.NumberFormatLocal = "@"
    Dim range1 As Range
    Dim exp_text As String
    Dim i As Long
    For Each range1 In target_range
        If IsNumeric(range1) Then
            exp_text = Format(range1, "0.00E-0")
            ' Remove exponent if it is 0
            i = InStr(exp_text, "E0")
            If i > 0 Then
                exp_text = Replace(exp_text, "E0", "")
            End If
            i = InStr(exp_text, "E")
            If i = 0 Then
                range1 = exp_text
            Else
                exp_text = Replace(exp_text, "E", "Å~10", Count:=1)
                range1 = exp_text
                range1.Characters(Start:=i + 3).Font.Superscript = True
            End If
        End If
    Next
End Sub
