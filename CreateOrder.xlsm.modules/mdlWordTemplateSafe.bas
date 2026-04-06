Attribute VB_Name = "mdlWordTemplateSafe"
Option Explicit

Private Const WORD_TEXT_CHUNK_SIZE As Long = 230
Private Const WD_COLLAPSE_END As Long = 0
Private Const WD_FIND_STOP As Long = 0

Public Sub ReplacePlaceholderText(ByVal wordDocument As Object, ByVal placeholder As String, ByVal replacementValue As Variant)
    Dim searchRange As Object

    If wordDocument Is Nothing Then Exit Sub

    Set searchRange = wordDocument.Content
    ReplacePlaceholderTextInRange searchRange, placeholder, replacementValue
End Sub

Private Sub ReplacePlaceholderTextInRange(ByVal searchRange As Object, ByVal placeholder As String, ByVal replacementValue As Variant)
    Dim replacementText As String

    If searchRange Is Nothing Then Exit Sub

    replacementText = NormalizeReplacementValue(replacementValue)

    With searchRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = placeholder
        .Forward = True
        .Wrap = WD_FIND_STOP

        Do While .Execute
            WriteRangeTextSafe searchRange, replacementText
            searchRange.Collapse WD_COLLAPSE_END
        Loop
    End With
End Sub

Private Sub WriteRangeTextSafe(ByVal targetRange As Object, ByVal replacementText As String)
    Dim startPos As Long
    Dim currentChunk As String

    targetRange.Text = vbNullString

    If Len(replacementText) = 0 Then Exit Sub

    For startPos = 1 To Len(replacementText) Step WORD_TEXT_CHUNK_SIZE
        currentChunk = Mid$(replacementText, startPos, WORD_TEXT_CHUNK_SIZE)
        targetRange.InsertAfter currentChunk
        targetRange.Collapse WD_COLLAPSE_END
    Next startPos
End Sub

Private Function NormalizeReplacementValue(ByVal replacementValue As Variant) As String
    If IsError(replacementValue) Then Exit Function
    If IsNull(replacementValue) Then Exit Function
    If IsEmpty(replacementValue) Then Exit Function

    NormalizeReplacementValue = CStr(replacementValue)
End Function
