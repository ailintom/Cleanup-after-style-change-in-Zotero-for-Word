Attribute VB_Name = "StyleChangeCleanup"
' ***** BEGIN LICENSE BLOCK *****
'
'MIT License
'
'Copyright (c) 2017 A. Ilin-Tomich (Johannes Gutenberg University Mainz)
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
' ***** END LICENSE BLOCK *****
Private Const ZoteroFieldIdentifier = " ADDIN ZOTERO_ITEM CSL_CITATION"
Private Const PunctuationPrecedingNoteReference = "[.,:;?!]"
'This macro places references before the punctuation marks and adds spaces before references
Sub CleanUpAfterChangingNotesToAuthorDate()
Dim uUndo As UndoRecord
Set uUndo = Application.UndoRecord
uUndo.StartCustomRecord ("Clean up citations after converting notes to author-date citations") 'Make the macro appear as a single operation on the Undo list
Dim fField As Field, rRange As Range, strPrevChar As String
Dim rPrevChar As Range
Dim rCurrentPosition As Range
Set rCurrentPosition = Selection.Range
For Each fField In ActiveDocument.Fields
    If Left(fField.Code, Len(ZoteroFieldIdentifier)) = ZoteroFieldIdentifier Then 'Check if this is a Zotero item field
        fField.Select
        Set rRange = Selection.Range
        If rRange.Start > ActiveDocument.Range.Start Then
            Set rPrevChar = ActiveDocument.Range(rRange.Start - 1, rRange.Start)
            If rPrevChar.Text Like PunctuationPrecedingNoteReference Then 'If the note reference was preceded by a punctuation sign, move it after the Author-Date citation
                Do While rPrevChar.Start > ActiveDocument.Range.Start
                    If ActiveDocument.Range(rPrevChar.Start - 1, rPrevChar.Start).Text Like PunctuationPrecedingNoteReference Then
                        rPrevChar.Start = rPrevChar.Start - 1 'if there is more than one punctuation sign preceding the reference, place them all after the reference
                    Else
                        Exit Do
                    End If
                Loop
                rRange.InsertAfter rPrevChar.Text
                rPrevChar.Delete
                If rRange.Start > ActiveDocument.Range.Start Then
                    Set rPrevChar = ActiveDocument.Range(rRange.Start - 1, rRange.Start)
                Else
                    GoTo Skip
                End If
            End If
            If Not (rPrevChar.Text = " " Or rPrevChar.Text = ChrW(160)) Then 'Insert a space before the Author-Date citation if there is no space or non-breaking space
                rRange.InsertBefore " "
            End If
        End If
    End If
Skip:
Next fField
rCurrentPosition.Select
uUndo.EndCustomRecord
End Sub

'This macro places footnote/endnote references after the punctuation marks and eliminates spaces before references
Sub CleanUpAfterChangingAuthorDateToNotes()
Dim uUndo As UndoRecord
Set uUndo = Application.UndoRecord
uUndo.StartCustomRecord ("Clean up citations after converting author-date citations to notes") 'Make the macro appear as a single operation on the Undo list
Dim fField As Field, rRange As Range, strPrevChar As String, fFootnote As Footnote, eEndNote As Endnote
Dim rCurrentPosition As Range
Set rCurrentPosition = Selection.Range
For Each fFootnote In ActiveDocument.Footnotes
    If fFootnote.Range.Fields.Count > 0 Then
    For Each fField In fFootnote.Range.Fields
        If Left(fField.Code, Len(ZoteroFieldIdentifier)) = ZoteroFieldIdentifier Then
        Call ProcessReference(fFootnote)
        End If
    Next fField
    End If
Next fFootnote
For Each eEndNote In ActiveDocument.Endnotes
    If eEndNote.Range.Fields.Count > 0 Then
    For Each fField In eEndNote.Range.Fields
        If Left(fField.Code, Len(ZoteroFieldIdentifier)) = ZoteroFieldIdentifier Then
        Call ProcessReference(eEndNote)
        End If
    Next fField
    End If
Next eEndNote
rCurrentPosition.Select
uUndo.EndCustomRecord
End Sub

'Removes a space before the Footnote/Endnote reference and places the punctuation signs (if any) before it
Private Sub ProcessReference(Note)
Dim iStart As Long, iEnd As Long
Dim rPrevChar As Range, rNextChar As Range
iStart = Note.Reference.Start
If iStart > ActiveDocument.Range.Start Then
    Set rPrevChar = ActiveDocument.Range(iStart - 1, iStart)
    If rPrevChar.Text = " " Or rPrevChar.Text = ChrW(160) Then rPrevChar.Delete
End If
iEnd = Note.Reference.End
If iEnd < ActiveDocument.Range.End Then
    Set rNextChar = ActiveDocument.Range(iEnd, iEnd + 1)
    If rNextChar.Text Like PunctuationPrecedingNoteReference Then
        Do While rNextChar.End < ActiveDocument.Range.End
        If ActiveDocument.Range(rNextChar.End, rNextChar.End + 1).Text Like PunctuationPrecedingNoteReference Then
            rNextChar.End = rNextChar.End + 1 'if there is more than one punctuation sign following the note reference, place them all before the reference
        Else
            Exit Do
        End If
        Loop
        Note.Reference.InsertBefore rNextChar.Text
        rNextChar.Delete
    End If
End If
End Sub

