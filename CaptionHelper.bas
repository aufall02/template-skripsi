' ============================================================
' Macro Caption Helper - Template Skripsi
' Insert caption Gambar/Tabel otomatis dengan nomor mengikuti BAB
' Auto-inserts Gambar/Tabel captions with chapter-linked numbering
' ============================================================

Sub InsertCaptionGambar()
    ' Insert caption "Gambar X.Y Judul" di posisi cursor
    ' Insert "Gambar X.Y Title" caption at cursor position
    Selection.Style = ActiveDocument.Styles("Caption")
    Selection.TypeText Text:="Gambar "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="SEQ Bab \c \* ARABIC", PreserveFormatting:=True
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="SEQ Gambar \s 1 \* ARABIC", PreserveFormatting:=True
    Selection.TypeText Text:=" "
    Selection.Fields.Update
    ' Cursor otomatis berhenti di sini, tinggal ketik judul
    ' Cursor stops here, just type the caption title
End Sub

Sub InsertCaptionTabel()
    ' Insert caption "Tabel X.Y Judul" di posisi cursor
    ' Insert "Tabel X.Y Title" caption at cursor position
    Selection.Style = ActiveDocument.Styles("Caption")
    Selection.TypeText Text:="Tabel "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="SEQ Bab \c \* ARABIC", PreserveFormatting:=True
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="SEQ Tabel \s 1 \* ARABIC", PreserveFormatting:=True
    Selection.TypeText Text:=" "
    Selection.Fields.Update
End Sub


Sub SetupShortcuts()
    CustomizationContext = ActiveDocument.AttachedTemplate

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF8), _
        KeyCategory:=wdKeyCategoryMacro, Command:="InsertCaptionGambar"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyT), _
        KeyCategory:=wdKeyCategoryMacro, Command:="InsertCaptionTabel"
        
End Sub
