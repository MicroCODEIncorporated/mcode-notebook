Attribute VB_Name = "NB_Macros"
Option Explicit

Sub NB_LineUp()
'
' NB_LineUp Macro = CTRL-SHIFT-K
' Macro written 06/5/2025 by Timothy J McGuire (ChatGPT)
'
    Dim startPos As Long, endPos As Long
    startPos = Selection.Start
    endPos = Selection.End

    If startPos = 0 Then Exit Sub ' Already at top

    Dim selText As String
    selText = Selection.Text

    Selection.Cut
    Selection.MoveUp Unit:=wdLine, Count:=1
    Dim newStart As Long
    newStart = Selection.Start
    Selection.Paste

    ' Restore selection of pasted content
    Selection.SetRange Start:=newStart, End:=newStart + Len(selText)

End Sub

Sub NB_LineDown()
'
' NB_LineDown Macro = CTRL-SHIFT-K
' Macro written 06/5/2025 by Timothy J McGuire (ChatGPT)
'
    Dim startPos As Long, endPos As Long
    startPos = Selection.Start
    endPos = Selection.End

    Dim selText As String
    selText = Selection.Text

    Selection.Cut
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Dim newStart As Long
    newStart = Selection.Start
    Selection.Paste

    ' Restore selection of pasted content
    Selection.SetRange Start:=newStart, End:=newStart + Len(selText)

End Sub
Sub NB_NewDay()
'
' NB_NewDay Macro = CTRL-SHIFT-K
' Macro written 03/11/2012 by Timothy J McGuire
'
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    Selection.Font.Color = wdColorBlue
    Selection.TypeText Text:="Date: "
    Selection.Font.Color = wdColorBlack
    Selection.InsertDateTime DateTimeFormat:="MMMM dd, yyyy, DDDD", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:=" (LOCATION) "
    Selection.Font.Bold = False
    Selection.Font.Size = 10
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    Selection.TypeParagraph
    
    NB_PunchIn
    
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = False
    Selection.Font.Size = 10
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    Selection.Font.Color = wdColorBlue
    Selection.Font.Bold = True
    Selection.TypeText Text:="TASKs…"
    Selection.Font.Bold = False
    Selection.Font.Color = wdColorBlack
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText

    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    Selection.Font.Color = wdColorBlue
    Selection.Font.Bold = True
    Selection.TypeText Text:="MEETINGs…"
    Selection.Font.Bold = False
    Selection.Font.Color = wdColorBlack
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    Selection.Font.Color = wdColorBlue
    Selection.Font.Bold = True
    Selection.TypeText Text:="INTERRUPTIONs…"
    Selection.Font.Bold = False
    Selection.Font.Color = wdColorBlack
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    Selection.Font.Color = wdColorBlue
    Selection.Font.Bold = True
    Selection.TypeText Text:="NOTEs…"
    Selection.Font.Bold = False
    Selection.Font.Color = wdColorBlack
    Selection.TypeParagraph
    Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    
End Sub
Sub NB_NoChecking()
'
' NB_NoChecking Macro = 'Red X' on Toolbar or CTRL-SHIFT-X
'
'
    Selection.LanguageID = wdEnglishUS
    Selection.NoProofing = True
    Application.CheckLanguage = False
    
End Sub

Sub NB_PunchIn()
'
' NB_PunchIn Macro = CTRL-SHIFT-I
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Color = wdColorBlue
    Selection.TypeText Text:="Start Time"
    Selection.Font.Color = wdColorBlack
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = False
    Selection.TypeParagraph
    Selection.TypeParagraph
    
End Sub

Sub NB_PunchOut()
'
' NB_PunchOut Macro = CTRL-SHIFT-O
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Color = wdColorBlue
    Selection.TypeText Text:="Stop Time"
    Selection.Font.Color = wdColorBlack
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = False
    Selection.TypeParagraph
    Selection.TypeParagraph
    
End Sub

Sub NB_Task()
'
' NB_Task Macro = CTRL-SHIFT-T
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorRed
    Selection.TypeText Text:="Task"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True
        
End Sub

Sub NB_Note()
'
' NB_Note Macro
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorYellow
    Selection.TypeText Text:="Note"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True
        
End Sub

Sub NB_Family()
'
' NB_Family Macro = CTRL-SHIFT-F
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorPink
    Selection.TypeText Text:="Family"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True
    
End Sub

Sub NB_Personal()
'
' NB_Family Macro = CTRL-SHIFT-L
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorPink
    Selection.TypeText Text:="Personal"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True
    
End Sub

Sub NB_Meeting()
'
' NB_Meeting Macro = CTRL-SHIFT-M
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorBlue
    Selection.Font.Color = wdColorWhite
    Selection.TypeText Text:="Meeting"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Color = wdColorBlack
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True
    
End Sub

Sub NB_Support()
'
' NB_Support Macro = CTRL-SHIFT-P
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorTurquoise
    Selection.TypeText Text:="Support"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True

End Sub

Sub NB_PhoneCall()
'
' NB_PhoneCall Macro = CTRL-SHIFT-V
' Macro written 10/30/2007 by Timothy J McGuire
'
    Selection.Font.Size = 10
    Selection.Font.Name = "Helvetica"
    Selection.Font.Bold = True
    Selection.Font.Shading.BackgroundPatternColor = wdColorYellow
    Selection.TypeText Text:="Phone Call"
    Selection.Font.Shading.BackgroundPatternColor = wdColorWhite
    Selection.Font.Bold = True
    Selection.TypeText Text:=":"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" ["
    Selection.Font.Bold = True
    Selection.InsertDateTime DateTimeFormat:="h:mm:ss am/pm", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.Font.Bold = False
    Selection.InsertDateTime DateTimeFormat:=" dd-MMM-yyyy", InsertAsField:= _
    False, DateLanguage:=wdEnglishUS, CalendarType:=wdCalendarWestern, _
    InsertAsFullWidth:=False
    Selection.TypeText Text:="] "
    Selection.Font.Bold = True

End Sub

Sub NBCodeStyle()
'
' NBCodeStyle Macro
'
'
    ' don't check spelling
    Selection.LanguageID = wdEnglishUS
    Selection.NoProofing = True
    Application.CheckLanguage = False
    
    ' gray background
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.BackgroundPatternColor = 15921906
    Selection.Font.Shading.BackgroundPatternColor = wdColorAutomatic
    
    ' coding font
    Selection.Font.Name = "ProFontWindows"
    Selection.Font.Size = 8
    Selection.Font.Bold = False
    
    ' light purple outline
    With Selection.ParagraphFormat
    
        .OutlineLevel = wdOutlineLevelBodyText
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleDot
            .LineWidth = wdLineWidth050pt
            .Color = 10498160
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleDot
            .LineWidth = wdLineWidth050pt
            .Color = 10498160
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDot
            .LineWidth = wdLineWidth050pt
            .Color = 10498160
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDot
            .LineWidth = wdLineWidth050pt
            .Color = 10498160
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
   End With
   With Options
        .DefaultBorderLineStyle = wdLineStyleDot
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = 10498160
    End With
    
End Sub

Sub NBPasteText()
'
' NBPasteText Macro
'
'
    Selection.PasteAndFormat (wdFormatPlainText)
    
End Sub


