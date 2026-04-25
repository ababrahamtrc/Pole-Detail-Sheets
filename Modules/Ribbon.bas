Attribute VB_Name = "Ribbon"
Public Sub macro1_1_onAction(ByVal control As IRibbonControl)
    Call NJUNS.createNJUNS
End Sub

Public Sub macro1_2_onAction(ByVal control As IRibbonControl)
    Call PoleClassHeightChecker.CheckPole
End Sub

Public Sub macro1_3_onAction(ByVal control As IRibbonControl)
    Call ExportPDS.ExportSinglePDS
End Sub

Public Sub macro1_4_onAction(ByVal control As IRibbonControl)
    Call ImportPDS.ImportSinglePDS
End Sub

Public Sub macro1_5_onAction(ByVal control As IRibbonControl)
    Call RemedyGen.calculateProposedMidspans
End Sub

Public Sub macro1_6_onAction(ByVal control As IRibbonControl)
    Call QACheck.QACheckPole
End Sub

Public Sub macro1_7_onAction(ByVal control As IRibbonControl)
    Call RemedyGen.RemedyGenerator
End Sub

Public Sub macro1_8_onAction(ByVal control As IRibbonControl)
    Call AutoFillForeign.FillForeignPole
End Sub

Public Sub macro1_9_onAction(ByVal control As IRibbonControl)
    Call FixPDS.fixAttachmentHeights
End Sub

Public Sub macro1_10_onAction(ByVal control As IRibbonControl)
    Call FixPDS.fixCommMakeReadyForm
End Sub

Public Sub macro1_11_onAction(ByVal control As IRibbonControl)
    Call CrewNotes.CrewNotesGenerator
End Sub

Public Sub macro1_12_onAction(ByVal control As IRibbonControl)
    Call Figures.getSheetFigures(ThisWorkbook.ActiveSheet)
End Sub

Public Sub macro1_13_onAction(ByVal control As IRibbonControl)
    Call CUExporter.ExportSingleSheetCUs
End Sub

Public Sub macro1_14_onAction(ByVal control As IRibbonControl)
    Call NJUNSGenerateClipboardCode.ExportSingleNJUNS
End Sub

Public Sub macro1_15_onAction(ByVal control As IRibbonControl)
    Call Photos.OpenPolePhoto
End Sub
