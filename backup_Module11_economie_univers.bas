Attribute VB_Name = "Module11"

Sub OLD_Import_Cognos_Economie_Univers()
On Error GoTo 0

'  Import_Cognos__Economie_Univers Macro
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'  This macro asks the user for Cognos report # FCA-OPE-FC-0200-02, imports it
'  and updates the related pivot table in order to prepare journal entries to reclass cost containment fees paid through Univers in the correct entities and accounts.
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'  Parts of this code were obtained from various non-copyrighted sources.
'  Coded by Gabriel Lainesse. gabriel.lainesse@qc.croixbleue.ca
'  This work is not protected by Copyright. Feel free to use it, or parts of it, but only for Croix Bleue related work.
'  By using this macro, you agree that you will not hold me (Gabriel Lainesse) responsible for any harm that
'  occurs as a result of your use of the macro.
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

'Checks if Macro is run from the correct workbook
Dim vMsgBox As Integer

If Not ActiveWorkbook.Sheets(1).Range("A1") = "RECLASSEMENT DES FRAIS D'�CONOMIE PAY�S VIA UNIVERS" Then
    
    vMsgBox = MsgBox("Vous devez ex�cuter la macro depuis le classeur d'analyse de reclassement des frais d'�conomie pay�s via Univers!" & vbNewLine & _
    "Assurez-vous �galement que la feuille SETUP soit la" & vbNewLine & "premi�re feuille du classeur et que le titre" & vbNewLine & _
    "'RECLASSEMENT DES FRAIS D'�CONOMIE PAY�S VIA UNIVERS' soit inscrit correctement dans la cellule A1.", _
        vbExclamation, "Mauvais classeur pour cette macro")
    Exit Sub
End If

' Prevents from running the macro without user consent, and thus prevent losing data.
vMsgBox = MsgBox("Voulez-vous vraiment lancer la proc�dure de mise � jour?", vbYesNo + vbQuestion + vbDefaultButton2, "Attention!")

If Not vMsgBox = 6 Then
    Worksheets("SETUP").Select
    vMsgBox = MsgBox("La mise � jour a �t� annul�e", vbOKOnly, "Mise � jour annul�e")
    Exit Sub
End If

'Asks if a copy of the previous month has been done
vMsgBox = MsgBox("Les anciennes donn�es seront �cras�es et ce de fa�on d�finitive." _
    & vbNewLine & "Assurez-vous d'avoir fait une copie du classeur avant de lancer la" & vbNewLine & "mise � jour." _
    & vbNewLine & vbNewLine & "Avez-vous d�j� fait une copie du classeur du mois pr�c�dent?", vbYesNo + vbQuestion + vbDefaultButton2, "Avez-vous fait une copie?")

If Not vMsgBox = 6 Then
    Worksheets("SETUP").Select
    vMsgBox = MsgBox("Veuillez faire une copie des donn�es ant�rieures avant de d�marrer cette macro.", vbOKOnly, "Mise � jour annul�e")
    Exit Sub
End If

'Checks if data is missing in SETUP sheet
If IsEmpty(ActiveWorkbook.Sheets(1).Range("MOISN")) Or IsEmpty(ActiveWorkbook.Sheets(1).Range("AN")) Or IsEmpty(ActiveWorkbook.Sheets(1).Range("SIGNATURE")) Then
    vMsgBox = MsgBox("Des donn�es sont manquantes dans la feuille SETUP." & vbNewLine & "La macro ne peut donc pas continuer." & vbNewLine _
        & "Veuillez remplir les champs orang�s de la feuille SETUP" & vbNewLine & "avant de proc�der � l'ex�cution de cette macro.", vbExclamation, "Donn�es manquantes")
    Exit Sub
End If


'Condition non utilis�e dans ce classeur
' Or IsEmpty(ActiveWorkbook.Sheets(1).Range("ECONOMIE")) Or IsEmpty(ActiveWorkbook.Sheets(1).Range("TAUXCHANGE")) Or IsEmpty(ActiveWorkbook.Sheets(1).Range("FRAISUS"))


' --------------------------------------------------------------------------------------------------------------------------------------------------
' PRE RUN CHECK COMPLETED
' --------------------------------------------------------------------------------------------------------------------------------------------------

'Storing current month in a variable
Dim vCurrentMonth As String
vCurrentMonth = Worksheets("SETUP").Range("MOIST").Value

'Storing current year in a variable
Dim vCurrentYear As String
vCurrentYear = Worksheets("SETUP").Range("AN").Value

' Declaring variables for Cognos extraction and copy-pasting into the analysis workbook.
Dim vFile As Variant
Dim wbCopyTo As Workbook
Dim wsCopyTo As Worksheet
Dim wbCopyFrom As Workbook
Dim wsCopyFrom As Worksheet
Dim vLastRow As Variant
Dim vLastRowNum As Integer
Dim vLastColumn As Variant
Dim vFirstPastedRow As Variant
Dim vFirstPastedColumn As Variant
Dim numRows As Variant
Dim numColumns As Variant

Dim vRangeCheck As Range
Dim vReportDataRange As Variant
Dim vReport2DataRange As Variant
Dim vPivotDataTrim As Variant
Dim vPivotDataFinal As Variant

Dim vWSCopyFromCount As Integer
Dim vWSCopyFromCount2 As Integer
Dim vWSCopyFrom As Worksheet
Dim vWSCopyToCount As Integer
Dim vWSCopyToCount2 As Integer
Dim vWSCopyTo As Worksheet

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
' STARTING IMPORTING PROCESS FOR REPORT 1 : FCA-OPE-FC-0200-02
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'Disabling Filters (NOT USED IN THIS MACRO)
'On Error Resume Next
'Worksheets("FIN-OPE-UN-0010-01").Select
'    If Worksheets("FIN-OPE-UN-0010-01").AutoFilterMode Then
 '
  '      If CurrentSheet.FilterMode Then
   '         Worksheets("FIN-OPE-UN-0010-01").ShowAllData
    '
     '   Else
     '       CurrentSheet.AutoFilterMode = False
     '   End If
    ' Else
   ' End If
On Error GoTo 0

'-------------------------------------------------------------
' Moving to destination sheet for first report

Worksheets("SOMMAIRE FCA-OPE-FC-0200-02").Select

'-------------------------------------------------------------
' Storing the analysis workbook as the destination workbook for the copy-paste process of the report (FCA-OPE-FC-0200-02)
Set wbCopyTo = ActiveWorkbook
Set wsCopyTo = ActiveSheet


'Deletes everything from the previous report summary
Cells.Select
Selection.ClearContents
Selection.ClearFormats

'-------------------------------------------------------------
'Open file with data to be copied
    
    vFile = Application.GetOpenFilename("Excel Files (*.xl*)," & _
    "*.xl*", 1, "S�lectionnez FCA-OPE-FC-0200-02 D�tail �conomie en devise - Univers pour " & vCurrentMonth & " " & " " & vCurrentYear, "Open", False)
    
    'If Cancel then Exit
    If TypeName(vFile) = "Boolean" Then
            vMsgBox = MsgBox("Aucun fichier n'a �t� fourni." & vbNewLine & "Veuillez d�marrer la macro de nouveau" & vbNewLine & "lorsque le fichier sera disponible." _
            , vbExclamation, "Aucun fichier fourni")
            wbCopyTo.Activate
            Worksheets("SETUP").Select
            Exit Sub
        
        Else
            Set wbCopyFrom = Workbooks.Open(vFile)
            Set wsCopyFrom = wbCopyFrom.Worksheets(1)
        End If
    
    If Not wbCopyFrom.Worksheets(1).Range("A4").Value = "D�tail �conomie en devise - Univers" Then
        vMsgBox = MsgBox("Ceci ne semble pas �tre le bon rapport." & vbNewLine & "Assurez-vous de fournir le rapport FCA-OPE-FC-0200-02 � cette �tape." _
                , vbExclamation, "Mauvais rapport fourni")
        wbCopyTo.Activate
        Worksheets("SETUP").Select
        wbCopyFrom.Close SaveChanges:=False
        Exit Sub
    End If
    

' Cr�ation de l'ent�te de la feuille Sommaire
wbCopyFrom.Worksheets(1).Range("A4").Copy
ActiveWorkbook.Worksheets("SOMMAIRE FCA-OPE-FC-0200-02").Range("A4").Paste
Application.CutCopyMode = False
wbCopyFrom.Worksheets(1).Range("A6").Copy
ActiveWorkbook.Worksheets("SOMMAIRE FCA-OPE-FC-0200-02").Range("A6").Paste
Application.CutCopyMode = False
wbCopyFrom.Worksheets(1).Range("A8:M8").Copy
ActiveWorkbook.Worksheets("SOMMAIRE FCA-OPE-FC-0200-02").Range("A6:M8").Paste
Application.CutCopyMode = False

' Cr�ation de l'ent�te de la feuille Totaux
wbCopyFrom.Worksheets(1).Range("A4").Copy
ActiveWorkbook.Worksheets("TOTAUX FCA-OPE-FC-0200-02").Range("A4").Paste
Application.CutCopyMode = False
wbCopyFrom.Worksheets(1).Range("A6").Copy
ActiveWorkbook.Worksheets("TOTAUX FCA-OPE-FC-0200-02").Range("A6").Paste
Application.CutCopyMode = False
wbCopyFrom.Worksheets(1).Range("A8:M8").Copy
ActiveWorkbook.Worksheets("TOTAUX FCA-OPE-FC-0200-02").Range("A6:M8").Paste
Application.CutCopyMode = False


'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

'Processus de copier-coller mis en boucle pour toutes les feuilles du rapport

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
vWSCopyFrom = wbCopyFrom.Worksheets(1)
vLastRowNum = 8
On Error Resume Next

For vWSCopyFromCount = 1 To wbCopyFrom.Worksheets.Count

    'Active la feuille du rapport en fonction de la progression de vWSCopyFromCount
    wbCopyFrom.Activate
    wbCopyFrom.Worksheets(vWSCopyFromCount).Select
    'D�fusionne les cellules fusionn�es dans la feuille actuelle
    Cells.UnMerge
    
    Dim vCellSearch As Range
    Dim vCellSearch2 As Range
    Dim vCellSearchLastValue As String
    
        For Each vCellSearch In Range("A9:A" & ActiveSheet.UsedRows.Count - 2)
            'Observe si la cellule actuelle contient des donn�es ou si elle contient des donn�es identiques � celles pr�c�demment enregistr�es.
            'Si ce n'est pas le cas, mets � jour la variable contenant la derni�re donn�e relev�e dans la colonne.
            If Not vCellSearch.Value Is Nothing Or Not vCellSearch.Value = vCellSearchLastValue Then
                vCellSearchLastValue = vCellSearch.Value
            End If
            'Si c'est le cas, inscrit la valeur de la derni�re cellule avec des donn�es dans cette cellule.
            If vCellSearch.Value Is Nothing Then
                vCellSearch.Value = vCellSearchLastValue
            End If
    
        Next vCellSearch

        For Each vCellSearch2 In Range("B9:B" & ActiveSheet.UsedRows.Count - 2)
            'Observe si la cellule actuelle contient des donn�es ou si elle contient des donn�es identiques � celles pr�c�demment enregistr�es.
            'Si ce n'est pas le cas, mets � jour la variable contenant la derni�re donn�e relev�e dans la colonne.
            If Not vCellSearch2.Value Is Nothing Or Not vCellSearch2.Value = vCellSearchLastValue Then
                vCellSearchLastValue = vCellSearch2.Value
            End If
            'Si c'est le cas, inscrit la valeur de la derni�re cellule avec des donn�es dans cette cellule.
            If vCellSearch2.Value Is Nothing Then
                vCellSearch2.Value = vCellSearchLastValue
            End If
    
        Next vCellSearch2

'--------------------------------------------------------------------------
'Copie les donn�es � la fin de la feuille SOMMAIRE FCA-OPE-FC-0200-02
'--------------------------------------------------------------------------
wbCopyFrom.Activate
wbCopyFrom.Worksheets(vWSCopyFromCount).Select
'Pr�vient le copier-coller s'il n'y a pas de donn�es dans la feuille en cours
If Not ActiveSheet.Range("A9").Value Is Nothing Then
    
ActiveSheet.Range("A9", Cells(Range("A100000").End(xlUp).Row - 2, Range("AAA9").End(xlToLeft).Column)).Select
Selection.Copy

'V�rifie si l'ent�te est la premi�re rang�e utilis�e dans le corps de la feuille sommaire (c'est le cas si rien n'a �t� copi� jusqu'� maintenant)
'Dans le cas contraire, utilise la fonction End(xlDown) (Control + Fl�che bas) pour trouver la derni�re ligne
    If wsCopyTo.Range("A8").Row = vLastRowNum Then
            wsCopyTo.Range("A9").Select
        Else
            wsCopyTo.Range("A9").Select
            Selection.End(xlDown).Offset(1, 0).Select

    End If

'Colle les donn�es copi�es dans la feuille sommaire
Selection.PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False

End If

'--------------------------------------------------------------------------
'Copie les totaux de chaque feuille � la fin de la feuille TOTAUX FCA-OPE-FC-0200-02
'--------------------------------------------------------------------------
'Change l'attribution de wsCopyTo � la feuille Totaux
wbCopyTo.Activate
wsCopyTo = wbCopyTo.Worksheets("TOTAUX FCA-OPE-FC-0200-02")
wsCopyTo.Select


'Remet � 1 l'index des feuilles du rapport FCA-OPE-FC-0200-02 afin de proc�der � une deuxi�me op�ration
vWSCopyFrom = wbCopyFrom.Worksheets(1)
On Error Resume Next
For vWSCopyFromCount2 = 1 To wbCopyFrom.Worksheets.Count

    'Active la feuille du rapport en fonction de la progression de vWSCopyFromCount
    wbCopyFrom.Activate
    wbCopyFrom.Worksheets(vWSCopyFromCount2).Select
    
        Dim vCellSearch3 As Range
        For Each vCellSearch3 In Range("A9:A" & ActiveSheet.UsedRows.Count)
            'Boucle sur les cellules de la colonne A afin de v�rifier laquelle contient un "Total". Lorsqu'une ligne de total est trouv�e, elle est copi�e-coll�e dans
                ' la feuille TOTAUX � la suite des autres.
            If Left(vCellSearch3.Value, 5) = "Total" Then
                    Range(vCellSearch3, Cells(vCellSearch3.Row, Range(Cells(vCellSearch3.Row, 100)).End(xlToLeft))).Select
                    Selection.Copy
                    wbCopyTo.Range("A8").End(xlDown).Offset(1, 0).Paste
            End If
    
        Next vCellSearch3


'Fin de la boucle
Next vWSCopyFromCount2

On Error GoTo 0


'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

'FIN DU PROCESSUS de copier-coller mis en boucle pour toutes les feuilles du rapport

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------



    
    
End Sub
    
    
    

