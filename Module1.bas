Attribute VB_Name = "Module1"
Option Explicit

Function CognosMerger(ByVal hasFooter As Boolean, ByVal copyTotals As Boolean, ByRef cognosWB As Workbook) As Worksheet
'Variables de classeurs, feuilles et plages
Dim cognosWS As Worksheet ' Variable pour chacune des feuilles originales du classeur
Dim mCell As Range  ' Variable de loop sur les cellules
Dim mCell2 As Range ' Variable de loop sur les cellules 2
Dim mCell3 As Range ' Variable de loop sur les cellules 3

'Variables pour la cr�ation de la feuille sommaire
Dim sheetCreationIteration As Integer 'Nombre de tentatives de cr�ation d'une feuille sommaire (afin de limiter la boucle s'il est impossible d'en cr�er une)
Dim summaryWS As Worksheet ' Feuille sommaire recevant les donn�es des feuilles
Dim worksheetNameCheck As Boolean 'Sert � v�rifier si le nom utilis� pour cr�er la feuille sommaire est d�j� utilis�

'Variables pour la recherche d'ent�te
Dim headerRow As Integer 'Stocke la rang�e de l'ent�te lorsque trouv�e
Dim headerRowVector As Integer 'Variable mesurant le nombre de cellules recherch�es pour trouver un ent�te (et arr�ter la recherche si le nombre limite est d�pass�)

'Variable de statut d'op�ration
Dim headerFound As Boolean 'L'ent�te a �t� trouv� pour cette feuille
Dim headerCopiedToSummary As Boolean 'L'ent�te a �t� copi� dans la feuille sommaire
Dim parsingDone As Boolean 'La recopie de l'ent�te a �t� faite dans les colonnes du d�but pour cette feuille

'Variables pour les op�rations de copie
Dim headerRowContent As String 'Stocke, pour chaque ligne de l'ent�te, le contenu de toutes les cellules de cette ligne afin de les recopier dans une colonne
Dim headerFirstColumn As Integer
Dim lastRow As Integer ' Derni�re rang�e trouv�e dans cette feuille
Dim lastColumn As Integer 'derni�re colonne trouv�e dans cette feuille
Dim summaryWSRowVector As Integer 'indique quelle rang�e, dans la feuille sommaire, sera la suivante � recevoir une copie d'une ligne d'une autre feuille

'Initialisation des variables de comptage et boolean de statut
sheetCreationIteration = 1
summaryWSRowVector = 2 '2 car la ligne 1 est r�serv�e � l'ent�te, lorsque celle-ci est trouv�e
headerCopiedToSummary = False
worksheetNameCheck = False
      
Set cognosWB = ActiveWorkbook
Set cognosWS = cognosWB.Worksheets(1)


Application.ScreenUpdating = False

'UserForm
CognosMergerForm.currentWorkbookLabel.Caption = cognosWB.Name
CognosMergerForm.currentSheetLabel.Caption = "En attente"
CognosMergerForm.currentCellLabel.Caption = "En attente"
CognosMergerForm.currentTaskLabel.Caption = "Cr�ation de la feuille sommaire..."
CognosMergerForm.Show False
CognosMergerForm.Repaint

'============================
'Cr�ation de la feuille Sommaire
cognosWB.Worksheets.Add Before:=Worksheets(1)
Set summaryWS = cognosWB.Sheets(1)
On Error GoTo createSheet
Application.DisplayAlerts = False

cognosWB.Sheets(1).Name = "COGNOSMERGER"
GoTo createSheetDone

'============================
'Si erreur dans l'action de nommer la feuille, essayer avec un suffixe num�ral it�ratif.
createSheet:
On Error GoTo createSheet

If sheetCreationIteration > 999 Then
    On Error GoTo 0
    Application.DisplayAlerts = True
    Call MsgBox("Le nombre maximal de feuilles CognosMerger a �t� atteint (999)." & vbNewLine & _
                "Veuillez en supprimer avant d'en cr�er des nouvelles, ou changez leurs noms", _
                vbCritical + vbOKOnly, "Trop de feuilles CognosMerger")
    End
    Else
End If

'Check si le nom de la feuille existe d�j�.
For Each cognosWS In cognosWB.Worksheets
    If cognosWS.Name = "COGNOSMERGER" & CStr(sheetCreationIteration) Or worksheetNameCheck = True Then
        worksheetNameCheck = True
        Exit For
    Else
    End If
Next cognosWS

'Si la feuille n'existe pas, cr�e la feuille; si elle existe, passe au prochain chiffre suffixe � 'COGNOSMERGER'.
If worksheetNameCheck = True Then
    sheetCreationIteration = sheetCreationIteration + 1
    worksheetNameCheck = False
    GoTo createSheet
    
    Else
    summaryWS.Name = "COGNOSMERGER" & CStr(sheetCreationIteration)
    
End If
    
'============================
createSheetDone:
Debug.Print "Feuille " & summaryWS.Name & " cr��e."
CognosMergerForm.currentTaskLabel.Caption = "Cr�ation de la feuille sommaire r�ussie!"
CognosMergerForm.Repaint

On Error GoTo 0
Application.DisplayAlerts = True

'============================
For Each cognosWS In cognosWB.Worksheets
    'Mise � 0 des variables uniques � chaque feuille
    headerRow = 0
    headerRowVector = 1
    lastRow = 0
    lastColumn = 0
    headerFound = False
    parsingDone = False

    'UserForm
    CognosMergerForm.currentSheetLabel.Caption = cognosWS.Name
    CognosMergerForm.currentTaskLabel.Caption = "S�lection de la prochaine feuille..."
    CognosMergerForm.Repaint

  
  
  'Test pour savoir si la feuille est l�gitime (contient des donn�es d'un rapport Cognos)
    If InStr(1, cognosWS.Name, "COGNOSMERGER", vbTextCompare) <> 0 Then
        Debug.Print "Feuille sommaire saut�e, car ne fait pas partie du rapport Cognos"
        GoTo nextWorksheet
    Else
    End If
    cognosWS.Activate
    
    'Check si la feuille est vide
    If cognosWS.UsedRange.Count = 1 Then
        Debug.Print "Feuille saut�e, car elle semble vide."
        GoTo nextWorksheet
    Else
        'Continue
    End If
    
    'D�fusionne toutes les cellules (DEVRAIT INT�GRER POWERDEFUSER)
    'UserForm
    CognosMergerForm.currentTaskLabel.Caption = "D�fusion des cellules..."
    CognosMergerForm.Repaint

    Cells.UnMerge
    
    
    'Trouve la derni�re ligne et la stocke dans une variable (utile lors de la copie des lignes une par une).
    lastRow = cognosWS.UsedRange.Rows.Count
    lastColumn = cognosWS.UsedRange.Columns.Count
    
'============================
'Trouve la premi�re rang�e de l'ent�te � l'aide des couleurs par d�faut des cellules d'ent�te des rapports Cognos.
selectNextHeaderStep:
            
        'Si l'ent�te a d�j� �t� copi� et que l'ent�te a d�j� �t� 'pars� dans les colonnes du d�but alors skip vers la copie des lignes
        Select Case True
          
        Case Not headerFound
            
             'Si l 'ent�te n'est pas trouv� apr�s 25 colonnes, demander si le programme doit continuer avec la prochaine feuille.
                      
            
                    'Trouve l'ent�te
                    
                For Each mCell In cognosWS.Range(cognosWS.Cells(1, headerRowVector), cognosWS.Cells(lastRow, headerRowVector))
                    CognosMergerForm.currentTaskLabel.Caption = "Recherche de l'ent�te..."
                    CognosMergerForm.currentCellLabel.Caption = mCell.Address
                    CognosMergerForm.Repaint

                    
                        If mCell.Interior.Color = 14865087 And mCell.Borders.Color = 11832160 Then
                            'UserForm
                            CognosMergerForm.currentTaskLabel.Caption = "Ent�te trouv�!"
                            CognosMergerForm.Repaint
                            headerRow = mCell.Row
                            headerFirstColumn = mCell.Column
                            headerFound = True
                            GoTo selectNextHeaderStep
                        Else
                        
                             If Not headerFound And headerRow = 0 And headerRowVector < 25 Then
                                headerRowVector = headerRowVector + 1
                                'incr�mente la variable de test et continue de chercher une ent�te dans la colonne suivante
                            ElseIf Not headerFound And headerRowVector > 25 Then
                                vMsgBox = MsgBox("L'ent�te n'a pas �t� trouv� pour la feuille " & cognosWS.Name & vbNewLine & "Voulez-vous continuer avec la prochaine feuille?", vbYesNo, "Ent�te non trouv�")
                                If vMsgBox = 6 Then
                                    GoTo nextWorksheet
                                ElseIf vMsgBox = 7 Then
                                    End
                                Else
                                    End
                                End If
                            Else
                            End If
                        End If
                    
                Next mCell
                GoTo nextWorksheet
         Case headerFound And Not parsingDone
            'PARSING PROCEDURE (recopie l'ent�te dans des colonnes)
            For Each mCell In cognosWS.Range(Cells(1, headerFirstColumn), Cells(headerRow - 1, headerFirstColumn))
                    CognosMergerForm.currentTaskLabel.Caption = "Cr�ation de colonnes pour les donn�es stock�s dans les ent�tes..."
                    CognosMergerForm.Repaint

                    'Ins�re une nouvelle colonne pour stocker les informations contenues dans l'ent�te.
                    cognosWS.Columns(headerFirstColumn).Insert
                    'TR�S IMPORTANT --> \/  \/  \/
                    headerFirstColumn = headerFirstColumn + 1 'Mets � jour la premi�re colonne des donn�es originales, puisqu'on vient d'en ins�rer une autre
                    lastColumn = lastColumn + 1         ' Et fait de m�me avec la variable contenant le num�ro de la derni�re colonne
                    cognosWS.Cells(headerRow, headerFirstColumn - 1).Value = "Ent�te, ligne " & CStr(mCell.Row)
                    
                    For Each mCell2 In cognosWS.Range(Cells(mCell.Row, headerFirstColumn), Cells(mCell.Row, lastColumn))
                        CognosMergerForm.currentTaskLabel.Caption = "Stockage des donn�es contenues dans l'ent�te..."
                        CognosMergerForm.currentCellLabel.Caption = mCell2.Address
                        CognosMergerForm.Repaint

                        'Concat�ne les informations de la ligne d'ent�te en cours.
                        If Not IsEmpty(mCell2.Value) Then
                            headerRowContent = headerRowContent & " " & CStr(mCell2.Value)
                        Else
                        End If
                        'Insertion des donn�es dans la premi�re cellule de la colonne, sous l'ent�te.
                    Next mCell2
                    '
                    cognosWS.Range(Cells(headerRow + 1, headerFirstColumn - 1), Cells(lastRow, headerFirstColumn - 1)).Value = headerRowContent
                    headerRowContent = ""
                Next mCell
                
                parsingDone = True
                GoTo selectNextHeaderStep

       
        
        Case headerFound And Not headerCopiedToSummary And parsingDone
                'Recopie de l'ent�te dans la feuille sommaire si �a n'a pas encore �t� fait (ne se produit qu'une seulle fois avec la feuille 1)
                CognosMergerForm.currentTaskLabel.Caption = "Recopie de l'ent�te dans la feuille sommaire..."
                    CognosMergerForm.Repaint

                
                'Copie de l'ent�te
                    cognosWS.Range(cognosWS.Cells(headerRow, 1), cognosWS.Cells(headerRow, lastColumn)).Copy
                'Colle l'ent�te dans la feuille sommaire
                    summaryWS.Cells(1, 1).PasteSpecial xlPasteAll
                    headerCopiedToSummary = True
                    GoTo selectNextHeaderStep
                    
                    
       
        Case headerFound And headerCopiedToSummary And parsingDone
            'Commence la proc�dure de recopie des lignes
            
        Case Else
           
        End Select
        
        
'===================================
copyRows:
CognosMergerForm.currentTaskLabel.Caption = "Recopie des lignes dans une feuille sommaire..."
'It�re sur les lignes de donn�es pour recopier les lignes valides
For Each mCell In cognosWS.Range(cognosWS.Cells(headerRow + 1, headerRow), cognosWS.Cells(lastRow, headerRow))
cognosWS.Activate
    
    CognosMergerForm.currentCellLabel.Caption = "Ligne # " & mCell.Row
    CognosMergerForm.Repaint

    'Analyse quel genre de ligne il s'agit
    Select Case True
    
        Case Not (copyTotals) And hasFooter
            Select Case True
                Case mCell.Row = lastRow _
                Or mCell.Interior.Color = 15856114 Or mCell.Interior.Color = 14671839 _
                Or mCell.Offset(0, 1).Interior.Color = 15856114 Or mCell.Offset(0, 1).Interior.Color = 14671839 _
                Or mCell.Offset(0, 2).Interior.Color = 15856114 Or mCell.Offset(0, 2).Interior.Color = 14671839
                'Skip cette ligne, cette ligne est grise, soit la couleur des totaux dans Cognos, ou encore il s'agit de la derni�re ligne,
                'soit la ligne de pieds de page.
                    GoTo nextLine
                Case Else
                    cognosWS.Range(Cells(mCell.Row, 1), Cells(mCell.Row, lastColumn)).Select
                    GoTo copyLine
            End Select
            
            
        Case Not (copyTotals) And Not (hasFooter)
            Select Case True
                Case mCell.Interior.Color = 15856114 Or mCell.Interior.Color = 14671839 _
                Or mCell.Offset(0, 1).Interior.Color = 15856114 Or mCell.Offset(0, 1).Interior.Color = 14671839 _
                Or mCell.Offset(0, 2).Interior.Color = 15856114 Or mCell.Offset(0, 2).Interior.Color = 14671839
                'Skip cette ligne, cette ligne est grise, soit la couleur des totaux dans Cognos
                    GoTo nextLine
                Case Else
                    cognosWS.Range(Cells(mCell.Row, 1), Cells(mCell.Row, lastColumn)).Select
                    GoTo copyLine
            End Select
            
        Case copyTotals And Not (hasFooter)
            'Skip this analysis
            cognosWS.Range(Cells(mCell.Row, 1), Cells(mCell.Row, lastColumn)).Select
            GoTo copyLine
        
        Case copyTotals And hasFooter
            Select Case True
                Case mCell.Row = lastRow
                'Skip cette ligne car il s'agit de la derni�re ligne, soit la ligne de pieds de page.
                    GoTo nextLine
                Case Else
                    cognosWS.Range(Cells(mCell.Row, 1), Cells(mCell.Row, lastColumn)).Select
                    GoTo copyLine
            End Select
        Case Else
            Call MsgBox("Les options n'ont pas �t� d�finies", vbCritical + vbOKOnly, "Options non d�finies")
            End
    End Select
GoTo nextWorksheet


copyLine:
    Selection.Copy
    summaryWS.Cells(summaryWSRowVector, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    summaryWSRowVector = summaryWSRowVector + 1
    
nextLine:
Next mCell

nextWorksheet:

DoEvents
Next cognosWS

CognosMergerForm.Hide
Application.ScreenUpdating = True
Set CognosMerger = summaryWS
End Function

