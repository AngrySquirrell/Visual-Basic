Sub Create_Stripe()
'
' Create_Stripe Macro
'

'
    
    'Transforme le format CSV en XLS
    Sheets("Stripe Virements").Select
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
        Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
        33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), _
        Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array( _
        46, 1), Array(47, 1), Array(48, 1), Array(49, 1)), TrailingMinusNumbers:=True
    Sheets("Stripe Opérations").Select
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
        Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
        33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), _
        Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array( _
        46, 1), Array(47, 1), Array(48, 1), Array(49, 1)), TrailingMinusNumbers:=True

    'Change les . par des ,
    Columns("A:AZ").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _
        ReplaceFormat:=True

    'Insert les trois colonnes
    Range("D1").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    
    'Saisi la 1er formule
    Range("D1").Select
    ActiveCell.EntireRow.Interior.ColorIndex = 36
    ActiveCell.FormulaR1C1 = "=RC[42]"
    'Etend la 1er formule
    Selection.AutoFill Destination:=Range("D1:D1000"), Type:=xlFillDefault

    'Saisi la 2nd formule
    Range("E1").Select
    ActiveCell.EntireRow.Interior.ColorIndex = 36
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[11]=""Charge"",RC[-1],RC[-1]&""bis"")"
    'Etend la 2nd formule
    Selection.AutoFill Destination:=Range("E1:E1000"), Type:=xlFillDefault
    
    'Saisi la 3rd formule
    Range("F1").Select
    ActiveCell.EntireRow.Interior.ColorIndex = 36
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-5],'Stripe Virements'!C[-5]:C[1],7,FALSE)"
    'Etend la 3rd formule
    Selection.AutoFill Destination:=Range("F1:F1000"), Type:=xlFillDefault
    
    Cells(1, 1).Select
End Sub
Function ifExist(Search) As Boolean
    Dim e As Range
    Set e = Feuil9.Cells.Find(What:=Search, LookIn:=xlValues, LookAt:=xlPart)
    ifExist = (Not e Is Nothing)
End Function
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
Sub Main_Stripe()
'
' Main_Stripe Macro
'

'

Dim i&, i2&, i°&, i§, order_id&, order_id2$, t, var, main
Dim d, d2, listCountry, net, sum, checksum

listCountry = Array( _
        "AD", "AL", "AM", "AT", "AX", _
        "AZ", "BA", "BE", "BG", "BY", _
        "CH", "CY", "CZ", "DE", "DK", _
        "EE", "ES", "FI", "FO", "FR", _
        "GB", "GE", "GG", "GI", "GR", _
        "HR", "HU", "IE", "IM", "IS", _
        "IT", "JE", "KZ", "LI", "LT", _
        "LU", "LV", "MC", "MD", "ME", _
        "MK", "MT", "NL", "NO", "PL", _
        "PT", "RO", "RS", "RU", "SE", _
        "SI", "SJ", "SK", "SM", "TR", _
        "UA", "VA" _
        )

i = 2
main = "Ventes 2022"
sum = 0


    'Check every lines
    Do Until IsEmpty(Cells(i, 1)) = True
    
        Sheets("Stripe Opérations").Select
    
        'Attribution de toutes les variables
        
        d = Cells(i, 8).Value2
        d = Format(d, "mm/dd/yyyy")
        d2 = CDate(Feuil9.Cells(2, 1).Value2)
        d2 = Format(d2, "mm/dd/yyyy")
        
        order_id = Cells(i, 4).Value
        order_id2 = Cells(i, 5).Value
        
        'MsgBox (ifExist(order_id2) & " | " & order_id2)
        'MsgBox (Format(d, "yyyy") & Format(d2, "yyyy"))
                
        'Vérifie que la ligne n'existe pas déjà dans 'Vente 2022'
                'Si = False alors la ligne N'EXISTE PAS, sinon MsgBox
        If ifExist(order_id2) = False And Format(d, "yyyy") = Format(d2, "yyyy") And ifExist(order_id) = True Then
        
            'MsgBox ("THEN initiated")
        
        'Check type to be 'Charge' - " & Cells(i, 16).Value)
            If Cells(i, 16).Value = "charge" Then
                MsgBox ("Case 'Charge' selected")
                MsgBox ("Commande inexistante : " & order_id)
                
            End If
            
        'Check type to be 'Refund' - " & Cells(i, 16).Value)
            If Cells(i, 16).Value = "refund" Then
                MsgBox ("Case 'Refund' selected")
                MsgBox ("d  " & d)
                MsgBox ("order_id  " & order_id)
                MsgBox ("order_id2  " & order_id2)
            
            'Change de Sheet")
                Sheets(main).Select
            
                'Filtre par la date du Refund")
                    ActiveSheet.Range("$A$1:$AB$12569").AutoFilter Field:=1, Criteria1:=Array( _
                    "="), Operator:=xlFilterValues, Criteria2:=Array(2, d)
                    Cells(1, 1).Select
            
                'Insert une nouvelle ligne à la fin")
                    Selection.End(xlDown).Offset(1, 0).Select
                    Selection.EntireRow.Insert
            
                'Récupère le numéro de la ligne créée")
                    i2 = ActiveCell.Row
            
                'Supprime le filtre")
                    ActiveSheet.Range("$A$1:$AB$12568").AutoFilter Field:=1
            
                'Cherche la case avec le n° de commande")
                    Cells.Find(What:=order_id, LookIn:=xlValues, LookAt:=xlPart).Activate
            
            
                'Copie/Colle la ligne trouvée dans celle fraîchement créée")
                    ActiveCell.EntireRow.Copy
                    Selection.End(xlDown).Offset(1, 0).EntireRow.PasteSpecial xlPasteValues
                
                'Coloration de la ligne fraîchement créée pour ajouter un peu de joie dans nos coeurs")
                    Cells(i2, 1).EntireRow.Interior.ColorIndex = 17
            
                'Date (A)")
                    Cells(i2, 1).Value2 = CDate(Format(d, "dd/mm/yyyy"))
            
                'Numéro de commande (B)
                    Cells(i2, 2).Value = order_id & "bis"
            
                'Total HT (I)
                    Sheets("Stripe").Select
                    var = Cells(i, 13).Value
                    Sheets(main).Select
                    Cells(i2, 9).Value = var / 1.2                                      ' I = TTC(I) / 1.2
            
                'Total TVA (J)
                    Cells(i2, 10).Value = var * (1 / 6)                                 ' J = I * 1/6
            
                'Total TTC (K)
                    Cells(i2, 11).Value = var                                           ' K = TTC(I)
            
                'Expéditions (L->N)
                    Cells(i2, 12).Value = 0                                             ' L = 0
                    Cells(i2, 13).Value = 0                                             ' M = 0
                    Cells(i2, 14).Value = 0                                             ' N = 0
            
                'Total HT (O)
                    Cells(i2, 15).Value = var / 1.2                                     ' O = I
            
                'Total TTC (P)
                    Cells(i2, 16).Value = var                                           ' P = TTC(I)
            
                'Commissions (Q->S)
                    Cells(i2, 17).Value = 0                                             ' Q = 0
                    Cells(i2, 18).Value = 0                                             ' R = 0
                    Cells(i2, 19).Value = 0                                             ' S = 0
            
                'Montant versé après commissions (T)
                    Cells(i2, 20).Value = var                                           ' T = TTC(I)
                
                'Date du virement(G)
                    var = Feuil14.Cells(i, 2).Value
                    Cells(i2, 7).Value = var                                            ' G = Date(B)
            
                'Montant du virement (H)
                    var = Feuil14.Cells(i, 6).Value
                    Cells(i2, 8).Value = var                                            ' H = F
                
            End If
            
            MsgBox ("Commande non trouvée sur 'Ventes 2022' :  " & order_id & " | Possible si en début d'année, si achat de l'an dernier est remboursé cette année)")
            
            Sheets("Stripe Opérations").Select
        
            'Vérifie si le pays d'achat n'appartient pas à l'Europe")
            If IsInArray(Feuil14.Cells(i, 30).Value, listCountry) = False Then
                MsgBox ("Payement hors Europe identifié " & i)
                'Applique la commission spéciale au pays
                Cells(i2, 17).Value = (Cells(i2, 16).Value * 0.029) + 0.25
        
            End If
        
            'Cherche la ligne avec l'id dans 'Ventes 2022' pour récuperer la valeur net")
            Sheets(main).Select
        
            net = Cells.Find(What:=order_id2, LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 18).Value
        
            Sheets("Stripe Opérations").Select
        
            'Vérifie si l'écart du virement net est supérieur ou égal à .05€")
            If Abs(net - Feuil14.Cells(i, 15)) >= 0.05 Then
                'Cherche la cellule net de l'id concerné et la colore en rouge
                Cells(i, 1).EntireRow.Interior.ColorIndex = 46

            End If
            
        End If
        
        'Fait la somme des opérations
        'sum = sum + Val(Feuil14.Cells(i, 15).Value)
        'MsgBox (sum)
        
        i = i + 1
    Loop
    
    
    
    
    Sheets("Stripe Opérations").Select
    i° = 2
    i§ = 2
    Do Until IsEmpty(Cells(i°, 1)) = True
    
    checksum = Val(Cells(i°, 6).Value)
    MsgBox (Cells(i°, 6).Value & " - " & Cells(i°, 6).Offset(1, 0).Value)
    
        If Cells(i°, 6).Value = Cells(i°, 6).Offset(1, 0).Value Then
        
            sum = sum + Val(Feuil14.Cells(i°, 15))
            MsgBox (checksum & "    -    " & sum)
            
        Else
            MsgBox (Abs(sum - checksum))
            If Abs(sum - checksum) >= 5 Then
                MsgBox ("Erreur de checksum, la somme des virement reçue est différent du montant annoncé : " & sum & " <> " & checksum)
            Else
                MsgBox ("La sum et checksum semblent être égaux : " & sum & " <> " & checksum)
            End If
            
            sum = 0
        End If
    
    i° = i° + 1
    
    Loop
    
End Sub
Sub Create_Paypal()
'
' Create_Paypal Macro
'

'

    Sheets("Paypal").Select
    
    'Regarde su la cellule B1 est vide
    If IsEmpty(Cells(1, 2)) = True Then
    'Converti le CVS en XLS
        Columns("A:A").Select
        Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
            Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
            ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
            (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
            Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
            33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), _
            Array(40, 1), Array(41, 1)), TrailingMinusNumbers:=True
    End If
    
    
    'Change les . par des ,
    Columns("A:AZ").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _
        ReplaceFormat:=True

    'Insert les deux colonnes
    Range("I1").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    
    'Saisi la 1er formule
    Range("I1").Select
    ActiveCell.EntireRow.Interior.ColorIndex = 36
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1],0,IF(OR(RC[23]=""FR"",RC[23]=""BE""),ROUND(0.029*RC[-1],2)+0.35,ROUND(RC[-1]*0.0489,2)+0.35)+RC[1])"
    'Etend la 1er formule
    Selection.AutoFill Destination:=Range("I1:I1000"), Type:=xlFillDefault
    
    'Saisi la 2nd formule
    Range("J1").Select
    ActiveCell.EntireRow.Interior.ColorIndex = 36
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[12],3)=""PRM"",RIGHT(RC[12],5),""NULL"")"
    Selection.AutoFill Destination:=Range("J1:J13"), Type:=xlFillDefault
    'Etend la 2nd formule
    Selection.AutoFill Destination:=Range("J1:J1000"), Type:=xlFillDefault
    
    'Saisi la 3rd formule
    Range("K1").Select
    ActiveCell.EntireRow.Interior.ColorIndex = 36
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""Remboursement de paiement"",RC[-1]&""bis"",RC[-1])"
    Selection.AutoFill Destination:=Range("K1:K13"), Type:=xlFillDefault
    'Etend la 3rd formule
    Selection.AutoFill Destination:=Range("K1:K1000"), Type:=xlFillDefault
    
    Cells(1, 1).Select
    
End Sub
Sub Main_Paypal()
'
' Main_Paypal Macro
'

'
Dim i&, i2&, checksum&, sum&, order_id, order_id2, d, d2, var, main$, import$
Dim i°&, i2°&

i = 2
i° = 2
checksum = 0
main = "Ventes 2022"
import = "Paypal"

'début de la 1er boucle
Do Until IsEmpty(Cells(i, 1)) = True

    Sheets(import).Select

    'Attribution des valeurs
    d = Cells(i, 1).Value2
    d = Format(d, "mm/dd/yyyy")
    d2 = CDate(Feuil9.Cells(2, 1).Value2)
    d2 = Format(d2, "mm/dd/yyyy")
    order_id = Cells(i, 10).Value
    order_id2 = Cells(i, 11).Value

    
    'Check if order_id2 exist in 'Ventes 2022' and year match between thoses two sheets
    If ifExist(order_id2) = False And Format(d, "yyyy") = Format(d2, "yyyy") Then
    
        'Inutile pour l'instant
        If Cells(i, 5).Value = "Paiement sur site" Then
        
        End If
        
        'Check type to be 'Refund'
        If Cells(i, 5).Value = "Remboursement de paiement" Then
            
            If ifExist(order_id) = True Then
            
            'Change de Sheet")
                Sheets(main).Select
            
            'Filtre par la date du Refund")
                ActiveSheet.Range("$A$1:$AB$12569").AutoFilter Field:=1, Criteria1:=Array( _
                "="), Operator:=xlFilterValues, Criteria2:=Array(2, d)
                Cells(1, 1).Select
            
            'Insert une nouvelle ligne à la fin")
                Selection.End(xlDown).Offset(1, 0).Select
                Selection.EntireRow.Insert
            
            'Récupère le numéro de la ligne créée")
                i2 = ActiveCell.Row
            
            'Supprime le filtre")
                ActiveSheet.Range("$A$1:$AB$12568").AutoFilter Field:=1
            
            'Cherche la case avec le n° de commande")
                Cells.Find(What:=order_id, LookIn:=xlValues, LookAt:=xlPart).Activate
                
            'Coloration de la ligne fraîchement créée pour ajouter un peu de joie dans ce monde de brutes")
                Cells(i2, 1).EntireRow.Interior.ColorIndex = 38
                
            'Copie/Colle la ligne trouvée dans celle fraîchement créée")
                ActiveCell.EntireRow.Copy
                Selection.End(xlDown).Offset(1, 0).EntireRow.PasteSpecial xlPasteValues
                
            'Date (A)")
                Cells(i2, 1).Value2 = CDate(Format(d, "dd/mm/yyyy"))
            
            'Numéro de commande (B)
                Cells(i2, 2).Value = order_id & "bis"
            
            'Total HT (I)
                Sheets(import).Select
                var = Cells(i, 13).Value
                Sheets(main).Select
                Cells(i2, 9).Value = var / 1.2                                      ' I = Net(M) / 1.2
            
            'Total TVA (J)
                Cells(i2, 10).Value = var * 0.2                                     ' J = I * 0.2
            
            'Total TTC (K)
                Cells(i2, 11).Value = Cells(i2, 9).Value + Cells(i2, 10).Value      ' K = I + J
            
            'Expéditions (L->N)
                Cells(i2, 12).Value = 0                                             ' L = 0
                Cells(i2, 13).Value = 0                                             ' M = 0
                Cells(i2, 14).Value = 0                                             ' N = 0
            
            'Total HT (O)
                Cells(i2, 15).Value = Cells(i2, 9).Value                            ' O = I
            
            'Total TTC (P)
                Cells(i2, 16).Value = Cells(i2, 11).Value                           ' P = K
            
            'Commissions (Q->S)
                Cells(i2, 17).Value = 0                                             ' Q = 0
                Cells(i2, 18).Value = 0                                             ' R = 0
                Cells(i2, 19).Value = 0                                             ' S = 0
            
            'Montant versé après commissions (T)
                Cells(i2, 20).Value = Cells(i2, 11).Value                           ' T = K
                
            'Date du virement(G)
                var = Feuil14.Cells(i, 1).Value
                Cells(i2, 7).Value = CDate(Format(var, "dd/mm/yyyy"))               ' G = Date(A)
                
            Else
            
                'Message d'erreur si le remboursement existe à l'an x mais pas le paiement
                MsgBox ("Commande passé l'an dernier, alors que le remboursement a été effectué cette année. ID conserné : " & order_id)
            
            End If
    
        End If

    End If
    
    i = i + 1
Loop

'Deuxième boucle    -   Vérifie que la somme des payement soit bien égal aux virements reçus
Do Until IsEmpty(Cells(i°, 1)) = True

    Sheets(import).Select
    
    'Check type differ from 'Virement Standard'
    If Cells(i°, 5).Value <> "Virement standard" Then
    
        'Fait la somme des virements
        sum = sum + Cells(i°, 13).Value
    
    Else
    
        'Attribue la checksum
        checksum = Cells(i°, 13).Value
        'Vérifie si la différence entre sum et checksum est inférieure à 1€
        If Abs(checksum - sum) >= 1 Then
        
            MsgBox ("Erreur | Le virement reçue semble différent de la somme des montants perçus (Normal si en début d'année) - Virement concerné : l." & i° & " Delta : " & Abs(sum - checksum))
            
        End If
    
    End If
    
    i° = i° + 1
Loop

End Sub
