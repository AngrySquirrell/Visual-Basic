Sub getOrder()
'
' getOrder Macro
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

    'Insert les deux colonnes
    Range("D1").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    
    'Saisi la 1er formule
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=RC[42]"
    'Etend la 1er formule
    Selection.AutoFill Destination:=Range("D1:D100"), Type:=xlFillDefault

    'Saisi la 2nd formule
    Range("E1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[11]=""Charge"",RC[-1],RC[-1]&""bis"")"
    'Etend la 2nd formule
    Selection.AutoFill Destination:=Range("E1:E100"), Type:=xlFillDefault
    
    'Saisi la 3rd formule
    Range("F1").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-5],'Stripe Virements'!C[-5]:C[1],7,FALSE)"
    'Etend la 3rd formule
    Selection.AutoFill Destination:=Range("F1:F100"), Type:=xlFillDefault
    
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
Sub searchRefund()
'
' searchRefund Macro
'

'

Dim i&, i2&, order_id, order_id2, t
Dim d, d2, listPays, net, sum, checksum

i = 2

    'Check every lines
    Do Until IsEmpty(Cells(i, 1)) = True
    
    Sheets("Stripe Opération").Select
    
        'Attribution de toutes les variables
        d = Cells(i, 8).Value2
        d = Format(d, "mm/dd/yyyy")
        d2 = CDate(Feuil9.Cells(2, 1).Value2)
        order_id = Cells(i, 4).Value
        order_id2 = Cells(i, 5).Value
        listPays = Array("AD", "AL", "AM", "AT", "AX", "AZ", "BA", "BE", "BG", "BY", "CH", "CY", "CZ", "DE", "DK", "EE", "ES", "FI", "FO", "FR", "GB", "GE", "GG", "GI", "GR", "HR", "HU", "IE", "IM", "IS", "IT", "JE", "KZ", "LI", "LT", "LU", "LV", "MC", "MD", "ME", "MK", "MT", "NL", "NO", "PL", "PT", "RO", "RS", "RU", "SE", "SI", "SJ", "SK", "SM", "TR", "UA", "VA")

        MsgBox (ifExist(order_id2) & " | " & order_id2)
                
        'Vérifie que la ligne n'existe pas déjà dans 'Vente 2022'
                'Si e = True alors la ligne N'EXISTE PAS, sinon MsgBox
        If ifExist(order_id2) = False And Format(d, "yyyy") = Format(d2, "yyyy") Then
        
            MsgBox ("THEN initiated")
        
        'Check type to be 'Charge'
            If Cells(i, 1).Value = "Charge" Then
                MsgBox ("Commande inexistante : " & order_id)
                
            End If
            
        'Check type to be 'Refund'
            If Cells(i, 1).Value = "Refund" Then
                MsgBox ("d  " & d)
                MsgBox ("order_id  " & order_id)
                MsgBox ("order_id2  " & order_id2)
            
            'Change de Sheet
                Sheets("Ventes 2022").Select
            
            'Filtre par la date du Refund
                ActiveSheet.Range("$A$1:$AB$12569").AutoFilter Field:=1, Criteria1:=Array( _
                "="), Operator:=xlFilterValues, Criteria2:=Array(2, d)
                Cells(1, 1).Select
            
            'Insert une nouvelle ligne à la fin
                Selection.End(xlDown).Offset(1, 0).Select
                Selection.EntireRow.Insert
            
            'Récupère le numéro de la ligne créée
                i2 = ActiveCell.Row
            
            'Supprime le filtre
                ActiveSheet.Range("$A$1:$AB$12568").AutoFilter Field:=1
            
            'Cherche la case avec le n° de commande
                Cells.Find(What:=order_id, After:=ActiveCell, LookIn:=xlValues, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                , SearchFormat:=False).Activate
            
            'Copie/Colle la ligne trouvée
                ActiveCell.EntireRow.Copy Selection.End(xlDown).Offset(1, 0).EntireRow
                MsgBox (i2)
                
            'Coloration de la ligne fraîchement créée
                Cells(i2, 1).EntireRow.Font.ColorIndex = 45
            
            'Date (A)
                Cells(i2, 1).Value2 = CDate(Format(d, "dd/mm/yyyy"))
            
            'Numéro de commande (B)
                Cells(i2, 2).Value = order_id & "bis"
            
            'Total HT (I)
                Sheets("Stripe").Select
                Var = Cells(i, 13).Value
                Sheets("Ventes 2022").Select
                Cells(i2, 9).Value = Var / 1.2
            
            'Total TVA (J)
                Cells(i2, 10).Value = Var * (1 / 6)
            
            'Total TTC (K)
                Cells(i2, 11).Value = Var
            
            'Expéditions (L->N)
                Cells(i2, 12).Value = 0
                Cells(i2, 13).Value = 0
                Cells(i2, 14).Value = 0
            
            'Total HT (O)
                Cells(i2, 15).Value = Var / 1.2
            
            'Total TTC (P)
                Cells(i2, 16).Value = Var
            
            'Commissions (Q->S)
                Cells(i2, 17).Value = 0
                Cells(i2, 18).Value = 0
                Cells(i2, 19).Value = 0
            
            'Montant versé après commissions (T)
                Cells(i2, 20).Value = Var
                
            'Date du virement(G)
                Var = Feuil14.Cells(i, 2).Value
                Cells(i2, 7).Value = Var
            
            'Montant du virement (H)
                Var = Feuil14.Cells(i, 6).Value
                Cells(i2, , 8).Value = Var
                
                
            End If
            
        End If
        Sheets("Stripe Oprérations").Activate
        
        'Vérifie si le pays d'achat n'appartient pas à l'Europe
        If IsInArray(Feuil14.Cells(i, 30).Value, listPays) = False Then
            'Applique la commission spéciale au pays
            Cells(i2, 17).Value = (Cells(i2, 16).Value * 0.029) + 0.25
        
        End If
        
        'Cherche la ligne avec l'id dans "Ventes 2022" pour récuperer la valeur net
        Sheets("Ventes 2022").Activate
        net = Cells.Find(What:=order_id2, LookIn:=xlValues, LookAt:=xlPart).Offset(0, 19).Value
        Sheets("Stripe Opération").Activate
        
        'Vérifie si l'écart du virement net est supérieur ou égal à .05€
        If Abs(net - Cells(i, 15)) >= 0.05 Then
            'Cherche la cellule net de l'id concerné et la colore en rouge
            Cells.Find(What:=order_id2, LookIn:=xlValues, LookAt:=xlPart).Offset(0, 19).Activate.Font.ColorIndex = 3

        End If
        
        sum = sum + Feuil14.Cells(i, 15)
        
        i = i + 1
    Loop
    
    checksum = Feuil13.Cells(2, 7).Value + Feuil13.Cells(3, 7).Value + Feuil13.Cells(4, 7).Value + Feuil13.Cells(5, 7).Value + Feuil13.Cells(6, 7).Value
    
    If sum <> checksum Then
        MsgBox ("Erreur de checksum, la somme des virement reçue est différent du montant annoncé : " + sum + " =/= " + checksum)
    End If
    
End Sub
