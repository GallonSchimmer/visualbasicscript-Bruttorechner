' Eingabe
netto = InputBox("Geben Sie den Netto-Betrag ein!","Eingabe Netto")
' Verarbeitung (Berechnung)
mwst = 0.19 * netto
brutto = netto + mwst
'Ausgabe
ergebnis = "Netto:" & vbTab & vbTab & FormatCurrency(netto) & _
           vbNewline & _
           "+ 19% MwSt:" & vbTab & FormatCurrency(mwst) & _
           vbNewline & _
           "Gesamt: " & vbTab & vbTab & FormatCurrency(brutto)


MsgBox ergebnis,,"Ergebnis"