Attribute VB_Name = "Módulo1"
Sub lineNum()

Dim numeroOrden, lineNum, cont, ox As Integer
numeroOrden = 9
lineNum = 2
ox = 3
Do While Cells(ox, numeroOrden).Value <> ""
  
cont = cont + 1
Cells(ox, lineNum).Value = cont


If (Cells(ox, numeroOrden).Value <> Cells(ox + 1, numeroOrden).Value) Then
cont = 0
End If

  ox = ox + 1

Loop

End Sub
Sub Amazon()
Dim hoja As String
hoja = "Amazon"

   Sheets("BASE").Select
   ActiveSheet.Range("$A$1:$BU$417").AutoFilter Field:=57, Criteria1:="<>"
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("E3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("U2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("U2").Select
    
    Sheets(hoja).Select
    Range("C3").Select
    ActiveSheet.Paste
    
    Sheets("BASE").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("I3").Select
    ActiveSheet.Paste
    
    Sheets("BASE").Select
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("F3").Select
    ActiveSheet.Paste
    
       Sheets("BASE").Select
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("F3").Select
    ActiveSheet.Paste
    iva
      Range("F:F").Select
    Selection.NumberFormat = "0.00"
    
    lineNum
    Sheets(hoja).Select
        Range("E1").Select
    Selection.AutoFilter
    Sheets("BASE").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

    
End Sub
Sub Meli()


Dim hoja As String
hoja = "Meli"
   
   
   Sheets("BASE").Select

    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BU$417").AutoFilter Field:=48, Criteria1:= _
        "Mercadolibre"


    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("E3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("U2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("U2").Select
    
    Sheets(hoja).Select
    
    Range("C3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select

    Selection.Copy
    Sheets(hoja).Select
    Range("I3").Select
    ActiveSheet.Paste
       Sheets("BASE").Select
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("F3").Select
    ActiveSheet.Paste
    iva
      Range("F:F").Select
    Selection.NumberFormat = "0.00"
    
    lineNum
    Sheets(hoja).Select
        Range("E1").Select
    Selection.AutoFilter
    Sheets("BASE").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

End Sub
Sub Tutti()

Dim hoja As String
hoja = "Tutti"
   
Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BU$417").AutoFilter Field:=1, Criteria1:="=#*", _
        Operator:=xlAnd
        
        

    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("E3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("U2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("U2").Select
    
    Sheets(hoja).Select
    
    Range("C3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select

    Selection.Copy
    Sheets(hoja).Select
    Range("I3").Select
    ActiveSheet.Paste
    
    Sheets("BASE").Select
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("F3").Select
    ActiveSheet.Paste
    
    iva
     Range("F:F").Select
    Selection.NumberFormat = "0.00"
   
    lineNum
    Sheets(hoja).Select
        Range("E1").Select
    Selection.AutoFilter
    Sheets("BASE").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        
        

End Sub
Sub Linio()

Dim hoja As String
hoja = "Linio"
   
   
   Sheets("BASE").Select

    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BU$417").AutoFilter Field:=48, Criteria1:= _
        "Linio"


    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("E3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("U2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("U2").Select
    
    Sheets(hoja).Select
    
    Range("C3").Select
    ActiveSheet.Paste
    Sheets("BASE").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select

    Selection.Copy
    Sheets(hoja).Select
    Range("I3").Select
    ActiveSheet.Paste
       Sheets("BASE").Select
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets(hoja).Select
    Range("F3").Select
    ActiveSheet.Paste
    iva
    Range("F:F").Select
    Selection.NumberFormat = "0.00"
    
    lineNum
    Sheets(hoja).Select
        Range("E1").Select
    Selection.AutoFilter
    Sheets("BASE").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData


End Sub
Sub Ejecutar()
Amazon
Meli
Tutti
Linio
End Sub
