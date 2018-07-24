Sub RUN_DAILY_COST_EXTRACTION()
 
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    
    'countries = Array("Austria", "Europe", "France", "Germany", "Italy", "Spain", "Switzerland", "UK", "USA")
    countries = Array("Germany", "Austria", "Europe", "France", "Italy", "Spain") ', "UK", "Netherlands", "Belgium", "USA", "Switzerland")
    'countries = Array("Germany", "Austria", "Europe", "France", "Italy", "Spain", "UK", "Netherlands", "Belgium", "USA", "Switzerland")
         
 
    'Clear sheet Result
    Windows("DailyCostExtraction.xlsm").Activate
    Sheets("Result").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp

    
    'Iterate over all CC-Sheets
    For Each country In countries
    
        'Clear sheet Country
        Windows("DailyCostExtraction.xlsm").Activate
        Sheets("Country").Select
        Cells.Select
        Application.CutCopyMode = False
        
        Selection.Delete Shift:=xlUp

        
        'Open next CC Sheet
        Filename = "Channel Controlling 2018 " & country & ".xlsx"
        Application.DisplayAlerts = False
        Workbooks.Open "Z:\800-Management\830-Controlling\833-Marketing\Channel Controlling 2018\" & Filename, UpdateLinks:=3
        Application.DisplayAlerts = True
        Application.Calculation = xlCalculationManual
        'tabs = Workbooks("Z:\800-Management\830-Controlling\833-Marketing\Channel Controlling 2016\" & Filename).Worksheets.Count
        
        
        'Copy Date once
        Windows(Filename).Activate
        tabs = ActiveWorkbook.Worksheets.Count
        Sheets(2).Select
        Range("A3:A407").Select
        Selection.Copy
        Windows("DailyCostExtraction.xlsm").Activate
        Sheets("Campaign").Select
        Range("B2").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
         :=False, Transpose:=False
     
      
     
        'Iterate over all Sheets except DataPivot within a CC-Sheet
        For n = 3 To tabs - 1

              
             ' Copy Costs to campaign
             Windows(Filename).Activate
             Sheets(n).Select
             Range("G3:G407").Select
             Selection.Copy
             Windows("DailyCostExtraction.xlsm").Activate
             Sheets("Campaign").Select
             Range("C2").Select
             Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
                 :=False, Transpose:=False
            
             ' Copy Affiliate to campaign
             Windows(Filename).Activate
             Sheets(n).Select
             Range("A502").Select
             Selection.Copy
             Windows("DailyCostExtraction.xlsm").Activate
             Sheets("Campaign").Select
             Range("A2:A406").Select
             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                 :=False, Transpose:=False
             
             ' Copy from campaign to country
             Sheets("Campaign").Select
             Range("A2:C2", Selection.End(xlDown)).Select
             Selection.Copy
             Sheets("Country").Select
             ActiveCell.Offset(1, 0).Range("A1").Select
             ActiveSheet.Paste
             Selection.End(xlDown).Select
                
        Next n
        
    
        'Copy Country label
        Windows(Filename).Activate
        Sheets(1).Select
        Range("A2").Copy
        Windows("DailyCostExtraction.xlsm").Activate
        Sheets("Country").Select
        ActiveCell.Offset(0, 3).Range("A1").Select
        Range(ActiveCell, Selection.End(xlUp)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("D1").Value = "Country"
        
        
        'Copy Country to Result
        
        
        Range("A2:D2").Select
        Range(Selection, Selection.End(xlDown)).Select
        'Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Result").Select
        ActiveCell.Offset(1, 0).Range("A1").Select
         ActiveSheet.Paste
        Selection.End(xlDown).Select
     
        
        Windows(Filename).Activate
        
        'ActiveWorkbook.Save 'dangerous!
        Application.Calculation = xlCalculationAutomatic
        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close savechanges:=False
        
    Next country
    
    'Add Labels
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "AffiliateGroup"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "MarketingCost"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Country"
    Range("A1").Select
    
    
    'Adds filter, deletion part done by Vaidehi :-)
       Columns("A:D").Select
    Range("D1").Activate
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$D$100000").AutoFilter Field:=3, Criteria1:="=0,0 €" _
        , Operator:=xlOr, Criteria2:="="

   
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
   
    MsgBox "Please, delete Null Marketing Cost rows and then DONE."
  
End Sub
