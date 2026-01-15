Attribute VB_Name = "modTranslation"
Option Explicit
Option Private Module

#Const DEVELOPMENT_MODE = True

Private Sub Test()

    '    Const TEST_FORMULA As String = "=LET(SalesData, Sales_Data, Mask, IF(C4 = """", 1, Sales_Data[Category] = C4) * IF(D4 = """", 1, Sales_Data[Product] = D4), Result, IFERROR(FILTER(Sales_Data, Mask), """"), Result)"
    '    Const EXP_FORMULA As String = "=LET(SalesData, Sales_Data, Mask, WENN(C4 = """", 1, Sales_Data[Category] = C4) * WENN(D4 = """", 1, Sales_Data[Product] = D4), Result, WENNFEHLER(FILTER(Sales_Data, Mask), """"), Result)"

    Const TEST_FORMULA As String = "=LET(SalesData, Sales_Data, Mask, IF(C4 = """", 1, Sales_Data[Category] = C4) * IF(D4 = """", 1, Sales_Data[Product] = D4) * (Sales_Data[UnitPrice] >= Price), Result, IFERROR(FILTER(Sales_Data, Mask), """"), Result)"
    Const EXP_FORMULA As String = "=LET(SalesData, Sales_Data, Mask, WENN(C4 = """", 1, Sales_Data[Category] = C4) * WENN(D4 = """", 1, Sales_Data[Product] = D4) * (Sales_Data[UnitPrice] >= Price), Result, WENNFEHLER(FILTER(Sales_Data, Mask), """"), Result)"

    Dim ActualTranslation As String
    ActualTranslation = TranslateEnUSFormulaToApplicationLanguage(TEST_FORMULA)

    If ActualTranslation = EXP_FORMULA Then
        Debug.Print "Correct translation."
    Else
        Debug.Print "Wrong translation."
        Debug.Print "Actual Formula:" & vbNewLine & TEST_FORMULA & vbNewLine & vbNewLine & "Translated formula: " & vbNewLine & ActualTranslation
    End If

End Sub

Private Sub TestTranslatePasteFormula()
    PasteTranslateFormula ActiveCell.Formula2, ActiveCell.Offset(-2, 0)
End Sub

Public Function CopyFormulaToEnglish(ByVal FormulaCell As Range) As String

    If FormulaCell Is Nothing Then Exit Function
    If FormulaCell.Cells.CountLarge > 1 Then Exit Function
    If Not FormulaCell.HasFormula Then Exit Function

    CopyFormulaToEnglish = GetCellFormula(FormulaCell)

End Function

Public Sub PasteTranslateFormula(ByVal enUSFormula As String _
                                 , ByVal AddTranslatedFormulaToCell As Range)

    If enUSFormula = vbNullString Then Exit Sub
    If AddTranslatedFormulaToCell Is Nothing Then Exit Sub

    AddTranslatedFormulaToCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(enUSFormula)

End Sub

Public Function TranslateEnUSFormulaToApplicationLanguage(ByVal Formula As String _
                                                          , Optional ByVal FormulaInSheet As Worksheet) As String

    Dim FormulaInBook As Workbook
    If FormulaInSheet Is Nothing Then
        Set FormulaInBook = ActiveWorkbook
    Else
        Set FormulaInBook = FormulaInSheet.Parent
    End If

    #If DEVELOPMENT_MODE Then
        Dim TranslatorCreator As OARobot.FormulaTranslatorFactory
        Set TranslatorCreator = New OARobot.FormulaTranslatorFactory
        Dim Translator As OARobot.FormulaTranslator
        Dim ParseResult As OARobot.FormulaParseResult
    #Else
        Dim TranslatorCreator As Object
        Set TranslatorCreator = CreateObject("OARobot.FormulaTranslatorFactory")
        Dim Translator As Object
        Dim ParseResult As Object
    #End If

    Set Translator = TranslatorCreator.CreateTranslator(Application)

    Set ParseResult = Translator.TranslateFormula(Formula, _
                                                  DefinedNames:=GetNames(FormulaInBook))
    If ParseResult.ParseSuccess Then
        TranslateEnUSFormulaToApplicationLanguage = ParseResult.Formula
    Else
        Err.Raise 13, "Formula parsing failed.", "Check if you have provided a valid formula or not."
    End If

End Function

Public Function TranslateApplicationLanguageFormulaToEnUS(ByVal Formula As String _
                                                          , Optional ByVal FormulaInSheet As Worksheet) As String

    Dim FormulaInBook As Workbook
    If FormulaInSheet Is Nothing Then
        Set FormulaInBook = ActiveWorkbook
    Else
        Set FormulaInBook = FormulaInSheet.Parent
    End If

    #If DEVELOPMENT_MODE Then
        Dim TranslatorCreator As OARobot.FormulaTranslatorFactory
        Set TranslatorCreator = New OARobot.FormulaTranslatorFactory
        Dim Translator As OARobot.FormulaTranslator
        Dim ParseResult As OARobot.FormulaParseResult
        Dim LocaleCreator As OARobot.FormulaLocaleInfoFactory
        Set LocaleCreator = New OARobot.FormulaLocaleInfoFactory
        Dim ApplicationLocale As OARobot.FormulaLocaleInfo
    #Else
        Dim TranslatorCreator As Object
        Set TranslatorCreator = CreateObject("OARobot.FormulaTranslatorFactory")
        Dim Translator As Object
        Dim ParseResult As Object
        Dim LocaleCreator As Object
        Set LocaleCreator = CreateObject("OARobot.FormulaLocaleInfoFactory")
        Dim ApplicationLocale As Object
    #End If

    Set Translator = TranslatorCreator.CreateTranslator(Application)

    Set ApplicationLocale = LocaleCreator.CreateFromExcel(Application)

    Set ParseResult = Translator.TranslateFormula(Formula, _
                                                  formulaLocale:=ApplicationLocale, _
                                                  translateTo:=LocaleCreator.EN_US, _
                                                  DefinedNames:=GetNames(FormulaInBook))
    If ParseResult.ParseSuccess Then
        TranslateApplicationLanguageFormulaToEnUS = ParseResult.Formula
    Else
        Err.Raise 13, "Formula parsing failed.", "Check if you have provided a valid formula or not."
    End If

End Function

Public Sub TranslateUsingExcelSettings()

    Dim txf As New OARobot.FormulaTranslatorFactory
    Dim tx As OARobot.FormulaTranslator

    Set tx = txf.CreateTranslator(Application)

    Dim s As String
    Dim X As OARobot.FormulaParseResult
    s = Sheet1.Cells(1, 1).Value
    Set X = tx.TranslateFormula(s)               ' Translates to Application settings
    Debug.Print X.ParseSuccess & ":" & X.Formula

    'Translates as specified

    Dim Locale As New OARobot.FormulaLocaleInfoFactory

    Dim fifi As OARobot.FormulaLocaleInfo

    Dim ApplicationLocale As OARobot.FormulaLocaleInfo
    Set ApplicationLocale = Locale.CreateFromExcel(Application)

    Set fifi = Locale.CreateFromLocaleName("fi-fi", "{", "}", "[", "]", "R", "C", "r", "c", ";", ";", "@", ";", ",")
    Set X = tx.TranslateFormula(s, formulaLocale:=Locale.EN_US, translateTo:=fifi)
    Debug.Print vbNewLine & X.ParseSuccess & ":" & X.Formula

End Sub

#If DEVELOPMENT_MODE Then
Public Function GetNames(Optional ByVal ForBook As Workbook) As OARobot.XLDefinedNames
#Else
Public Function GetNames(Optional ByVal ForBook As Workbook) As Object
#End If

If ForBook Is Nothing Then Set ForBook = ActiveWorkbook

#If DEVELOPMENT_MODE Then
    Dim NamesFactory As New OARobot.DefinedNamesFactory
    Set GetNames = NamesFactory.Create(ForBook)
#Else
    Set GetNames = CreateObject("OARobot.DefinedNamesFactory").Create(ForBook)
#End If

End Function


