Attribute VB_Name = "COMTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Format Formula")
Private Sub TestFormatFormula()
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestFormula As String
    TestFormula = "=LAMBDA([IsLowerCase], LET(IsLowerCaseSanitized, IF(ISOMITTED(IsLowerCase), FALSE, IsLowerCase), Chars, CHAR(CODE(""A"") + SEQUENCE(26) - 1), Result, IF(IsLowerCaseSanitized, LOWER(Chars), Chars), Result))(TRUE)"
    
    'Act:
    Dim FormattedFormula As String
    FormattedFormula = FormatFormula(TestFormula)
    
    'Assert:
    Assert.IsFalse Text.IsStartsWith(FormattedFormula, "Failed to parse")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

