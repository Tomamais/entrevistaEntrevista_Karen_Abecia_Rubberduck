Attribute VB_Name = "TestModule1"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")

End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestarFuncaoEhNumero()
    Dim valor As String, resultado As Boolean
    valor = "Karen"
    resultado = ENumero(valor)
    
    Assert.Equals resultado = False
End Sub
