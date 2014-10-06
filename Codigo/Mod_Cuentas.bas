Attribute VB_Name = "Mod_Cuentas"
Option Explicit
Public Type pjs
    NamePJ As String
    LvlPJ As Byte
    ClasePJ As eClass
End Type
Public Type Acc
    name As String
    Pass As String
    
    CantPjs As Byte
    PJ(1 To 8) As pjs
End Type
Public Cuenta As Acc

Public Sub CrearCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String, ByVal email As String)
Dim ciclo As Byte
'¿Posee caracteres invalidos?
If Not AsciiValidos(name) Or LenB(name) = 0 Then
    Call WriteErrorMsg(UserIndex, "Nombre invalido.")
    Exit Sub
End If

'Si ya existe la cuenta
If FileExist(App.Path & "\Cuentas\" & name & ".bgao", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre de la cuenta ya existe, por favor ingresa otro.")
    Exit Sub
End If

Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "NOMBRE", name)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "PASSWORD", Pass)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "MAIL", email)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "FECHA_CREACION", Now)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "BAN", "0")

'************************RELLENO LOS PJs************************'
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "CANTIDAD_PJS", "0")
For ciclo = 1 To 8
    Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ" & ciclo, "")
Next ciclo
'************************************************************'

Call EnviarCuenta(UserIndex, "", "", "", "", "", "", "", "", "0", "1")
End Sub

Public Sub ConectarCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String)
'Si NO existe la cuenta
If Not FileExist(App.Path & "\Cuentas\" & name & ".bgao", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre de la cuenta es inexistente.")
    Exit Sub
End If

With Cuenta
'Si la contraseña es correcta
If Pass = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "PASSWORD") Then
    If GetVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "BAN") <> "0" Then
        Call WriteErrorMsg(UserIndex, "Se ha denegado el acceso a tu cuenta por mal comportamiento en el servidor. Por favor comunicate con los administradores del juego para más información.")
    Else
        .CantPjs = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
        
        .PJ(1).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ1")
        .PJ(2).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ2")
        .PJ(3).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ3")
        .PJ(4).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ4")
        .PJ(5).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ5")
        .PJ(6).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ6")
        .PJ(7).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ7")
        .PJ(8).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ8")
        
        Call EnviarCuenta(UserIndex, .PJ(1).NamePJ, .PJ(2).NamePJ, .PJ(3).NamePJ, .PJ(4).NamePJ, _
        .PJ(5).NamePJ, .PJ(6).NamePJ, .PJ(7).NamePJ, .PJ(8).NamePJ, .CantPjs, "1")
        
        Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
    End If
Else
    Call WriteErrorMsg(UserIndex, "La contraseña es incorrecta. Por favor intentalo nuevamente.")
    Exit Sub
End If
End With
End Sub

Public Sub AgregarPersonaje(ByVal UserIndex As Integer, ByVal UserName As String, ByVal UserIndexFile As String)
Dim CantidadPJs As Byte

If FileExist(App.Path & "\Charfile\" & UserName & ".CHR", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre del chango ya existe, por favor ingresa otro.")
    Exit Sub
End If

If Not FileExist(App.Path & "\Charfile\" & UserName & ".CHR", vbNormal) Then
    CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
    
    WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs + 1
    WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & (CantidadPJs + 1), UserName
    
End If

Call SaveUser(UserIndex, UserIndexFile)

End Sub

Public Sub ActualizarCuenta(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal UserName As String)
With Cuenta
'Actualizamos la cuenta.
        .CantPjs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
        
        .PJ(1).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ1")
        .PJ(2).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ2")
        .PJ(3).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ3")
        .PJ(4).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ4")
        .PJ(5).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ5")
        .PJ(6).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ6")
        .PJ(7).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ7")
        .PJ(8).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ8")
        
        Call EnviarCuenta(UserIndex, .PJ(1).NamePJ, .PJ(2).NamePJ, .PJ(3).NamePJ, .PJ(4).NamePJ, _
        .PJ(5).NamePJ, .PJ(6).NamePJ, .PJ(7).NamePJ, .PJ(8).NamePJ, .CantPjs, "1")
        
        Call WriteVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)

End With
End Sub

Public Sub BorrarPersonaje(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal IndiceUser As String)
Dim CantidadPJs As Byte
Dim NamePJ As String
Dim c As String
Dim d As String
Dim f As String
Dim g As String
Dim H As Byte
Dim i As String
Dim j As String
    
'Consulto el nombre del PJ a eliminar
NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & IndiceUser)

CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")

WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs - 1
WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & IndiceUser, ""

        Call WriteErrorMsg(UserIndex, "Personaje eliminado con éxito.")
    Call Kill(App.Path & "\Charfile\" & NamePJ & ".CHR")

With Cuenta
'Actualizamos la cuenta.
        .CantPjs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
        
        .PJ(1).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ1")
        .PJ(2).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ2")
        .PJ(3).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ3")
        .PJ(4).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ4")
        .PJ(5).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ5")
        .PJ(6).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ6")
        .PJ(7).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ7")
        .PJ(8).NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ8")
        
        Call EnviarCuenta(UserIndex, .PJ(1).NamePJ, .PJ(2).NamePJ, .PJ(3).NamePJ, .PJ(4).NamePJ, _
        .PJ(5).NamePJ, .PJ(6).NamePJ, .PJ(7).NamePJ, .PJ(8).NamePJ, .CantPjs, "1")
        
        Call WriteVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)

End With


 
End Sub

