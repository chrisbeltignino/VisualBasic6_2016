Attribute VB_Name = "Sentencias"
'Abrir Bd e indicar que somos el usuario cliente
Sub main()
With Base
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\SistemaUsuarios.mdb;Persist Security Info=False"
     Login.Show
End With
End Sub

'si la tabla esta abierta entonces que se cierre
'abrirla de nuevo y seleccionar la tabla USUARIOS
Sub Usuarioss()
With Usuarios
    If .State = 1 Then .Close
    .Open "select * from Usuarios", Base, adOpenStatic, adLockOptimistic
End With
End Sub

