' ============================================================
' GENERADOR DE ÍNDICE + FOLIADOR — CRUBC Los Ríos
' Oficina Técnica
' ============================================================
' Doble clic para abrir la aplicación.
' Este archivo debe estar en la misma carpeta que:
'   generador_indice.py
' ============================================================

Option Explicit

Dim oShell, oFSO, scriptDir, scriptPath
Dim pythonPath, pythonCmd
Dim resultado

Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

' Carpeta donde está este .vbs
scriptDir  = oFSO.GetParentFolderName(WScript.ScriptFullName)
scriptPath = scriptDir & "\generador_indice.py"

' Verificar que el script existe
If Not oFSO.FileExists(scriptPath) Then
    MsgBox "No se encontró generador_indice.py" & vbCrLf & vbCrLf & _
           "Asegúrate de que este archivo esté en la misma carpeta que:" & vbCrLf & _
           "   generador_indice.py", _
           vbCritical, "CRUBC Los Ríos — Error"
    WScript.Quit 1
End If

' Buscar pythonw (sin consola) en rutas comunes
Dim rutas(5)
rutas(0) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
           "\Programs\Python\Python314\pythonw.exe"
rutas(1) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
           "\Programs\Python\Python313\pythonw.exe"
rutas(2) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
           "\Programs\Python\Python312\pythonw.exe"
rutas(3) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
           "\Programs\Python\Python311\pythonw.exe"
rutas(4) = "C:\Python312\pythonw.exe"
rutas(5) = "C:\Python311\pythonw.exe"

pythonPath = ""
Dim i
For i = 0 To 5
    If oFSO.FileExists(rutas(i)) Then
        pythonPath = rutas(i)
        Exit For
    End If
Next

' Si no encontró ruta fija, intentar pythonw del PATH
If pythonPath = "" Then
    On Error Resume Next
    resultado = oShell.Run("pythonw --version", 0, True)
    On Error GoTo 0
    If resultado = 0 Then
        pythonPath = "pythonw"
    End If
End If

' Último intento: python normal del PATH
If pythonPath = "" Then
    On Error Resume Next
    resultado = oShell.Run("python --version", 0, True)
    On Error GoTo 0
    If resultado = 0 Then
        pythonPath = "pythonw"
    End If
End If

' Python no encontrado
If pythonPath = "" Then
    MsgBox "No se encontró Python en este equipo." & vbCrLf & vbCrLf & _
           "Para instalar Python:" & vbCrLf & _
           "1. Ve a: https://www.python.org/downloads/" & vbCrLf & _
           "2. Descarga Python 3.12 o superior" & vbCrLf & _
           "3. Marca 'Add Python to PATH' durante la instalación" & vbCrLf & vbCrLf & _
           "Luego instala las librerías en CMD:" & vbCrLf & _
           "pip install pdfplumber python-docx openpyxl pillow", _
           vbCritical, "CRUBC Los Ríos — Python no encontrado"
    WScript.Quit 1
End If

' Ejecutar sin consola (0 = ventana oculta)
pythonCmd = """" & pythonPath & """ """ & scriptPath & """"
oShell.Run pythonCmd, 0, False

WScript.Quit 0
