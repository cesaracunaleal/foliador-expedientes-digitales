' Gestor de Expedientes Digitales v7.1
' Lanzador sin consola
Option Explicit
Dim oShell, oFSO, scriptDir, scriptPath, configPath, batPath
Dim pythonPath, pythonwPath, pythonExe, pythonCmd, resultado, i
Dim localApp, rutas(13)
Dim tmpVer, tmpScript, tmpResult
Dim fsW, fsR, checkOut, verCmd
Dim m1, m2, m3, faltantes, respLib, respCfg
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
scriptDir  = oFSO.GetParentFolderName(WScript.ScriptFullName)
scriptPath = scriptDir & "\generador_indice.py"
configPath = scriptDir & "\config_indice.json"
batPath    = scriptDir & "\instalar_librerias.bat"

' Paso 1: verificar script
If Not oFSO.FileExists(scriptPath) Then
    MsgBox "No se encontro generador_indice.py en esta carpeta.", vbCritical, "Error"
    WScript.Quit 1
End If

' Paso 2: buscar Python
localApp = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
rutas(0)  = localApp & "\Programs\Python\Python314\pythonw.exe"
rutas(1)  = localApp & "\Programs\Python\Python313\pythonw.exe"
rutas(2)  = localApp & "\Programs\Python\Python312\pythonw.exe"
rutas(3)  = localApp & "\Programs\Python\Python311\pythonw.exe"
rutas(4)  = localApp & "\Programs\Python\Python310\pythonw.exe"
rutas(5)  = localApp & "\Programs\Python\Python314\python.exe"
rutas(6)  = localApp & "\Programs\Python\Python313\python.exe"
rutas(7)  = localApp & "\Programs\Python\Python312\python.exe"
rutas(8)  = localApp & "\Programs\Python\Python311\python.exe"
rutas(9)  = localApp & "\Programs\Python\Python310\python.exe"
rutas(10) = "C:\Python314\pythonw.exe"
rutas(11) = "C:\Python313\pythonw.exe"
rutas(12) = "C:\Python312\pythonw.exe"
rutas(13) = "C:\Python311\pythonw.exe"

pythonPath = ""
For i = 0 To 13
    If oFSO.FileExists(rutas(i)) Then
        pythonPath = rutas(i)
        Exit For
    End If
Next

' Paso 3: intentar PATH
If pythonPath = "" Then
    tmpVer = oShell.ExpandEnvironmentStrings("%TEMP%") & "\pycheck.txt"
    On Error Resume Next
    oShell.Run "cmd /c pythonw --version > " & Chr(34) & tmpVer & Chr(34), 0, True
    On Error GoTo 0
    If oFSO.FileExists(tmpVer) Then
        oFSO.DeleteFile tmpVer
        pythonPath = "pythonw"
    End If
End If
If pythonPath = "" Then
    On Error Resume Next
    resultado = oShell.Run("cmd /c python --version", 0, True)
    On Error GoTo 0
    If resultado = 0 Then pythonPath = "python"
End If

' Paso 4: Python no encontrado
If pythonPath = "" Then
    m1 = "No se encontro Python en este equipo." & vbCrLf & vbCrLf
    m1 = m1 & "Para instalarlo:" & vbCrLf
    m1 = m1 & "  1. Ve a: https://www.python.org/downloads/" & vbCrLf
    m1 = m1 & "  2. Descarga Python 3.12 o superior" & vbCrLf
    m1 = m1 & "  3. Ejecuta el instalador" & vbCrLf
    m1 = m1 & "  4. Marca la casilla Add Python to PATH" & vbCrLf & vbCrLf
    m1 = m1 & "Luego ejecuta instalar_librerias.bat"
    MsgBox m1, vbCritical, "Python no encontrado"
    WScript.Quit 1
End If

' Paso 5: obtener python.exe para verificacion
If InStr(LCase(pythonPath), "pythonw.exe") > 0 Then
    pythonExe = Replace(pythonPath, "pythonw.exe", "python.exe")
ElseIf LCase(pythonPath) = "pythonw" Then
    pythonExe = "python"
Else
    pythonExe = pythonPath
End If

' Paso 6: verificar librerias
tmpScript = oShell.ExpandEnvironmentStrings("%TEMP%") & "\crubc_check.py"
tmpResult = oShell.ExpandEnvironmentStrings("%TEMP%") & "\crubc_out.txt"
Set fsW = oFSO.CreateTextFile(tmpScript, True)
fsW.WriteLine "import importlib.util"
fsW.WriteLine "L = ['pdfplumber', 'docx', 'openpyxl', 'PIL']"
fsW.WriteLine "M = [x for x in L if importlib.util.find_spec(x) is None]"
fsW.WriteLine "print(','.join(M) if M else 'OK')"
fsW.Close
verCmd = "cmd /c " & Chr(34) & pythonExe & Chr(34)
verCmd = verCmd & " " & Chr(34) & tmpScript & Chr(34)
verCmd = verCmd & " > " & Chr(34) & tmpResult & Chr(34)
On Error Resume Next
oShell.Run verCmd, 0, True
On Error GoTo 0
checkOut = "OK"
If oFSO.FileExists(tmpResult) Then
    Set fsR = oFSO.OpenTextFile(tmpResult, 1)
    If Not fsR.AtEndOfStream Then checkOut = Trim(fsR.ReadLine())
    fsR.Close
    oFSO.DeleteFile tmpResult
End If
If oFSO.FileExists(tmpScript) Then oFSO.DeleteFile tmpScript

' Paso 7: librerias faltantes
If checkOut <> "OK" And checkOut <> "" Then
    faltantes = Replace(checkOut, "PIL", "pillow")
    faltantes = Replace(faltantes, "docx", "python-docx")
    m2 = "Faltan librerias necesarias:" & vbCrLf & vbCrLf
    m2 = m2 & "   " & faltantes & vbCrLf & vbCrLf
    m2 = m2 & "Deseas instalarlas ahora?"
    respLib = MsgBox(m2, vbExclamation + vbYesNo, "Librerias faltantes")
    If respLib = vbYes Then
        If oFSO.FileExists(batPath) Then
            oShell.Run Chr(34) & batPath & Chr(34), 1, True
        Else
            MsgBox "No se encontro instalar_librerias.bat", vbCritical, "Error"
            WScript.Quit 1
        End If
    Else
        MsgBox "Ejecuta instalar_librerias.bat e intenta nuevamente.", vbInformation, "Aviso"
        WScript.Quit 0
    End If
End If

' Paso 8: config opcional
If Not oFSO.FileExists(configPath) Then
    m3 = "No se encontro config_indice.json" & vbCrLf & vbCrLf
    m3 = m3 & "Se usara la configuracion por defecto." & vbCrLf
    m3 = m3 & "Deseas continuar?"
    respCfg = MsgBox(m3, vbQuestion + vbYesNo, "Configuracion")
    If respCfg = vbNo Then WScript.Quit 0
End If

' Paso 9: abrir aplicacion sin consola
pythonwPath = Replace(pythonExe, "python.exe", "pythonw.exe")
If oFSO.FileExists(pythonwPath) Then
    pythonCmd = Chr(34) & pythonwPath & Chr(34) & Chr(32) & Chr(34) & scriptPath & Chr(34)
Else
    pythonCmd = Chr(34) & pythonExe & Chr(34) & Chr(32) & Chr(34) & scriptPath & Chr(34)
End If
oShell.Run pythonCmd, 0, False
WScript.Quit 0
