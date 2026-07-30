; export_articulos.ahk  (AutoHotkey v2 — descargalo gratis de autohotkey.com)
;
; Automatiza en GDS: Importar / Exportar > Exportación Rápida,
; guardando siempre en la misma ruta fija, y después llama al script
; que sube ese archivo a tu web.
;
; SEGURIDAD: si detecta que estuviste usando el mouse o el teclado
; hace poco, NO hace nada (para no interrumpirte ni escribir en el
; lugar equivocado). Solo actúa si la PC estuvo quieta un rato.
;
; REQUISITO: GDS tiene que estar abierto y con la pantalla
; "Modificación de Artículos" ya visible (la de la grilla de artículos)
; cuando este script corre. Si GDS está cerrado o en otra pantalla, no
; va a funcionar solo — este script no abre GDS ni inicia sesión por vos.
;
; Antes de usarlo:
; 1) Creá la carpeta C:\GDS_export en tu PC.
; 2) Revisá que exportPath abajo apunte ahí.
; 3) Completá tu ADMIN_SECRET en upload_articulos.ps1 (archivo aparte).

exportPath := "C:\GDS_export\articulos.xls"
minutosDeInactividadRequeridos := 3

; Si hubo actividad de mouse/teclado en los últimos X minutos, no hace nada.
idleMs := A_TimeIdlePhysical
if (idleMs < minutosDeInactividadRequeridos * 60000) {
    ExitApp  ; se salta esta corrida en silencio; a la hora siguiente reintenta
}

if !WinExist("Modificación de Artículos") {
    ExitApp  ; GDS no está en esa pantalla ahora mismo; no interrumpe nada, reintenta la próxima hora
}

WinActivate("Modificación de Artículos")
WinWaitActive("Modificación de Artículos",, 5)

; Selecciona el menú nativo: Importar / Exportar > Exportación Rápida
WinMenuSelectItem("Modificación de Artículos",, "Importar / Exportar", "Exportación Rápida")

; Espera a que aparezca el diálogo "Guardar en"
if WinWait("Guardar en",, 8) {
    WinActivate("Guardar en")
    Sleep(400)
    ; Reemplaza el nombre/ruta completa y guarda
    Send("^a")
    SendText(exportPath)
    Sleep(200)
    Send("{Enter}")
    Sleep(600)
    ; Si pregunta por sobreescribir el archivo existente, confirma
    if WinWait("Confirmar",, 2) {
        Send("{Enter}")
    }
} else {
    MsgBox("No apareció el diálogo para guardar. Revisá que GDS esté en la pantalla correcta.")
    ExitApp
}

; Espera a que el archivo se termine de escribir en disco
Loop 20 {
    if FileExist(exportPath)
        break
    Sleep(500)
}

; Sube el archivo al backend
RunWait('powershell.exe -ExecutionPolicy Bypass -File "' A_ScriptDir '\upload_articulos.ps1"',, "Hide")
