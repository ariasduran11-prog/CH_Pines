@echo off
chcp 65001 >nul
color 0b
title CH Pines v2.0 Pro - Ejecutable Portable
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                    🚀 CH PINES v2.0 PROFESIONAL                   ║
echo ║                     Versión Ejecutable Portable                    ║
echo ║                        Desarrollado por David Arias               ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
echo 💎 Características de esta versión:
echo    ✨ No requiere instalación
echo    📊 Funciona directamente desde cualquier carpeta
echo    🔧 Compatible con Windows 10/11
echo    💾 Portable - funciona desde USB
echo.
echo 🚀 Iniciando CH Pines v2.0 Pro...
echo.

REM Verificar que el ejecutable existe
if not exist "CH_Pines_v2.0_Pro.exe" (
    echo ❌ ERROR: CH_Pines_v2.0_Pro.exe no encontrado
    echo 💡 Asegúrate de ejecutar desde la carpeta correcta
    echo.
    pause
    exit /b 1
)

REM Verificar que la plantilla existe
if not exist "Plantilla.xlsx" (
    echo ⚠️  AVISO: Plantilla.xlsx no encontrada
    echo 💡 El programa creará una plantilla básica automáticamente
    echo.
)

echo ✅ Archivos verificados - Ejecutando aplicación...
echo.

REM Ejecutar el programa
start "" "CH_Pines_v2.0_Pro.exe"

echo 🎉 CH Pines v2.0 Pro iniciado correctamente
echo 👨‍💻 Desarrollado por David Arias (layoutjda@gmail.com)
echo.
echo Presiona cualquier tecla para cerrar esta ventana...
pause >nul