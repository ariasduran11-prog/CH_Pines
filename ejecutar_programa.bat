@echo off
chcp 65001 >nul
color 0b
title CH Pines v2.0 - Generador Profesional MikroTik
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                    🚀 CH PINES v2.0 PROFESIONAL                   ║
echo ║                     Generador MikroTik Optimizado                  ║
echo ║                        Desarrollado por David Arias               ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
echo 🎯 Versión 2.0 con características profesionales:
echo    ✨ Diseño limpio y moderno
echo    📊 Exportación Excel perfecta
echo    🧹 Código optimizado y limpio
echo    🔧 Conexión manual estable
echo.
echo 🔍 Verificando sistema...

REM Verificar Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ ERROR: Python no está instalado
    echo 💡 Ejecuta 'instalar.bat' primero
    echo.
    pause
    exit /b 1
)

REM Verificar archivo principal
if not exist "winbox_style_generator.py" (
    echo ❌ ERROR: Archivo principal no encontrado
    echo 💡 Asegúrate de estar en la carpeta correcta
    echo.
    pause
    exit /b 1
)

echo ✅ Sistema verificado - Iniciando aplicación...
echo.
echo ════════════════════════════════════════════════════════════════════
echo.

REM Ejecutar el programa
python winbox_style_generator.py

echo.
echo ════════════════════════════════════════════════════════════════════
echo 🎉 Sesión finalizada
echo 👨‍💻 Desarrollado por David Arias (layoutjda@gmail.com)
echo 📞 ¿Necesitas soporte? ¡Contáctame!
echo.
echo Presiona cualquier tecla para cerrar...
pause >nul