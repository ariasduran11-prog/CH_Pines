@echo off
chcp 65001 >nul
color 0b
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                    🚀 CH PINES v2.0 - INSTALADOR                  ║
echo ║                      Generador MikroTik Profesional                ║
echo ║                        Desarrollado por David Arias               ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
echo 🎯 Instalando la versión más avanzada con:
echo   ✨ Diseño profesional limpio
echo   📊 Exportación Excel optimizada (sin páginas extra)
echo   🎨 Interfaz moderna sin iconos innecesarios
echo   🧹 Código optimizado y limpio
echo   🔧 Solo conexión manual (más estable)
echo.
echo ════════════════════════════════════════════════════════════════════

echo.
echo [1/5] 🔍 Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ ERROR: Python no está instalado
    echo.
    echo 📥 Descarga Python desde: https://python.org/downloads
    echo ⚠️  Durante la instalación marca: "Add Python to PATH"
    echo 💡 Luego ejecuta este instalador nuevamente
    echo.
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo ✅ Python %PYTHON_VERSION% encontrado

echo.
echo [2/5] 📦 Actualizando pip...
python -m pip install --upgrade pip --quiet
echo ✅ pip actualizado correctamente

echo.
echo [3/5] 🔧 Instalando dependencias profesionales...
echo    📡 Instalando paramiko (conexión SSH)...
pip install paramiko --quiet
if %errorlevel% neq 0 (
    echo ❌ Error instalando paramiko
    pause
    exit /b 1
)
echo    🔐 Instalando cryptography (seguridad)...
pip install cryptography --quiet
if %errorlevel% neq 0 (
    echo ❌ Error instalando cryptography
    pause
    exit /b 1
)
echo    📊 Instalando openpyxl (Excel)...
pip install openpyxl --quiet
if %errorlevel% neq 0 (
    echo ❌ Error instalando openpyxl
    pause
    exit /b 1
)
echo ✅ Todas las dependencias instaladas

echo.
echo [4/5] 🧪 Verificando instalación completa...
python -c "import paramiko; import cryptography; import openpyxl; import tkinter; print('✅ Verificación exitosa: Todas las librerías funcionando')" 2>nul
if %errorlevel% neq 0 (
    echo ❌ Error en la verificación
    echo 💡 Algunas librerías pueden no estar instaladas correctamente
    pause
    exit /b 1
)

echo.
echo [5/5] 📋 Verificando archivos del programa...
if not exist "winbox_style_generator.py" (
    echo ❌ ERROR: Archivo principal no encontrado
    echo 💡 Asegúrate de ejecutar desde la carpeta del programa
    pause
    exit /b 1
)
if not exist "Plantilla.xlsx" (
    echo ⚠️  AVISO: Plantilla.xlsx no encontrada
    echo 💡 El programa creará una plantilla básica automáticamente
)
echo ✅ Archivos verificados

echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                    🎉 INSTALACIÓN COMPLETADA                      ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
echo 🚀 Para ejecutar el programa:
echo    📝 Opción 1: python winbox_style_generator.py
echo    🎯 Opción 2: ejecutar_programa.bat (recomendado)
echo.
echo 💎 Características instaladas v2.0:
echo    ✨ Interfaz profesional limpia
echo    📊 Exportación Excel perfecta (sin páginas extra)
echo    🎨 Diseño moderno optimizado
echo    🧹 Código limpio y optimizado
echo    🔧 Conexión manual estable
echo.
echo 👨‍💻 Desarrollado por: David Arias (layoutjda@gmail.com)
echo 📞 ¿Necesitas soporte? ¡Contáctame!
echo.
echo Presiona cualquier tecla para continuar...
pause >nul