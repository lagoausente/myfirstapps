Este repositorio contiene varias aplicaciones desarrolladas en Python para facilitar tareas comunes, como procesar archivos Excel y renombrar archivos en masa.
📌 Herramientas Incluidas
📊 Procesador de Excel

    Filtra filas por texto específico.
    Selecciona y combina columnas de múltiples archivos Excel.
    Guarda los datos procesados en un único archivo Excel.

🔹 Ubicación: procesador_excel/
🔹 Ejecutable disponible: Sí (procesador_excel.exe)
🗂️ Renombrador de Archivos

    Cambia la extensión de múltiples archivos de una carpeta.
    Permite renombrar archivos con prefijos o sufijos personalizados.

🔹 Ubicación: renombrador_archivos/
🔹 Ejecutable disponible: Sí (renombrador_archivos.exe)
📦 Instalación y Uso

    Clonar el repositorio:

git clone https://github.com/tu_usuario/tu_repositorio.git
cd tu_repositorio

(Opcional) Crear un entorno virtual
Si deseas ejecutar el código en un entorno limpio:

python -m venv venv
source venv/bin/activate  # En Linux/Mac
venv\Scripts\activate     # En Windows

Instalar dependencias:

pip install -r requirements.txt

Ejecutar una aplicación específica:

python procesador_excel/main.py

o

    python renombrador_archivos/main.py

⚙️ Compilación a Ejecutable

Si deseas convertir el código en un .exe:

pyinstaller --onefile --windowed --icon=icono.ico procesador_excel/main.py

📄 Licencia

Este proyecto está bajo la licencia MIT. Puedes usarlo, modificarlo y compartirlo libremente.
✨ Contribuciones

Si tienes ideas o mejoras, ¡puedes enviar un pull request o abrir un issue!