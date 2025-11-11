# Generador de Bingos

Generador de cartones de bingo personalizados con interfaz gráfica, previsualización y exportación a PDF.

##  Captura de pantalla

<img width="700" height="420" alt="image" src="https://github.com/user-attachments/assets/1d30e001-4a3e-4fe5-ab18-dc83430def8b" />

## 🌟 Características

- Generación de cartones de bingo personalizados.
- Previsualización en tiempo real.
- Exportación a PDF con múltiples cartones por hoja.
- Barra de progreso durante la generación.
- Ventana de carga y mensajes de éxito.
- Soporte para hojas tamaño Oficio y Carta.
- Números de serie personalizados.
- Interfaz gráfica moderna y fácil de usar.

## Requisitos

- Python 3.8 o superior
- customtkinter
- reportlab
- PyMuPDF (fitz)
- Pillow

## Instalación

1. Clona el repositorio:
```
git clone https://github.com/tuusuario/generador-bingo.git
cd generador-bingo
```
2. Instala las dependencias:

```
pip install customtkinter reportlab PyMuPDF Pillow
```
3. Ejecuta el programa:
```
python bingo.py
```
## Compilación a .exe

1. Instala PyInstaller:
```
pip install pyinstaller
```
2. Compila el programa:
```
pyinstaller --onefile --windowed bingo.py
```

3. El archivo .exe se generará en la carpeta `dist`.

Gracias por usar mi generador de bingos!  
No dudes en abrir un issue si encuentras algún problema o tienes alguna sugerencia.


