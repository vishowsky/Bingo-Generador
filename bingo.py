import customtkinter as ctk
from PIL import Image
from customtkinter import CTkImage
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import black
from datetime import datetime
import fitz
import random
import sys
import threading
import os

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def get_primary_monitor_geometry(root):
    if sys.platform.startswith("win"):
        import ctypes
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        sw = user32.GetSystemMetrics(0)
        sh = user32.GetSystemMetrics(1)
        return sw, sh
    else:
        return root.winfo_screenwidth(), root.winfo_screenheight()

def center_window(root):
    root.update_idletasks()
    sw, sh = get_primary_monitor_geometry(root)
    w = root.winfo_width()
    h = root.winfo_height()
    x = (sw - w) // 2
    y = (sh - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

def generar_carton_bingo():
    carton = {"B":[], "I":[], "N":[], "G":[], "O":[]}
    rangos = {"B":(1,15), "I":(16,30), "N":(31,45), "G":(46,60), "O":(61,75)}
    for col in carton:
        carton[col] = random.sample(range(rangos[col][0], rangos[col][1]+1),5)
    carton["N"][2] = "FREE"
    return [[carton[col][row] for col in carton] for row in range(5)]

def dibujar_carton(pdf, x, y, ancho_carton, carton):
    ancho_col = ancho_carton / 5
    alto_carton = ancho_carton
    alto_fila = alto_carton / 6
    pdf.setLineWidth(1)
    pdf.setStrokeColor(black)
    pdf.rect(x, y - alto_carton, ancho_carton, alto_carton)
    pdf.setFont("Helvetica-Bold", 13)
    for i, letra in enumerate(['B','I','N','G','O']):
        pdf.drawCentredString(x + i*ancho_col + ancho_col/2, y - alto_fila/2, letra)
    pdf.setLineWidth(0.7)
    for k in range(1, 6):
        pdf.line(x + k*ancho_col, y, x + k*ancho_col, y - alto_carton)
    for k in range(1, 6):
        pdf.line(x, y - k*alto_fila, x + ancho_carton, y - k*alto_fila)
    pdf.setFont("Helvetica", 12)
    for fila in range(5):
        for col in range(5):
            valor = carton[fila][col]
            cx = x + col*ancho_col + ancho_col/2
            cy = y - alto_fila - fila*alto_fila - alto_fila/2 + 2
            if valor == "FREE":
                pdf.setFont("Helvetica-Bold", 12)
                pdf.setFillColor(black)
                pdf.drawCentredString(cx, cy, "FREE")
            else:
                pdf.setFont("Helvetica", 12)
                pdf.setFillColor(black)
                pdf.drawCentredString(cx, cy, str(valor))

def generar_pdf_preview(titulo, filas, cols, size_name, sn):
    ancho_pagina = 215.9 * mm if size_name == "Oficio" else 215.9 * mm
    alto_pagina = 330.2 * mm if size_name == "Oficio" else 279.4 * mm
    separacion = 12
    margen_hoja_alto = 20 * mm
    margen_hoja_ancho = 13 * mm
    ancho_disponible = ancho_pagina - 2 * margen_hoja_ancho - (cols - 1) * separacion
    alto_disponible = alto_pagina - 2 * margen_hoja_alto - (filas - 1) * separacion
    ancho_carton = min(ancho_disponible / cols, alto_disponible / filas)
    alto_carton = ancho_carton
    total_ancho = cols * ancho_carton + (cols - 1) * separacion
    total_alto = filas * alto_carton + (filas - 1) * separacion
    offset_x = (ancho_pagina - total_ancho) / 2
    offset_y = (alto_pagina - total_alto) / 2

    pdf = canvas.Canvas("preview.pdf", pagesize=(ancho_pagina, alto_pagina))
    carton_base = generar_carton_bingo()
    pdf.setFont("Times-BoldItalic", 38 if filas <= 4 else 28)
    y_titulo = alto_pagina - margen_hoja_alto/2
    pdf.drawCentredString(ancho_pagina / 2, y_titulo, titulo)
    for fila in range(filas):
        for col in range(cols):
            x = offset_x + col * (ancho_carton + separacion)
            y = alto_pagina - offset_y - fila * (alto_carton + separacion)
            dibujar_carton(pdf, x, y, ancho_carton, carton_base)
    pdf.setFont("Helvetica", 14)
    pdf.drawRightString(ancho_pagina - 7 * mm, 6 * mm, f"SN {sn}")
    pdf.showPage()
    pdf.save()

    doc = fitz.open("preview.pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=150)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img = img.resize((430, 650), Image.LANCZOS)
    img.save("preview.png")
    doc.close()
    return img

def generate_pdf(titulo, num_hojas, copias_por_hoja, filas, cols, size_name, inicio_sn):
    ancho_pagina = 215.9 * mm if size_name == "Oficio" else 215.9 * mm
    alto_pagina = 330.2 * mm if size_name == "Oficio" else 279.4 * mm
    separacion = 12
    margen_hoja_alto = 20 * mm
    margen_hoja_ancho = 13 * mm
    ancho_disponible = ancho_pagina - 2 * margen_hoja_ancho - (cols - 1) * separacion
    alto_disponible = alto_pagina - 2 * margen_hoja_alto - (filas - 1) * separacion
    ancho_carton = min(ancho_disponible / cols, alto_disponible / filas)
    alto_carton = ancho_carton
    total_ancho = cols * ancho_carton + (cols - 1) * separacion
    total_alto = filas * alto_carton + (filas - 1) * separacion
    offset_x = (ancho_pagina - total_ancho) / 2
    offset_y = (alto_pagina - total_alto) / 2

    fecha = datetime.now().strftime("%Y%m%d")
    filename = f"{titulo.replace(' ', '_')}-{fecha}.pdf"
    pdf = canvas.Canvas(filename, pagesize=(ancho_pagina, alto_pagina))

    total_cartones = num_hojas * copias_por_hoja
    sn_count = inicio_sn

    for i in range(num_hojas):
        carton_base = generar_carton_bingo()
        for copia in range(copias_por_hoja):
            sn_str = f"{sn_count}"
            pdf.setFont("Times-BoldItalic", 38 if filas <= 4 else 28)
            y_titulo = alto_pagina - margen_hoja_alto / 2
            pdf.drawCentredString(ancho_pagina / 2, y_titulo, titulo)
            for fila in range(filas):
                for col in range(cols):
                    x = offset_x + col * (ancho_carton + separacion)
                    y = alto_pagina - offset_y - fila * (alto_carton + separacion)
                    dibujar_carton(
                        pdf,
                        x,
                        y,
                        ancho_carton,
                        carton_base,
                    )
            pdf.setFont("Helvetica", 14)
            pdf.drawRightString(215.9 * mm - 7 * mm, 6 * mm, f"SN {sn_str}")
            pdf.showPage()
            porcentaje = ((i * copias_por_hoja) + (copia + 1)) / total_cartones
        sn_count += 1

    pdf.save()
    return filename

class AboutWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Acerca de")
        self.geometry("500x400")
        self.minsize(500, 400)
        self.maxsize(500, 400)
        self.resizable(False, False)
        self.grab_set()
        info = (
            "Generador de Bingos\n"
            "Versión 1.0\n"
            "Licencia: MIT\n"
            "Desarrollado por Vicente Carreño\n"
            "GitHub: https://github.com/Vishowsky\n"
        )
        ctk.CTkLabel(self, text=info, justify="left", font=("Helvetica", 12)).pack(pady=20)
        ctk.CTkButton(self, text="Cerrar", command=self.destroy).pack(pady=10)
        self.after(0, lambda: center_window(self))

class BingoModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("1200x800")
        self.minsize(1200, 800)
        self.maxsize(1200, 800)
        self.resizable(False, False)
        self.title("Generador de Bingos")
        self.after(0, lambda: center_window(self))
        frame = ctk.CTkFrame(self)
        frame.pack(side="left", fill="y", padx=20, pady=20, expand=False)
        ctk.CTkLabel(frame, text="Título:").pack(pady=7)
        self.titulo_entry = ctk.CTkEntry(frame)
        self.titulo_entry.insert(0, "Bingo")
        self.titulo_entry.pack()
        ctk.CTkLabel(frame, text="Tamaño de hoja:").pack(pady=7)
        self.size_var = ctk.StringVar(value="Oficio")
        ctk.CTkOptionMenu(frame, variable=self.size_var, values=["Oficio", "Carta"]).pack()
        ctk.CTkLabel(frame, text="Cartones por hoja:").pack(pady=7)
        self.cuadricula_var = ctk.StringVar(value="4x3")
        ctk.CTkOptionMenu(frame, variable=self.cuadricula_var, values=["3x3", "4x3"]).pack()
        ctk.CTkLabel(frame, text="Cantidad de hojas:").pack(pady=7)
        self.nhojas_entry = ctk.CTkEntry(frame)
        self.nhojas_entry.insert(0, "1")
        self.nhojas_entry.pack()
        ctk.CTkLabel(frame, text="Copias por hoja:").pack(pady=7)
        self.copias_entry = ctk.CTkEntry(frame)
        self.copias_entry.insert(0, "1")
        self.copias_entry.pack()
        ctk.CTkLabel(frame, text="Inicio correlativo:").pack(pady=7)
        self.sn_entry = ctk.CTkEntry(frame)
        self.sn_entry.insert(0, "1")
        self.sn_entry.pack()
        ctk.CTkButton(frame, text="Previsualizar", command=self.preview).pack(pady=15)
        ctk.CTkButton(frame, text="Generar PDF", command=self.generar_pdf).pack(pady=7)

        self.preview_frame = ctk.CTkFrame(self)
        self.preview_frame.pack(padx=20, pady=20, expand=True)
        self.preview_label = ctk.CTkLabel(self.preview_frame, text="", font=("Helvetica", 16, "bold"))
        self.preview_label.pack()
        self.preview_panel = ctk.CTkLabel(self.preview_frame, text="")
        self.preview_panel.pack(pady=10)

        self.about_btn = ctk.CTkButton(self, text="Acerca de", command=self.show_about)
        self.about_btn.place(relx=0.98, rely=0.98, anchor="se")

    def validate_fields(self):
        try:
            if not self.titulo_entry.get().strip():
                return False, "El título no puede estar vacío."
            nhojas = int(self.nhojas_entry.get())
            copias = int(self.copias_entry.get())
            sn = int(self.sn_entry.get())
            if nhojas < 1 or copias < 1 or sn < 1:
                return False, "Cantidad de hojas, copias y correlativo deben ser mayor que 0."
        except:
            return False, "Ingrese solo números válidos."
        return True, ""

    def preview(self):
        valid, msg = self.validate_fields()
        if not valid:
            ctk.CTkMessagebox(title="Error", message=msg, icon="cancel")
            return
        titulo = self.titulo_entry.get()
        fil, col = (4,3) if self.cuadricula_var.get() == "4x3" else (3,3)
        size_name = self.size_var.get()
        sn = int(self.sn_entry.get())
        img = generar_pdf_preview(titulo, fil, col, size_name, sn)
        imgtk = CTkImage(light_image=img, size=(430, 650))
        self.preview_panel.configure(image=imgtk)
        self.preview_panel.image = imgtk
        self.preview_label.configure(text="Números de ejemplo")

    def generar_pdf(self):
        valid, msg = self.validate_fields()
        if not valid:
            ctk.CTkMessagebox(title="Error", message=msg, icon="cancel")
            return
        titulo = self.titulo_entry.get()
        num_hojas = int(self.nhojas_entry.get())
        copias_por_hoja = int(self.copias_entry.get())
        inicio_sn = int(self.sn_entry.get())
        filas, cols = (4, 3) if self.cuadricula_var.get() == "4x3" else (3, 3)
        size_name = self.size_var.get()

        total_cartones = num_hojas * copias_por_hoja
        fecha = datetime.now().strftime("%Y%m%d")
        filename = f"{titulo.replace(' ', '_')}-{fecha}.pdf"
        stop_event = threading.Event()

        def run_generation():
            try:
                pdf = canvas.Canvas(filename, pagesize=(215.9 * mm, 330.2 * mm if size_name == "Oficio" else 279.4 * mm))
                sn_count = inicio_sn
                for i in range(num_hojas):
                    if stop_event.is_set():
                        break
                    carton_base = generar_carton_bingo()
                    for copia in range(copias_por_hoja):
                        if stop_event.is_set():
                            break
                        sn_str = f"{sn_count}"
                        pdf.setFont("Times-BoldItalic", 38 if filas <= 4 else 28)
                        y_titulo = 330.2 * mm - 20 * mm / 2 if size_name == "Oficio" else 279.4 * mm - 20 * mm / 2
                        pdf.drawCentredString(215.9 * mm / 2, y_titulo, titulo)
                        for fila in range(filas):
                            for col in range(cols):
                                x = 13 * mm + col * (min((215.9 * mm - 26 * mm) / cols, (330.2 * mm if size_name == "Oficio" else 279.4 * mm - 40 * mm) / filas) + 12)
                                y = (330.2 * mm if size_name == "Oficio" else 279.4 * mm) - 20 * mm - fila * (min((215.9 * mm - 26 * mm) / cols, (330.2 * mm if size_name == "Oficio" else 279.4 * mm - 40 * mm) / filas) + 12)
                                dibujar_carton(pdf, x, y, min((215.9 * mm - 26 * mm) / cols, (330.2 * mm if size_name == "Oficio" else 279.4 * mm - 40 * mm) / filas), carton_base)
                        pdf.setFont("Helvetica", 14)
                        pdf.drawRightString(215.9 * mm - 7 * mm, 6 * mm, f"SN {sn_str}")
                        pdf.showPage()
                        porcentaje = ((i * copias_por_hoja) + (copia + 1)) / total_cartones
                        progress_bar.set(porcentaje)
                        progress_label.configure(text=f"{int(porcentaje * 100)}%")
                        loading.update_idletasks()
                    sn_count += 1
                if not stop_event.is_set():
                    pdf.save()
                    self.after(0, lambda: mostrar_popup_exito(filename))
                else:
                    if os.path.exists(filename):
                        os.remove(filename)
                    self.after(0, lambda: mostrar_popup_cancelado())
            except Exception as e:
                if os.path.exists(filename):
                    os.remove(filename)
                    self.after(0, lambda: mostrar_popup_error(str(e)))

        def mostrar_popup_exito(filename):
            okpop = ctk.CTkToplevel(self)
            okpop.title("¡Bingo listo!")
            okpop.geometry("350x120")
            okpop.grab_set()
            okpop.resizable(False, False)
            okpop.after(0, lambda: center_window(okpop))
            label = ctk.CTkLabel(
                okpop,
                text=f"Bingo generado con éxito:\n{filename}",
                font=("Helvetica", 16),
                justify="center",
            )
            label.pack(expand=True, pady=20)
            okbtn = ctk.CTkButton(okpop, text="Aceptar", command=okpop.destroy)
            okbtn.pack(pady=5)

        def mostrar_popup_cancelado():
            okpop = ctk.CTkToplevel(self)
            okpop.title("Generación cancelada")
            okpop.geometry("350x120")
            okpop.grab_set()
            okpop.resizable(False, False)
            okpop.after(0, lambda: center_window(okpop))
            label = ctk.CTkLabel(
                okpop,
                text="La generación fue cancelada y el archivo eliminado.",
                font=("Helvetica", 16),
                justify="center",
            )
            label.pack(expand=True, pady=20)
            okbtn = ctk.CTkButton(okpop, text="Aceptar", command=okpop.destroy)
            okbtn.pack(pady=5)

        def mostrar_popup_error(msg):
            okpop = ctk.CTkToplevel(self)
            okpop.title("Error")
            okpop.geometry("350x120")
            okpop.grab_set()
            okpop.resizable(False, False)
            okpop.after(0, lambda: center_window(okpop))
            label = ctk.CTkLabel(
                okpop,
                text=f"Ocurrió un error:\n{msg}",
                font=("Helvetica", 16),
                justify="center",
            )
            label.pack(expand=True, pady=20)
            okbtn = ctk.CTkButton(okpop, text="Aceptar", command=okpop.destroy)
            okbtn.pack(pady=5)

        loading = ctk.CTkToplevel(self)
        loading.title("Generando Bingos")
        loading.geometry("400x180")
        loading.grab_set()
        loading.resizable(False, False)
        loading.after(0, lambda: center_window(loading))
        wrapper = ctk.CTkFrame(loading)
        wrapper.pack(expand=True, fill="both", padx=20, pady=20)
        label_main = ctk.CTkLabel(
            wrapper, text="Generando cartones...", font=("Helvetica", 18)
        )
        label_main.pack(pady=10)
        progress_bar = ctk.CTkProgressBar(wrapper, width=300)
        progress_bar.set(0)
        progress_bar.pack(pady=10)
        progress_label = ctk.CTkLabel(wrapper, text="0%", font=("Helvetica", 14))
        progress_label.pack()
        cancel_btn = ctk.CTkButton(
            wrapper,
            text="Cancelar",
            command=lambda: stop_event.set(),
            fg_color="red",
            hover_color="darkred"
        )
        cancel_btn.pack(pady=10)

        def update_button():
            if not stop_event.is_set():
                cancel_btn.configure(text="Cerrar", command=loading.destroy, fg_color="green", hover_color="darkgreen")

        loading.after(100, update_button)
        threading.Thread(target=run_generation, daemon=True).start()

    def show_about(self):
        AboutWindow(self)

if __name__ == "__main__":
    app = BingoModernApp()
    app.mainloop()
