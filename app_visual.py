from datetime import datetime

import customtkinter as ctk
from tkinter import messagebox

from main import (
    actualizar_venta_desde_gui,
    actualizar_prenda_desde_gui,
    buscar_ingreso_por_codigo,
    eliminar_ingresos_desde_gui,
    obtener_prenda_desde_gui,
    obtener_venta_desde_gui,
    calcular_prendas_vencidas,
    calcular_resumen_general,
    calcular_rendicion_proveedora,
    calcular_ventas_pendientes,
    crear_proveedora_desde_gui,
    exportar_lote_remarque_excel,
    exportar_rendicion_excel,
    guardar_decisiones_remarque_desde_gui,
    guardar_lote_remarque_desde_gui,
    guardar_ingreso_desde_gui,
    guardar_venta_desde_gui,
    normalizar_obs_descuento_proveedora,
    obtener_lote_remarque_proveedora_desde_gui,
    obtener_lote_decision_remarque_desde_gui,
    obtener_proveedora_desde_gui,
    obtener_ruta_en_base,
    obtener_todas_las_proveedoras_desde_gui,
    registrar_devolucion_desde_gui,
    reversar_ventas_desde_gui,
    validar_venta_pendiente_por_codigo,
)

# CONFIGURACION GENERAL
ctk.set_appearance_mode("Light")   # "Light", "Dark" o "System"
ctk.set_default_color_theme("blue")  # "blue", "green" o "dark-blue"

# ESTILOS GENERALES
COLOR_FONDO = "white"
COLOR_TEXTO = "black"
COLOR_HOVER = "#F2F2F2"

FUENTE_TITULO = ("Arial", 28, "bold")
FUENTE_SUBTITULO = ("Arial", 16)
FUENTE_BOTON = ("Arial", 18)
FUENTE_FLECHA = ("Arial", 22)
FUENTE_FORMULARIO_TITULO = ("Arial", 24, "bold")


class AppFashionReset(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Fashion Reset")
        self.geometry("900x600")
        self.configure(fg_color=COLOR_FONDO)
        self._configurar_icono()

        self.rendicion_actual = None

        self.label_titulo = ctk.CTkLabel(
            self,
            text="FASHION RESET",
            font=FUENTE_TITULO
        )
        self.label_titulo.pack(pady=(10, 5))

        self.frame_botones = ctk.CTkFrame(self, fg_color=COLOR_FONDO)
        self.frame_botones.pack(pady=30, padx=30, fill="both", expand=True)

        self.frame_ingreso = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_venta = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_consultas = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_resumen_general = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_pendientes_validacion = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_proveedoras = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_editar = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_devolucion = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)
        self.frame_rendicion = ctk.CTkScrollableFrame(self, fg_color=COLOR_FONDO)

        boton_ingreso = ctk.CTkButton(
            self.frame_botones,
            text="Cargar Ingreso",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.abrir_ventana_ingreso
        )
        boton_ingreso.pack(pady=(10, 0))

        linea_ingreso = ctk.CTkFrame(
            self.frame_botones,
            width=220,
            height=1,
            fg_color=COLOR_TEXTO
        )
        linea_ingreso.pack(pady=(0, 10))

        self.boton_venta = ctk.CTkButton(
            self.frame_botones,
            text="Cargar Venta",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.abrir_ventana_venta
        )
        self.boton_venta.pack(pady=10)

        self.boton_consultas = ctk.CTkButton(
            self.frame_botones,
            text="Consultas y Rendición",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.abrir_ventana_consultas
        )
        self.boton_consultas.pack(pady=10)

        self.boton_proveedoras = ctk.CTkButton(
            self.frame_botones,
            text="Proveedoras",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.abrir_ventana_proveedoras
        )
        self.boton_proveedoras.pack(pady=10)

        self.boton_editar = ctk.CTkButton(
            self.frame_botones,
            text="Editar",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.abrir_ventana_editar
        )
        self.boton_editar.pack(pady=10)

        self.boton_devolucion = ctk.CTkButton(
            self.frame_botones,
            text="Devolución",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.abrir_modulo_devolucion
        )
        self.boton_devolucion.pack(pady=10)

        self.boton_salir = ctk.CTkButton(
            self.frame_botones,
            text="Salir",
            font=FUENTE_BOTON,
            width=220,
            height=40,
            fg_color="transparent",
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=self.destroy
        )
        self.boton_salir.pack(pady=20)

    def _configurar_icono(self):
        ruta_icono = obtener_ruta_en_base("Icono Fashion Reset.ico")
        try:
            self.iconbitmap(ruta_icono)
        except Exception:
            pass

    def _ocultar_frames_secundarios(self):
        self.frame_ingreso.pack_forget()
        self.frame_venta.pack_forget()
        self.frame_consultas.pack_forget()
        self.frame_resumen_general.pack_forget()
        self.frame_pendientes_validacion.pack_forget()
        self.frame_proveedoras.pack_forget()
        self.frame_editar.pack_forget()
        self.frame_devolucion.pack_forget()
        self.frame_rendicion.pack_forget()

    def _crear_encabezado(self, frame_padre, titulo, comando_volver):
        frame_encabezado = ctk.CTkFrame(frame_padre, fg_color=COLOR_FONDO)
        frame_encabezado.pack(fill="x", padx=10, pady=(5, 5))

        boton_atras = ctk.CTkButton(
            frame_encabezado,
            text="←",
            font=FUENTE_FLECHA,
            width=40,
            height=40,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=0,
            corner_radius=0,
            command=comando_volver
        )
        boton_atras.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="w")

        label_titulo = ctk.CTkLabel(
            frame_encabezado,
            text=titulo,
            font=FUENTE_TITULO,
            text_color=COLOR_TEXTO
        )
        label_titulo.grid(row=0, column=1, sticky="w")

        return frame_encabezado

    def _resetear_scroll(self, frame):
        try:
            self.after(0, lambda: frame._parent_canvas.yview_moveto(0))
        except Exception:
            pass

    def _aplicar_mascara_fecha(self, entry):
        entry.bind("<KeyRelease>", lambda event, e=entry: self._formatear_fecha_en_entry(e))

    def _formatear_fecha_en_entry(self, entry):
        texto = entry.get()
        digitos = "".join(caracter for caracter in texto if caracter.isdigit())[:8]

        if len(digitos) <= 2:
            formateado = digitos
        elif len(digitos) <= 4:
            formateado = f"{digitos[:2]}/{digitos[2:]}"
        else:
            formateado = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"

        if texto == formateado:
            return

        entry.delete(0, "end")
        entry.insert(0, formateado)
        entry.icursor("end")

    def abrir_ventana_ingreso(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_ingreso.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_ingreso.winfo_children():
            widget.destroy()

        self.filas_ingreso = []
        self.fila_actual_ingreso = 1

        self._crear_encabezado(self.frame_ingreso, "CARGAR INGRESO", self.volver_menu_principal)

        frame_lote = ctk.CTkFrame(self.frame_ingreso, fg_color=COLOR_FONDO)
        frame_lote.pack(fill="x", padx=20, pady=10)

        label_fecha = ctk.CTkLabel(frame_lote, text="Fecha ingreso", text_color=COLOR_TEXTO)
        label_fecha.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        entry_fecha = ctk.CTkEntry(frame_lote, width=140)
        entry_fecha.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self._aplicar_mascara_fecha(entry_fecha)

        label_proveedora = ctk.CTkLabel(frame_lote, text="Código proveedora", text_color=COLOR_TEXTO)
        label_proveedora.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        entry_proveedora = ctk.CTkEntry(frame_lote, width=160)
        entry_proveedora.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.frame_tabla_ingreso = ctk.CTkFrame(self.frame_ingreso, fg_color=COLOR_FONDO)
        self.frame_tabla_ingreso.pack(fill="x", padx=20, pady=10)

        encabezados = ["NÚMERO", "ARTÍCULO", "MARCA", "TALLE", "COLOR", "PRECIO", "OBS"]
        for col, encabezado in enumerate(encabezados):
            label = ctk.CTkLabel(
                self.frame_tabla_ingreso,
                text=encabezado,
                text_color=COLOR_TEXTO,
                font=FUENTE_BOTON
            )
            label.grid(row=0, column=col, padx=5, pady=5, sticky="w")

        self.agregar_filas_ingreso(10)

        frame_botones_inferiores = ctk.CTkFrame(self.frame_ingreso, fg_color=COLOR_FONDO)
        frame_botones_inferiores.pack(fill="x", padx=20, pady=15)

        boton_agregar_filas = ctk.CTkButton(
            frame_botones_inferiores,
            text="Agregar 10 filas más",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=lambda: self.agregar_filas_ingreso(10)
        )
        boton_agregar_filas.pack(side="left", padx=10)

        boton_guardar = ctk.CTkButton(
            frame_botones_inferiores,
            text="Guardar lote",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=lambda: self.guardar_lote_ingreso(entry_fecha, entry_proveedora)
        )
        boton_guardar.pack(side="right", padx=10)

    def abrir_ventana_venta(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_venta.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_venta.winfo_children():
            widget.destroy()

        self.filas_venta = []
        self.fila_actual_venta = 1

        self._crear_encabezado(self.frame_venta, "CARGAR VENTA", self.volver_menu_principal)

        frame_lote = ctk.CTkFrame(self.frame_venta, fg_color=COLOR_FONDO)
        frame_lote.pack(fill="x", padx=20, pady=10)

        label_fecha = ctk.CTkLabel(frame_lote, text="Fecha venta", text_color=COLOR_TEXTO)
        label_fecha.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        entry_fecha = ctk.CTkEntry(frame_lote, width=140)
        entry_fecha.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self._aplicar_mascara_fecha(entry_fecha)

        label_cliente = ctk.CTkLabel(frame_lote, text="Cliente", text_color=COLOR_TEXTO)
        label_cliente.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        entry_cliente = ctk.CTkEntry(frame_lote, width=160)
        entry_cliente.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.tipo_pago_opciones = [
            "TRANSFERENCIA",
            "EFECTIVO",
            "DESCUENTO A PROVEEDORA",
            "OTRO"
        ]
        self.validacion_opciones = ["PENDIENTE", "PAGADO"]

        self.frame_tabla_venta = ctk.CTkFrame(self.frame_venta, fg_color=COLOR_FONDO)
        self.frame_tabla_venta.pack(fill="x", padx=20, pady=(10, 0))

        encabezados = [
            "CÓDIGO PRENDA",
            "ARTÍCULO",
            "MARCA",
            "TALLE",
            "COLOR",
            "PRECIO VENTA",
            "TIPO PAGO",
            "VALIDACIÓN",
            "OBS VENTA"
        ]
        for col, encabezado in enumerate(encabezados):
            label = ctk.CTkLabel(
                self.frame_tabla_venta,
                text=encabezado,
                text_color=COLOR_TEXTO,
                font=FUENTE_BOTON
            )
            label.grid(row=0, column=col, padx=5, pady=5, sticky="w")

        self.agregar_filas_venta(10)

        frame_botones_inferiores = ctk.CTkFrame(self.frame_venta, fg_color=COLOR_FONDO)
        frame_botones_inferiores.pack(fill="x", padx=20, pady=15)
        frame_botones_inferiores.grid_columnconfigure(0, weight=1)

        frame_acciones_izquierda = ctk.CTkFrame(frame_botones_inferiores, fg_color=COLOR_FONDO)
        frame_acciones_izquierda.grid(row=0, column=0, sticky="w")

        self.label_total_venta = ctk.CTkLabel(
            frame_acciones_izquierda,
            text="TOTAL DEL LOTE: $0",
            text_color=COLOR_TEXTO,
            font=FUENTE_BOTON
        )
        self.label_total_venta.pack(side="left", padx=(10, 12))

        boton_agregar_filas = ctk.CTkButton(
            frame_acciones_izquierda,
            text="Agregar 10 filas más",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=lambda: self.agregar_filas_venta(10)
        )
        boton_agregar_filas.pack(side="left", padx=0)

        boton_guardar = ctk.CTkButton(
            frame_botones_inferiores,
            text="Guardar lote",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=lambda: self.guardar_lote_venta(entry_fecha, entry_cliente)
        )
        boton_guardar.grid(row=0, column=1, padx=10, sticky="e")
        self._resetear_scroll(self.frame_venta)

    def abrir_ventana_consultas(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_consultas.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_consultas.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_consultas, "CONSULTAS Y RENDICIÓN", self.volver_menu_principal)

        frame_contenido = ctk.CTkFrame(self.frame_consultas, fg_color=COLOR_FONDO)
        frame_contenido.pack(fill="both", expand=True, padx=20, pady=20)

        label_info = ctk.CTkLabel(
            frame_contenido,
            text="Elegí la consulta que querés abrir.",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        )
        label_info.pack(anchor="w", pady=(0, 20))

        boton_rendicion = ctk.CTkButton(
            frame_contenido,
            text="Rendición de cuentas",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_ventana_rendicion
        )
        boton_rendicion.pack(anchor="w", pady=(0, 10))

        boton_resumen_general = ctk.CTkButton(
            frame_contenido,
            text="Resumen general del mes",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_ventana_resumen_general
        )
        boton_resumen_general.pack(anchor="w", pady=(0, 10))

        boton_pendientes = ctk.CTkButton(
            frame_contenido,
            text="Ventas pendientes de validación",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_ventana_pendientes_validacion
        )
        boton_pendientes.pack(anchor="w", pady=(0, 10))

        label_pendiente = ctk.CTkLabel(
            frame_contenido,
            text="Las demás consultas las podemos ir agregando una por una.",
            font=FUENTE_SUBTITULO,
            text_color="#555555"
        )
        label_pendiente.pack(anchor="w", pady=(10, 0))

    def abrir_ventana_proveedoras(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_proveedoras.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_proveedoras.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_proveedoras, "PROVEEDORAS", self.volver_menu_principal)

        frame_contenido = ctk.CTkFrame(self.frame_proveedoras, fg_color=COLOR_FONDO)
        frame_contenido.pack(fill="both", expand=True, padx=20, pady=20)

        label_info = ctk.CTkLabel(
            frame_contenido,
            text="Elegí la acción que querés realizar.",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        )
        label_info.pack(anchor="w", pady=(0, 20))

        boton_nueva = ctk.CTkButton(
            frame_contenido,
            text="Nueva proveedora",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_formulario_proveedora
        )
        boton_nueva.pack(anchor="w", pady=(0, 10))

        boton_ver = ctk.CTkButton(
            frame_contenido,
            text="Ver proveedora",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_consulta_proveedora
        )
        boton_ver.pack(anchor="w", pady=(0, 10))

    def abrir_ventana_editar(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_editar.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_editar.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_editar, "EDITAR", self.volver_menu_principal)

        frame_contenido = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_contenido.pack(fill="both", expand=True, padx=20, pady=20)

        label_info = ctk.CTkLabel(
            frame_contenido,
            text="Elegí lo que querés editar.",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        )
        label_info.pack(anchor="w", pady=(0, 20))

        boton_editar_venta = ctk.CTkButton(
            frame_contenido,
            text="Editar venta",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_editar_venta
        )
        boton_editar_venta.pack(anchor="w", pady=(0, 10))

        boton_editar_prenda = ctk.CTkButton(
            frame_contenido,
            text="Editar prenda",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_editar_prenda
        )
        boton_editar_prenda.pack(anchor="w", pady=(0, 10))

        boton_eliminar_prenda = ctk.CTkButton(
            frame_contenido,
            text="Eliminar prenda",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_eliminar_prendas
        )
        boton_eliminar_prenda.pack(anchor="w", pady=(0, 10))

        boton_reversar_venta = ctk.CTkButton(
            frame_contenido,
            text="Reversar venta",
            font=FUENTE_BOTON,
            width=280,
            height=42,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.abrir_reversar_ventas
        )
        boton_reversar_venta.pack(anchor="w", pady=(0, 10))

    def abrir_modulo_devolucion(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_devolucion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_devolucion.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_devolucion, "DEVOLUCIÓN", self.volver_menu_principal)

        frame_contenido = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_contenido.pack(fill="both", expand=True, padx=20, pady=20)

        label_info = ctk.CTkLabel(
            frame_contenido,
            text="Elegí la parte del proceso que querés trabajar.",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        )
        label_info.pack(anchor="w", pady=(0, 20))

        botones = [
            ("Ver prendas vencidas", self.abrir_ventana_prendas_vencidas),
            ("Remarcar", self.abrir_ventana_remarque),
            ("Aprobación de remarque", self.abrir_ventana_aprobacion_remarque),
            ("Devolución", self.abrir_ventana_devolucion),
        ]

        for texto, comando in botones:
            boton = ctk.CTkButton(
                frame_contenido,
                text=texto,
                font=FUENTE_BOTON,
                width=320,
                height=42,
                fg_color=COLOR_FONDO,
                hover_color=COLOR_HOVER,
                text_color=COLOR_TEXTO,
                border_width=1,
                border_color=COLOR_TEXTO,
                command=comando
            )
            boton.pack(anchor="w", pady=(0, 10))

    def abrir_ventana_prendas_vencidas(self):
        self._ocultar_frames_secundarios()
        self.frame_devolucion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_devolucion.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_devolucion, "PRENDAS VENCIDAS", self.abrir_modulo_devolucion)

        frame_acciones = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_acciones.pack(fill="x", padx=20, pady=10)

        boton_cargar = ctk.CTkButton(
            frame_acciones,
            text="Cargar prendas vencidas",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.cargar_prendas_vencidas
        )
        boton_cargar.pack(anchor="w")

        self.label_estado_prendas_vencidas = ctk.CTkLabel(
            self.frame_devolucion,
            text="Todavía no cargaste el listado de prendas vencidas.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_prendas_vencidas.pack(fill="x", padx=20, pady=(0, 10))

        self.prendas_vencidas_vars = {
            "cantidad_prendas": ctk.StringVar(value="-")
        }

        frame_resumen = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_resumen.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(
            frame_resumen,
            text="Cantidad prendas",
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=0, column=0, padx=10, pady=(0, 4), sticky="w")
        ctk.CTkLabel(
            frame_resumen,
            textvariable=self.prendas_vencidas_vars["cantidad_prendas"],
            text_color=COLOR_TEXTO,
            font=FUENTE_BOTON
        ).grid(row=1, column=0, padx=10, pady=(0, 8), sticky="w")

        self.textbox_prendas_vencidas = ctk.CTkTextbox(
            self.frame_devolucion,
            height=380,
            font=("Consolas", 14),
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color="#D9D9D9"
        )
        self.textbox_prendas_vencidas.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self._mostrar_prendas_vencidas("Todavía no hay resultados para mostrar.")

    def abrir_ventana_remarque(self):
        self._ocultar_frames_secundarios()
        self.frame_devolucion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_devolucion.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_devolucion, "REMARCAR", self.abrir_modulo_devolucion)

        frame_busqueda = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_busqueda.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_busqueda, text="Código proveedora", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_remarque_proveedora = ctk.CTkEntry(frame_busqueda, width=180)
        self.entry_remarque_proveedora.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar lote",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_lote_remarque
        )
        boton_buscar.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_estado_remarque = ctk.CTkLabel(
            self.frame_devolucion,
            text="Buscá una proveedora para cargar el lote de remarque.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_remarque.pack(fill="x", padx=20, pady=(0, 10))

        frame_formulario = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_formulario.pack(fill="x", padx=20, pady=10)

        self.remarque_vars = {
            "nombre_proveedora": ctk.StringVar(value="-"),
            "cantidad_prendas": ctk.StringVar(value="-"),
        }

        ctk.CTkLabel(frame_formulario, text="Proveedora", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=(6, 2), sticky="w"
        )
        ctk.CTkLabel(
            frame_formulario,
            textvariable=self.remarque_vars["nombre_proveedora"],
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=1, column=0, padx=10, pady=(0, 8), sticky="w")

        ctk.CTkLabel(frame_formulario, text="Prendas vencidas", text_color=COLOR_TEXTO).grid(
            row=0, column=1, padx=10, pady=(6, 2), sticky="w"
        )
        ctk.CTkLabel(
            frame_formulario,
            textvariable=self.remarque_vars["cantidad_prendas"],
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=1, column=1, padx=10, pady=(0, 8), sticky="w")

        self.label_info_lote_remarque = ctk.CTkLabel(
            frame_formulario,
            text="La fecha del lote se guardará automáticamente con la fecha de hoy.",
            text_color="#666666",
            font=("Arial", 14)
        )
        self.label_info_lote_remarque.grid(row=1, column=2, padx=10, pady=(0, 8), sticky="w")

        self.frame_tabla_remarque = ctk.CTkScrollableFrame(self.frame_devolucion, fg_color=COLOR_FONDO, height=320)
        self.frame_tabla_remarque.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        encabezados = [
            ("Código", 0, 90),
            ("Artículo", 1, 140),
            ("Marca", 2, 100),
            ("Talle", 3, 70),
            ("Color", 4, 90),
            ("Precio actual", 5, 100),
            ("Días", 6, 60),
            ("Último remarque", 7, 110),
            ("Precio remarcado", 8, 120),
        ]
        for texto, columna, ancho in encabezados:
            ctk.CTkLabel(
                self.frame_tabla_remarque,
                text=texto,
                text_color=COLOR_TEXTO,
                width=ancho
            ).grid(row=0, column=columna, padx=4, pady=(0, 6), sticky="w")

        self.filas_remarque = []
        self.fila_actual_remarque = 1
        self.lote_remarque_actual = None

        frame_acciones = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_acciones.pack(fill="x", padx=20, pady=(0, 20))

        boton_guardar = ctk.CTkButton(
            frame_acciones,
            text="Guardar lote",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.guardar_lote_remarque
        )
        boton_guardar.pack(side="right")

        boton_exportar = ctk.CTkButton(
            frame_acciones,
            text="Exportar lote",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.exportar_lote_remarque_actual
        )
        boton_exportar.pack(side="right", padx=(0, 10))

    def abrir_ventana_devolucion(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_devolucion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_devolucion.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_devolucion, "DEVOLUCIÓN", self.volver_menu_principal)

        frame_busqueda = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_busqueda.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_busqueda, text="Código prenda", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_devolucion_codigo = ctk.CTkEntry(frame_busqueda, width=180)
        self.entry_devolucion_codigo.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar prenda",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_prenda_para_devolucion
        )
        boton_buscar.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_estado_devolucion = ctk.CTkLabel(
            self.frame_devolucion,
            text="Buscá una prenda por código para registrar la devolución.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_devolucion.pack(fill="x", padx=20, pady=(0, 10))

        self.label_detalle_devolucion = ctk.CTkLabel(
            self.frame_devolucion,
            text="Todavía no hay una prenda cargada para devolución.",
            text_color="#666666",
            font=FUENTE_SUBTITULO,
            justify="left",
            anchor="w"
        )
        self.label_detalle_devolucion.pack(fill="x", padx=20, pady=(0, 20))

        boton_devolver = ctk.CTkButton(
            self.frame_devolucion,
            text="Registrar devolución",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.registrar_devolucion_desde_pantalla
        )
        boton_devolver.pack(anchor="e", padx=20, pady=(0, 20))

    def abrir_ventana_aprobacion_remarque(self):
        self._ocultar_frames_secundarios()
        self.frame_devolucion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_devolucion.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_devolucion, "APROBACIÓN DE REMARQUE", self.abrir_modulo_devolucion)

        frame_busqueda = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_busqueda.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_busqueda, text="Código proveedora", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_aprobacion_remarque_proveedora = ctk.CTkEntry(frame_busqueda, width=180)
        self.entry_aprobacion_remarque_proveedora.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar pendientes",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_lote_aprobacion_remarque
        )
        boton_buscar.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_estado_aprobacion_remarque = ctk.CTkLabel(
            self.frame_devolucion,
            text="Buscá una proveedora para cargar sus remarques pendientes.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_aprobacion_remarque.pack(fill="x", padx=20, pady=(0, 10))

        frame_resumen = ctk.CTkFrame(self.frame_devolucion, fg_color=COLOR_FONDO)
        frame_resumen.pack(fill="x", padx=20, pady=10)

        self.aprobacion_remarque_vars = {
            "nombre_proveedora": ctk.StringVar(value="-"),
            "cantidad_prendas": ctk.StringVar(value="-"),
        }

        ctk.CTkLabel(frame_resumen, text="Proveedora", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=(6, 2), sticky="w"
        )
        ctk.CTkLabel(
            frame_resumen,
            textvariable=self.aprobacion_remarque_vars["nombre_proveedora"],
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=1, column=0, padx=10, pady=(0, 8), sticky="w")

        ctk.CTkLabel(frame_resumen, text="Pendientes", text_color=COLOR_TEXTO).grid(
            row=0, column=1, padx=10, pady=(6, 2), sticky="w"
        )
        ctk.CTkLabel(
            frame_resumen,
            textvariable=self.aprobacion_remarque_vars["cantidad_prendas"],
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=1, column=1, padx=10, pady=(0, 8), sticky="w")

        self.frame_tabla_aprobacion_remarque = ctk.CTkScrollableFrame(
            self.frame_devolucion,
            fg_color=COLOR_FONDO,
            height=320
        )
        self.frame_tabla_aprobacion_remarque.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        encabezados = [
            ("Código", 0, 90),
            ("Artículo", 1, 130),
            ("Precio actual", 2, 100),
            ("Precio remarcado", 3, 120),
            ("Fecha remarque", 4, 100),
            ("Decisión", 5, 120),
        ]
        for texto, columna, ancho in encabezados:
            ctk.CTkLabel(
                self.frame_tabla_aprobacion_remarque,
                text=texto,
                text_color=COLOR_TEXTO,
                width=ancho
            ).grid(row=0, column=columna, padx=4, pady=(0, 6), sticky="w")

        self.filas_aprobacion_remarque = []
        self.fila_actual_aprobacion_remarque = 1
        self.lote_aprobacion_remarque_actual = None
        self.decision_remarque_opciones = ["PENDIENTE", "APROBADO", "DEVOLVER"]

        boton_guardar = ctk.CTkButton(
            self.frame_devolucion,
            text="Guardar decisiones",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.guardar_decisiones_remarque
        )
        boton_guardar.pack(anchor="e", padx=20, pady=(0, 20))

    def abrir_editar_venta(self):
        self._ocultar_frames_secundarios()
        self.frame_editar.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_editar.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_editar, "EDITAR VENTA", self.abrir_ventana_editar)

        frame_busqueda = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_busqueda.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_busqueda, text="Código prenda", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_editar_venta_codigo = ctk.CTkEntry(frame_busqueda, width=160)
        self.entry_editar_venta_codigo.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar venta",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_venta_para_editar
        )
        boton_buscar.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_estado_editar_venta = ctk.CTkLabel(
            self.frame_editar,
            text="Buscá una venta por código de prenda.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_editar_venta.pack(fill="x", padx=20, pady=(0, 10))

        frame_formulario = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_formulario.pack(fill="x", padx=20, pady=10)

        campos = [
            ("Fecha venta", "fecha_venta", 0, 0),
            ("Precio venta", "precio_venta", 0, 1),
            ("Cliente", "cliente", 2, 0),
            ("Obs venta", "obs_venta", 2, 1),
        ]

        self.editar_venta_entries = {}
        for texto, clave, fila, columna in campos:
            ctk.CTkLabel(frame_formulario, text=texto, text_color=COLOR_TEXTO).grid(
                row=fila, column=columna, padx=10, pady=(6, 2), sticky="w"
            )
            entry = ctk.CTkEntry(frame_formulario, width=220)
            entry.grid(row=fila + 1, column=columna, padx=10, pady=(0, 8), sticky="w")
            self.editar_venta_entries[clave] = entry
            if clave == "fecha_venta":
                self._aplicar_mascara_fecha(entry)

        ctk.CTkLabel(frame_formulario, text="Tipo pago", text_color=COLOR_TEXTO).grid(
            row=4, column=0, padx=10, pady=(6, 2), sticky="w"
        )
        self.editar_venta_tipo_pago = ctk.CTkOptionMenu(
            frame_formulario,
            values=["TRANSFERENCIA", "EFECTIVO", "DESCUENTO A PROVEEDORA", "OTRO"],
            width=220,
            fg_color=COLOR_FONDO,
            button_color=COLOR_FONDO,
            button_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            dropdown_fg_color=COLOR_FONDO,
            dropdown_text_color=COLOR_TEXTO
        )
        self.editar_venta_tipo_pago.grid(row=5, column=0, padx=10, pady=(0, 8), sticky="w")

        ctk.CTkLabel(frame_formulario, text="Validación", text_color=COLOR_TEXTO).grid(
            row=4, column=1, padx=10, pady=(6, 2), sticky="w"
        )
        self.editar_venta_validacion = ctk.CTkOptionMenu(
            frame_formulario,
            values=["PENDIENTE", "PAGADO"],
            width=220,
            fg_color=COLOR_FONDO,
            button_color=COLOR_FONDO,
            button_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            dropdown_fg_color=COLOR_FONDO,
            dropdown_text_color=COLOR_TEXTO
        )
        self.editar_venta_validacion.grid(row=5, column=1, padx=10, pady=(0, 8), sticky="w")

        self.label_editar_venta_info = ctk.CTkLabel(
            self.frame_editar,
            text="Todavía no hay una venta cargada para editar.",
            text_color="#666666",
            font=FUENTE_SUBTITULO
        )
        self.label_editar_venta_info.pack(fill="x", padx=20, pady=(0, 10))

        boton_guardar = ctk.CTkButton(
            self.frame_editar,
            text="Guardar cambios",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.guardar_edicion_venta
        )
        boton_guardar.pack(anchor="e", padx=20, pady=(0, 20))

    def abrir_editar_prenda(self):
        self._ocultar_frames_secundarios()
        self.frame_editar.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_editar.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_editar, "EDITAR PRENDA", self.abrir_ventana_editar)

        frame_busqueda = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_busqueda.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_busqueda, text="Código prenda", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_editar_prenda_codigo = ctk.CTkEntry(frame_busqueda, width=160)
        self.entry_editar_prenda_codigo.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar prenda",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_prenda_para_editar
        )
        boton_buscar.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.label_estado_editar_prenda = ctk.CTkLabel(
            self.frame_editar,
            text="Buscá una prenda por código.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_editar_prenda.pack(fill="x", padx=20, pady=(0, 10))

        frame_formulario = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_formulario.pack(fill="x", padx=20, pady=10)

        campos = [
            ("Artículo", "articulo", 0, 0),
            ("Marca", "marca", 0, 1),
            ("Talle", "talle", 2, 0),
            ("Color", "color", 2, 1),
            ("Precio", "precio", 4, 0),
            ("Obs ingreso", "obs_ingreso", 4, 1),
        ]

        self.editar_prenda_entries = {}
        for texto, clave, fila, columna in campos:
            ctk.CTkLabel(frame_formulario, text=texto, text_color=COLOR_TEXTO).grid(
                row=fila, column=columna, padx=10, pady=(6, 2), sticky="w"
            )
            entry = ctk.CTkEntry(frame_formulario, width=220)
            entry.grid(row=fila + 1, column=columna, padx=10, pady=(0, 8), sticky="w")
            self.editar_prenda_entries[clave] = entry

        self.label_editar_prenda_info = ctk.CTkLabel(
            self.frame_editar,
            text="Todavía no hay una prenda cargada para editar.",
            text_color="#666666",
            font=FUENTE_SUBTITULO
        )
        self.label_editar_prenda_info.pack(fill="x", padx=20, pady=(0, 10))

        boton_guardar = ctk.CTkButton(
            self.frame_editar,
            text="Guardar cambios",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.guardar_edicion_prenda
        )
        boton_guardar.pack(anchor="e", padx=20, pady=(0, 20))

    def abrir_eliminar_prendas(self):
        self._ocultar_frames_secundarios()
        self.frame_editar.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_editar.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_editar, "ELIMINAR PRENDAS", self.abrir_ventana_editar)

        frame_contenido = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_contenido.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            frame_contenido,
            text="Pegá uno o varios códigos de prenda, uno por línea.",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        ).pack(anchor="w", pady=(0, 10))

        self.label_estado_eliminar_prendas = ctk.CTkLabel(
            frame_contenido,
            text="Todavía no cargaste códigos para eliminar.",
            text_color="#666666",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_eliminar_prendas.pack(fill="x", pady=(0, 10))

        self.textbox_eliminar_prendas = ctk.CTkTextbox(
            frame_contenido,
            height=220,
            font=("Consolas", 14),
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color="#D9D9D9"
        )
        self.textbox_eliminar_prendas.pack(fill="both", expand=True, pady=(0, 10))

        self.textbox_resultado_eliminar_prendas = ctk.CTkTextbox(
            frame_contenido,
            height=180,
            font=("Consolas", 13),
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color="#D9D9D9"
        )
        self.textbox_resultado_eliminar_prendas.pack(fill="both", expand=True, pady=(0, 10))
        self.textbox_resultado_eliminar_prendas.insert("1.0", "Todavía no hay resultados para mostrar.")
        self.textbox_resultado_eliminar_prendas.configure(state="disabled")

        boton_confirmar = ctk.CTkButton(
            frame_contenido,
            text="Eliminar prendas",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.eliminar_prendas_desde_pantalla
        )
        boton_confirmar.pack(anchor="e")

    def abrir_reversar_ventas(self):
        self._ocultar_frames_secundarios()
        self.frame_editar.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_editar.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_editar, "REVERSAR VENTAS", self.abrir_ventana_editar)

        frame_contenido = ctk.CTkFrame(self.frame_editar, fg_color=COLOR_FONDO)
        frame_contenido.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            frame_contenido,
            text="Pegá uno o varios códigos de prenda vendidos, uno por línea.",
            font=FUENTE_SUBTITULO,
            text_color=COLOR_TEXTO
        ).pack(anchor="w", pady=(0, 10))

        self.label_estado_reversar_ventas = ctk.CTkLabel(
            frame_contenido,
            text="Todavía no cargaste códigos para reversar.",
            text_color="#666666",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_reversar_ventas.pack(fill="x", pady=(0, 10))

        self.textbox_reversar_ventas = ctk.CTkTextbox(
            frame_contenido,
            height=220,
            font=("Consolas", 14),
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color="#D9D9D9"
        )
        self.textbox_reversar_ventas.pack(fill="both", expand=True, pady=(0, 10))

        self.textbox_resultado_reversar_ventas = ctk.CTkTextbox(
            frame_contenido,
            height=180,
            font=("Consolas", 13),
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color="#D9D9D9"
        )
        self.textbox_resultado_reversar_ventas.pack(fill="both", expand=True, pady=(0, 10))
        self.textbox_resultado_reversar_ventas.insert("1.0", "Todavía no hay resultados para mostrar.")
        self.textbox_resultado_reversar_ventas.configure(state="disabled")

        boton_confirmar = ctk.CTkButton(
            frame_contenido,
            text="Reversar ventas",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.reversar_ventas_desde_pantalla
        )
        boton_confirmar.pack(anchor="e")

    def abrir_ventana_resumen_general(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_resumen_general.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_resumen_general.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_resumen_general, "RESUMEN GENERAL", self.abrir_ventana_consultas)

        hoy = datetime.now()

        frame_filtros = ctk.CTkFrame(self.frame_resumen_general, fg_color=COLOR_FONDO)
        frame_filtros.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_filtros, text="Mes", text_color=COLOR_TEXTO).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_resumen_mes = ctk.CTkEntry(frame_filtros, width=80)
        self.entry_resumen_mes.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_resumen_mes.insert(0, f"{hoy.month:02d}")

        ctk.CTkLabel(frame_filtros, text="Año", text_color=COLOR_TEXTO).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        self.entry_resumen_anio = ctk.CTkEntry(frame_filtros, width=100)
        self.entry_resumen_anio.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        self.entry_resumen_anio.insert(0, str(hoy.year))

        boton_buscar = ctk.CTkButton(
            frame_filtros,
            text="Buscar resumen",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_resumen_general
        )
        boton_buscar.grid(row=1, column=2, padx=10, pady=5, sticky="w")

        self.label_resumen_estado = ctk.CTkLabel(
            self.frame_resumen_general,
            text="Ingresá mes y año para generar el resumen general.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_resumen_estado.pack(fill="x", padx=20, pady=(0, 10))

        frame_resumen = ctk.CTkFrame(self.frame_resumen_general, fg_color=COLOR_FONDO)
        frame_resumen.pack(fill="x", padx=20, pady=(0, 10))

        self.resumen_general_vars = {
            "cantidad_total_vendida": ctk.StringVar(value="-"),
            "total_vendido": ctk.StringVar(value="-"),
            "total_proveedoras": ctk.StringVar(value="-"),
            "total_fashion_reset": ctk.StringVar(value="-"),
            "total_descuentos": ctk.StringVar(value="-"),
            "total_neto_a_pagar": ctk.StringVar(value="-"),
        }

        labels_resumen = [
            ("Cantidad total vendida", "cantidad_total_vendida"),
            ("Total vendido", "total_vendido"),
            ("Total proveedoras (60%)", "total_proveedoras"),
            ("Total Fashion Reset (40%)", "total_fashion_reset"),
            ("Total descuentos", "total_descuentos"),
            ("Total neto a pagar", "total_neto_a_pagar"),
        ]

        for indice, (texto, clave) in enumerate(labels_resumen):
            fila = indice // 2
            columna = (indice % 2) * 2

            ctk.CTkLabel(
                frame_resumen,
                text=f"{texto}:",
                text_color=COLOR_TEXTO,
                font=FUENTE_SUBTITULO
            ).grid(row=fila, column=columna, padx=(0, 8), pady=6, sticky="w")

            ctk.CTkLabel(
                frame_resumen,
                textvariable=self.resumen_general_vars[clave],
                text_color=COLOR_TEXTO,
                font=FUENTE_SUBTITULO
            ).grid(row=fila, column=columna + 1, padx=(0, 20), pady=6, sticky="w")

        label_detalle = ctk.CTkLabel(
            self.frame_resumen_general,
            text="Desglose por proveedora",
            text_color=COLOR_TEXTO,
            font=FUENTE_FORMULARIO_TITULO
        )
        label_detalle.pack(anchor="w", padx=20, pady=(10, 5))

        self.textbox_resumen_general = ctk.CTkTextbox(
            self.frame_resumen_general,
            height=260,
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO
        )
        self.textbox_resumen_general.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.textbox_resumen_general.configure(font=("Consolas", 13))
        self.textbox_resumen_general.insert("1.0", "Todavía no hay resultados para mostrar.")
        self.textbox_resumen_general.configure(state="disabled")

    def abrir_ventana_pendientes_validacion(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_pendientes_validacion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_pendientes_validacion.winfo_children():
            widget.destroy()

        self._crear_encabezado(
            self.frame_pendientes_validacion,
            "VENTAS PENDIENTES DE VALIDACIÓN",
            self.abrir_ventana_consultas
        )

        frame_acciones = ctk.CTkFrame(self.frame_pendientes_validacion, fg_color=COLOR_FONDO)
        frame_acciones.pack(fill="x", padx=20, pady=10)

        boton_buscar = ctk.CTkButton(
            frame_acciones,
            text="Buscar pendientes",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_ventas_pendientes
        )
        boton_buscar.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            frame_acciones,
            text="Código prenda",
            text_color=COLOR_TEXTO
        ).pack(side="left", padx=(20, 8))

        self.entry_validar_pendiente = ctk.CTkEntry(frame_acciones, width=120)
        self.entry_validar_pendiente.pack(side="left", padx=(0, 10))

        boton_validar = ctk.CTkButton(
            frame_acciones,
            text="Validar venta",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.validar_venta_pendiente_desde_pantalla
        )
        boton_validar.pack(side="left", padx=(0, 10))

        self.label_pendientes_estado = ctk.CTkLabel(
            frame_acciones,
            text="Mostrá las ventas pendientes de validación.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_pendientes_estado.pack(side="left", padx=10)

        frame_resumen = ctk.CTkFrame(self.frame_pendientes_validacion, fg_color=COLOR_FONDO)
        frame_resumen.pack(fill="x", padx=20, pady=(0, 10))

        self.pendientes_vars = {
            "cantidad_pendientes": ctk.StringVar(value="-"),
            "total_importe_pendiente": ctk.StringVar(value="-"),
        }

        ctk.CTkLabel(
            frame_resumen,
            text="Cantidad pendientes:",
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=0, column=0, padx=(0, 8), pady=6, sticky="w")
        ctk.CTkLabel(
            frame_resumen,
            textvariable=self.pendientes_vars["cantidad_pendientes"],
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=0, column=1, padx=(0, 20), pady=6, sticky="w")

        ctk.CTkLabel(
            frame_resumen,
            text="Total importe pendiente:",
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=0, column=2, padx=(0, 8), pady=6, sticky="w")
        ctk.CTkLabel(
            frame_resumen,
            textvariable=self.pendientes_vars["total_importe_pendiente"],
            text_color=COLOR_TEXTO,
            font=FUENTE_SUBTITULO
        ).grid(row=0, column=3, padx=(0, 20), pady=6, sticky="w")

        label_detalle = ctk.CTkLabel(
            self.frame_pendientes_validacion,
            text="Detalle de ventas pendientes",
            text_color=COLOR_TEXTO,
            font=FUENTE_FORMULARIO_TITULO
        )
        label_detalle.pack(anchor="w", padx=20, pady=(10, 5))

        self.textbox_pendientes = ctk.CTkTextbox(
            self.frame_pendientes_validacion,
            height=260,
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO
        )
        self.textbox_pendientes.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.textbox_pendientes.configure(font=("Consolas", 13))
        self.textbox_pendientes.insert("1.0", "Todavía no hay resultados para mostrar.")
        self.textbox_pendientes.configure(state="disabled")

    def abrir_formulario_proveedora(self):
        self._ocultar_frames_secundarios()
        self.frame_proveedoras.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_proveedoras.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_proveedoras, "NUEVA PROVEEDORA", self.abrir_ventana_proveedoras)

        frame_formulario = ctk.CTkFrame(self.frame_proveedoras, fg_color=COLOR_FONDO)
        frame_formulario.pack(fill="x", padx=20, pady=10)

        campos = [
            ("Código proveedora", "codigo", 0, 0),
            ("Nombre proveedora", "nombre", 0, 1),
            ("Teléfono", "telefono", 2, 0),
            ("Banco", "banco", 2, 1),
            ("Número cuenta", "numero_cuenta", 4, 0),
            ("Titular cuenta", "titular_cuenta", 4, 1),
            ("Alias", "alias", 6, 0),
            ("Observación", "obs", 6, 1),
        ]

        self.proveedora_entries = {}
        for texto, clave, fila, columna in campos:
            ctk.CTkLabel(frame_formulario, text=texto, text_color=COLOR_TEXTO).grid(
                row=fila, column=columna, padx=10, pady=(6, 2), sticky="w"
            )
            entry = ctk.CTkEntry(frame_formulario, width=260)
            entry.grid(row=fila + 1, column=columna, padx=10, pady=(0, 8), sticky="w")
            self.proveedora_entries[clave] = entry

        boton_guardar = ctk.CTkButton(
            self.frame_proveedoras,
            text="Guardar proveedora",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.guardar_proveedora_desde_pantalla
        )
        boton_guardar.pack(anchor="e", padx=20, pady=(10, 10))

        self.label_estado_proveedora = ctk.CTkLabel(
            self.frame_proveedoras,
            text="Completá los datos para registrar una nueva proveedora.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_proveedora.pack(fill="x", padx=20, pady=(0, 20))

    def abrir_consulta_proveedora(self):
        self._ocultar_frames_secundarios()
        self.frame_proveedoras.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_proveedoras.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_proveedoras, "VER PROVEEDORA", self.abrir_ventana_proveedoras)

        frame_busqueda = ctk.CTkFrame(self.frame_proveedoras, fg_color=COLOR_FONDO)
        frame_busqueda.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_busqueda, text="Código proveedora", text_color=COLOR_TEXTO).grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_buscar_proveedora = ctk.CTkEntry(frame_busqueda, width=160)
        self.entry_buscar_proveedora.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_busqueda,
            text="Buscar proveedora",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_proveedora_desde_pantalla
        )
        boton_buscar.grid(row=1, column=1, padx=10, pady=5, sticky="w")


        boton_todas = ctk.CTkButton(
            frame_busqueda,
            text="Todas",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.mostrar_todas_las_proveedoras
        )
        boton_todas.grid(row=1, column=2, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(frame_busqueda, text="Filtrar prendas", text_color=COLOR_TEXTO).grid(
            row=0, column=3, padx=10, pady=5, sticky="w"
        )
        self.filtro_estado_proveedora = ctk.StringVar(value="TODAS")
        self.option_filtro_estado_proveedora = ctk.CTkOptionMenu(
            frame_busqueda,
            values=["TODAS", "DISPONIBLES", "VENDIDAS"],
            variable=self.filtro_estado_proveedora,
            width=140,
            fg_color=COLOR_FONDO,
            button_color=COLOR_FONDO,
            button_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            dropdown_fg_color=COLOR_FONDO,
            dropdown_text_color=COLOR_TEXTO,
            command=lambda _valor: self.actualizar_filtro_proveedora()
        )
        self.option_filtro_estado_proveedora.grid(row=1, column=3, padx=10, pady=5, sticky="w")

        self.label_estado_ver_proveedora = ctk.CTkLabel(
            self.frame_proveedoras,
            text="Ingresá un código para consultar la proveedora.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_estado_ver_proveedora.pack(fill="x", padx=20, pady=(0, 10))

        self.datos_proveedora_vars = {
            "nombre_proveedora": ctk.StringVar(value="-"),
            "telefono": ctk.StringVar(value="-"),
            "banco": ctk.StringVar(value="-"),
            "numero_cuenta": ctk.StringVar(value="-"),
            "titular_cuenta": ctk.StringVar(value="-"),
            "alias": ctk.StringVar(value="-"),
            "obs_ingreso": ctk.StringVar(value="-"),
            "estado": ctk.StringVar(value="-"),
            "total_prendas": ctk.StringVar(value="-"),
            "prendas_disponibles": ctk.StringVar(value="-"),
            "prendas_vendidas": ctk.StringVar(value="-"),
            "prendas_devueltas": ctk.StringVar(value="-"),
        }

        frame_datos = ctk.CTkFrame(self.frame_proveedoras, fg_color=COLOR_FONDO)
        frame_datos.pack(fill="x", padx=20, pady=(0, 10))

        labels = [
            ("Nombre", "nombre_proveedora"),
            ("Teléfono", "telefono"),
            ("Banco", "banco"),
            ("Número cuenta", "numero_cuenta"),
            ("Titular cuenta", "titular_cuenta"),
            ("Alias", "alias"),
            ("Observación", "obs_ingreso"),
            ("Estado", "estado"),
            ("Total prendas", "total_prendas"),
            ("Disponibles", "prendas_disponibles"),
            ("Vendidas", "prendas_vendidas"),
            ("Devueltas", "prendas_devueltas"),
        ]

        for indice, (texto, clave) in enumerate(labels):
            fila = indice // 2
            columna = (indice % 2) * 2
            ctk.CTkLabel(frame_datos, text=f"{texto}:", text_color=COLOR_TEXTO, font=FUENTE_SUBTITULO).grid(
                row=fila, column=columna, padx=(0, 8), pady=6, sticky="w"
            )
            ctk.CTkLabel(frame_datos, textvariable=self.datos_proveedora_vars[clave], text_color=COLOR_TEXTO, font=FUENTE_SUBTITULO).grid(
                row=fila, column=columna + 1, padx=(0, 20), pady=6, sticky="w"
            )

        label_detalle = ctk.CTkLabel(
            self.frame_proveedoras,
            text="Detalle de prendas",
            text_color=COLOR_TEXTO,
            font=FUENTE_FORMULARIO_TITULO
        )
        label_detalle.pack(anchor="w", padx=20, pady=(10, 5))

        self.textbox_detalle_proveedora = ctk.CTkTextbox(
            self.frame_proveedoras,
            height=240,
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO
        )
        self.textbox_detalle_proveedora.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.textbox_detalle_proveedora.configure(font=("Consolas", 13))
        self.textbox_detalle_proveedora.insert("1.0", "Todavía no hay resultados para mostrar.")
        self.textbox_detalle_proveedora.configure(state="disabled")
        self.detalle_proveedora_actual = []

    def abrir_ventana_rendicion(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack_forget()
        self.frame_rendicion.pack(pady=(5, 10), padx=30, fill="both", expand=True)

        for widget in self.frame_rendicion.winfo_children():
            widget.destroy()

        self._crear_encabezado(self.frame_rendicion, "RENDICIÓN DE CUENTAS", self.abrir_ventana_consultas)

        hoy = datetime.now()

        frame_filtros = ctk.CTkFrame(self.frame_rendicion, fg_color=COLOR_FONDO)
        frame_filtros.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(frame_filtros, text="Mes", text_color=COLOR_TEXTO).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_rendicion_mes = ctk.CTkEntry(frame_filtros, width=80)
        self.entry_rendicion_mes.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_rendicion_mes.insert(0, f"{hoy.month:02d}")

        ctk.CTkLabel(frame_filtros, text="Año", text_color=COLOR_TEXTO).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        self.entry_rendicion_anio = ctk.CTkEntry(frame_filtros, width=100)
        self.entry_rendicion_anio.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        self.entry_rendicion_anio.insert(0, str(hoy.year))

        ctk.CTkLabel(frame_filtros, text="Código proveedora", text_color=COLOR_TEXTO).grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.entry_rendicion_codigo = ctk.CTkEntry(frame_filtros, width=160)
        self.entry_rendicion_codigo.grid(row=1, column=2, padx=10, pady=5, sticky="w")

        boton_buscar = ctk.CTkButton(
            frame_filtros,
            text="Buscar rendición",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            command=self.buscar_rendicion
        )
        boton_buscar.grid(row=1, column=3, padx=10, pady=5, sticky="w")

        self.boton_exportar_rendicion = ctk.CTkButton(
            frame_filtros,
            text="Exportar Excel",
            font=FUENTE_BOTON,
            fg_color=COLOR_FONDO,
            hover_color=COLOR_HOVER,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO,
            state="disabled",
            command=self.exportar_rendicion_actual
        )
        self.boton_exportar_rendicion.grid(row=1, column=4, padx=10, pady=5, sticky="w")

        self.label_rendicion_estado = ctk.CTkLabel(
            self.frame_rendicion,
            text="Ingresá mes, año y código de proveedora para generar la rendición.",
            text_color="#555555",
            font=FUENTE_SUBTITULO
        )
        self.label_rendicion_estado.pack(fill="x", padx=20, pady=(0, 10))

        frame_resumen = ctk.CTkFrame(self.frame_rendicion, fg_color=COLOR_FONDO)
        frame_resumen.pack(fill="x", padx=20, pady=(0, 10))

        self.resumen_vars = {
            "nombre_proveedora": ctk.StringVar(value="-"),
            "cantidad_prendas": ctk.StringVar(value="-"),
            "total_vendido": ctk.StringVar(value="-"),
            "comision_proveedora": ctk.StringVar(value="-"),
            "comision_fashion_reset": ctk.StringVar(value="-"),
            "total_descuentos": ctk.StringVar(value="-"),
            "saldo_final": ctk.StringVar(value="-"),
        }

        labels_resumen = [
            ("Nombre proveedora", "nombre_proveedora"),
            ("Cantidad prendas", "cantidad_prendas"),
            ("Total vendido", "total_vendido"),
            ("Comisión proveedora", "comision_proveedora"),
            ("Comisión Fashion Reset", "comision_fashion_reset"),
            ("Descuentos", "total_descuentos"),
            ("Saldo final a pagar", "saldo_final"),
        ]

        for indice, (texto, clave) in enumerate(labels_resumen):
            fila = indice // 2
            columna = (indice % 2) * 2

            ctk.CTkLabel(
                frame_resumen,
                text=f"{texto}:",
                text_color=COLOR_TEXTO,
                font=FUENTE_SUBTITULO
            ).grid(row=fila, column=columna, padx=(0, 8), pady=6, sticky="w")

            ctk.CTkLabel(
                frame_resumen,
                textvariable=self.resumen_vars[clave],
                text_color=COLOR_TEXTO,
                font=FUENTE_SUBTITULO
            ).grid(row=fila, column=columna + 1, padx=(0, 20), pady=6, sticky="w")

        label_detalle = ctk.CTkLabel(
            self.frame_rendicion,
            text="Detalle de ventas",
            text_color=COLOR_TEXTO,
            font=FUENTE_FORMULARIO_TITULO
        )
        label_detalle.pack(anchor="w", padx=20, pady=(10, 5))

        self.textbox_rendicion = ctk.CTkTextbox(
            self.frame_rendicion,
            height=260,
            fg_color=COLOR_FONDO,
            text_color=COLOR_TEXTO,
            border_width=1,
            border_color=COLOR_TEXTO
        )
        self.textbox_rendicion.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.textbox_rendicion.configure(font=("Consolas", 13))
        self.textbox_rendicion.insert("1.0", "Todavía no hay resultados para mostrar.")
        self.textbox_rendicion.configure(state="disabled")

    def guardar_lote_venta(self, entry_fecha, entry_cliente):
        fecha_venta = entry_fecha.get().strip()
        cliente = entry_cliente.get().strip()

        if not fecha_venta or not cliente:
            messagebox.showerror("Guardar lote", "Fecha venta y cliente son obligatorios.")
            return

        guardadas = 0
        ignoradas = 0
        errores = []
        codigos_guardados = []

        for indice, fila in enumerate(self.filas_venta, start=1):
            codigo_prenda = fila[0].get().strip()
            precio_texto = fila[5].get().strip()
            tipo_pago = fila[6].get().strip().upper()
            validacion = fila[7].get().strip().upper()
            obs_venta = fila[8].get().strip()

            if not codigo_prenda:
                if not any([
                    fila[1].get().strip(),
                    fila[2].get().strip(),
                    fila[3].get().strip(),
                    fila[4].get().strip(),
                    precio_texto,
                    fila[8].get().strip()
                ]):
                    ignoradas += 1
                    continue
                ignoradas += 1
                continue

            if not precio_texto:
                resultado_busqueda, datos_prenda = buscar_ingreso_por_codigo(codigo_prenda)
                if resultado_busqueda:
                    precio_texto = str(datos_prenda["precio_lista"] or "").strip()
                    fila[5].configure(state="normal")
                    fila[5].delete(0, "end")
                    fila[5].insert(0, self._format_precio(precio_texto))
                    fila[5].configure(state="disabled")
                else:
                    errores.append(f"Fila {indice}: Falta precio para la prenda {codigo_prenda}.")
                    fila[0].configure(border_color="red")
                    continue

            precio_texto_sin_separadores = precio_texto.replace(".", "").replace(",", "")

            if tipo_pago == "DESCUENTO A PROVEEDORA" and not obs_venta:
                errores.append(
                    f"Fila {indice}: OBS VENTA es obligatorio cuando el tipo de pago es "
                    "DESCUENTO A PROVEEDORA (código de la proveedora)."
                )
                try:
                    fila[8].configure(border_color="red")
                except Exception:
                    pass
                continue

            obs_venta = normalizar_obs_descuento_proveedora(tipo_pago, obs_venta)

            resultado, mensaje = guardar_venta_desde_gui(
                fecha_venta,
                cliente,
                tipo_pago,
                validacion,
                codigo_prenda,
                fila[1].get().strip(),
                fila[2].get().strip(),
                fila[3].get().strip(),
                fila[4].get().strip(),
                precio_texto_sin_separadores,
                precio_texto_sin_separadores,
                obs_venta
            )

            if resultado:
                guardadas += 1
                codigos_guardados.append(codigo_prenda)
                fila[0].configure(border_color=COLOR_TEXTO, state="normal")
                fila[0].delete(0, "end")
                fila[1].delete(0, "end")
                fila[1].configure(state="disabled")
                fila[2].delete(0, "end")
                fila[2].configure(state="disabled")
                fila[3].delete(0, "end")
                fila[3].configure(state="disabled")
                fila[4].delete(0, "end")
                fila[4].configure(state="disabled")
                fila[5].delete(0, "end")
                fila[5].configure(state="disabled")
                fila[6].set("EFECTIVO")
                fila[8].delete(0, "end")
                self._actualizar_tipo_pago_venta("EFECTIVO", fila[8], fila[7])
            else:
                errores.append(f"Fila {indice}: {mensaje}")
                fila[0].configure(border_color="red")

        resumen = [
            f"Guardadas: {guardadas}",
            f"Ignoradas: {ignoradas}",
            f"Errores: {len(errores)}"
        ]
        if codigos_guardados:
            resumen.append("")
            resumen.append("Códigos guardados:")
            resumen.extend(codigos_guardados)

        self._actualizar_total_venta()

        if errores:
            resumen.append("")
            resumen.extend(errores)
            messagebox.showwarning("Resumen de guardado", "\n".join(resumen))
        else:
            messagebox.showinfo("Resumen de guardado", "\n".join(resumen))
            if guardadas:
                self.abrir_ventana_venta()

    def guardar_lote_ingreso(self, entry_fecha, entry_proveedora):
        fecha_ingreso = entry_fecha.get().strip()
        codigo_proveedora = entry_proveedora.get().strip()

        if not fecha_ingreso or not codigo_proveedora:
            messagebox.showerror("Guardar lote", "Fecha ingreso y código proveedora son obligatorios.")
            return

        guardadas = 0
        ignoradas = 0
        errores = []
        codigos_guardados = []

        for indice, fila in enumerate(self.filas_ingreso, start=1):
            numero_prenda = fila[0].get().strip()
            articulo = fila[1].get().strip()
            marca = fila[2].get().strip()
            talle = fila[3].get().strip()
            color = fila[4].get().strip()
            precio_texto = fila[5].get().strip()
            obs_ingreso = fila[6].get().strip()

            if not any([numero_prenda, articulo, marca, talle, color, precio_texto, obs_ingreso]):
                ignoradas += 1
                continue

            if not numero_prenda or not articulo:
                errores.append(f"Fila {indice}: N° prenda y artículo son obligatorios.")
                continue

            resultado, mensaje = guardar_ingreso_desde_gui(
                fecha_ingreso,
                codigo_proveedora,
                numero_prenda,
                articulo,
                marca,
                talle,
                color,
                precio_texto,
                obs_ingreso
            )

            if resultado:
                guardadas += 1
                codigos_guardados.append(str(mensaje))
                for entrada in fila:
                    entrada.delete(0, "end")
            else:
                errores.append(f"Fila {indice}: {mensaje}")

        resumen = [
            f"Guardadas: {guardadas}",
            f"Ignoradas: {ignoradas}",
            f"Errores: {len(errores)}"
        ]
        if codigos_guardados:
            resumen.append("")
            resumen.append("Códigos guardados:")
            resumen.extend(codigos_guardados)

        if errores:
            resumen.append("")
            resumen.extend(errores)
            messagebox.showwarning("Resumen de guardado", "\n".join(resumen))
        else:
            messagebox.showinfo("Resumen de guardado", "\n".join(resumen))

    def buscar_resumen_general(self):
        resultado, datos = calcular_resumen_general(
            self.entry_resumen_mes.get().strip(),
            self.entry_resumen_anio.get().strip()
        )

        if not resultado:
            self.label_resumen_estado.configure(text=datos, text_color="#B00020")
            self._limpiar_resumen_general()
            self._mostrar_desglose_resumen_general("No se encontraron resultados para mostrar.")
            return

        self.label_resumen_estado.configure(
            text=f"Resumen generado para {datos['mes']:02d}/{datos['anio']}.",
            text_color="#2E5E2E"
        )
        self.resumen_general_vars["cantidad_total_vendida"].set(str(datos["cantidad_total_vendida"]))
        self.resumen_general_vars["total_vendido"].set(self._format_moneda(datos["total_vendido"]))
        self.resumen_general_vars["total_proveedoras"].set(self._format_moneda(datos["total_proveedoras"]))
        self.resumen_general_vars["total_fashion_reset"].set(self._format_moneda(datos["total_fashion_reset"]))
        self.resumen_general_vars["total_descuentos"].set(self._format_moneda(datos["total_descuentos"]))
        self.resumen_general_vars["total_neto_a_pagar"].set(self._format_moneda(datos["total_neto_a_pagar"]))
        self._mostrar_desglose_resumen_general(
            self._construir_tabla_resumen_general(datos["detalle_proveedoras"])
        )

    def buscar_ventas_pendientes(self):
        resultado, datos = calcular_ventas_pendientes()

        if not resultado:
            self.label_pendientes_estado.configure(text=datos, text_color="#B00020")
            self._limpiar_pendientes_validacion()
            self._mostrar_ventas_pendientes("No se encontraron resultados para mostrar.")
            return

        self.label_pendientes_estado.configure(
            text="Ventas pendientes cargadas correctamente.",
            text_color="#2E5E2E"
        )
        self.pendientes_vars["cantidad_pendientes"].set(str(datos["cantidad_pendientes"]))
        self.pendientes_vars["total_importe_pendiente"].set(
            self._format_moneda(datos["total_importe_pendiente"])
        )
        self._mostrar_ventas_pendientes(
            self._construir_tabla_pendientes(datos["ventas_pendientes"])
        )

    def guardar_proveedora_desde_pantalla(self):
        resultado, datos = crear_proveedora_desde_gui(
            self.proveedora_entries["codigo"].get().strip(),
            self.proveedora_entries["nombre"].get().strip(),
            self.proveedora_entries["telefono"].get().strip(),
            self.proveedora_entries["banco"].get().strip(),
            self.proveedora_entries["numero_cuenta"].get().strip(),
            self.proveedora_entries["titular_cuenta"].get().strip(),
            self.proveedora_entries["alias"].get().strip(),
            self.proveedora_entries["obs"].get().strip(),
        )

        if not resultado:
            self.label_estado_proveedora.configure(text=datos, text_color="#B00020")
            return

        for entry in self.proveedora_entries.values():
            entry.delete(0, "end")

        self.label_estado_proveedora.configure(
            text=f"Proveedora {datos['codigo_proveedora']} guardada correctamente.",
            text_color="#2E5E2E"
        )

    def buscar_venta_para_editar(self):
        codigo_prenda = self.entry_editar_venta_codigo.get().strip()
        resultado, datos = obtener_venta_desde_gui(codigo_prenda)

        if not resultado:
            self.label_estado_editar_venta.configure(text=datos, text_color="#B00020")
            self._limpiar_formulario_editar_venta()
            return

        self.editar_venta_entries["fecha_venta"].delete(0, "end")
        self.editar_venta_entries["fecha_venta"].insert(0, datos["fecha_venta"])
        self.editar_venta_entries["precio_venta"].delete(0, "end")
        self.editar_venta_entries["precio_venta"].insert(0, self._format_precio(datos["precio_venta"]))
        self.editar_venta_entries["cliente"].delete(0, "end")
        self.editar_venta_entries["cliente"].insert(0, datos["cliente"])
        self.editar_venta_entries["obs_venta"].delete(0, "end")
        self.editar_venta_entries["obs_venta"].insert(0, datos["obs_venta"])
        self.editar_venta_tipo_pago.set(datos["tipo_pago"] or "EFECTIVO")
        self.editar_venta_validacion.set(datos["validacion"] or "PENDIENTE")
        self.label_estado_editar_venta.configure(
            text=f"Venta {datos['codigo_prenda']} cargada para edición.",
            text_color="#2E5E2E"
        )
        self.label_editar_venta_info.configure(
            text=(
                f"Proveedora: {datos['codigo_proveedora']} | "
                f"Artículo: {datos['articulo']} | Marca: {datos['marca']} | "
                f"Talle: {datos['talle']} | Color: {datos['color']}"
            ),
            text_color=COLOR_TEXTO
        )

    def guardar_edicion_venta(self):
        codigo_prenda = self.entry_editar_venta_codigo.get().strip()
        resultado, mensaje = actualizar_venta_desde_gui(
            codigo_prenda,
            self.editar_venta_entries["fecha_venta"].get().strip(),
            self.editar_venta_entries["cliente"].get().strip(),
            self.editar_venta_tipo_pago.get().strip(),
            self.editar_venta_validacion.get().strip(),
            self.editar_venta_entries["precio_venta"].get().strip(),
            self.editar_venta_entries["obs_venta"].get().strip(),
        )

        if not resultado:
            self.label_estado_editar_venta.configure(text=mensaje, text_color="#B00020")
            return

        self.label_estado_editar_venta.configure(text=mensaje, text_color="#2E5E2E")
        self.buscar_venta_para_editar()

        resumen_actualizado = "\n".join([
            f"Código prenda: {codigo_prenda}",
            f"Fecha venta: {self.editar_venta_entries['fecha_venta'].get().strip()}",
            f"Precio venta: {self.editar_venta_entries['precio_venta'].get().strip()}",
            f"Cliente: {self.editar_venta_entries['cliente'].get().strip()}",
            f"Tipo pago: {self.editar_venta_tipo_pago.get().strip()}",
            f"Validación: {self.editar_venta_validacion.get().strip()}",
            f"Obs venta: {self.editar_venta_entries['obs_venta'].get().strip() or '-'}",
        ])
        messagebox.showinfo("Venta editada", f"{mensaje}\n\n{resumen_actualizado}")

    def buscar_prenda_para_editar(self):
        codigo_prenda = self.entry_editar_prenda_codigo.get().strip()
        resultado, datos = obtener_prenda_desde_gui(codigo_prenda)

        if not resultado:
            self.label_estado_editar_prenda.configure(text=datos, text_color="#B00020")
            self._limpiar_formulario_editar_prenda()
            return

        self.editar_prenda_entries["articulo"].delete(0, "end")
        self.editar_prenda_entries["articulo"].insert(0, datos["articulo"])
        self.editar_prenda_entries["marca"].delete(0, "end")
        self.editar_prenda_entries["marca"].insert(0, datos["marca"])
        self.editar_prenda_entries["talle"].delete(0, "end")
        self.editar_prenda_entries["talle"].insert(0, datos["talle"])
        self.editar_prenda_entries["color"].delete(0, "end")
        self.editar_prenda_entries["color"].insert(0, datos["color"])
        self.editar_prenda_entries["precio"].delete(0, "end")
        self.editar_prenda_entries["precio"].insert(0, self._format_precio(datos["precio"]))
        self.editar_prenda_entries["obs_ingreso"].delete(0, "end")
        self.editar_prenda_entries["obs_ingreso"].insert(0, datos["obs_ingreso"])
        self.label_estado_editar_prenda.configure(
            text=f"Prenda {datos['codigo_prenda']} cargada para edición.",
            text_color="#2E5E2E"
        )
        self.label_editar_prenda_info.configure(
            text=f"Proveedora: {datos['codigo_proveedora']} | Estado: {datos['estado']}",
            text_color=COLOR_TEXTO
        )

    def guardar_edicion_prenda(self):
        codigo_prenda = self.entry_editar_prenda_codigo.get().strip()
        resultado, mensaje = actualizar_prenda_desde_gui(
            codigo_prenda,
            self.editar_prenda_entries["articulo"].get().strip(),
            self.editar_prenda_entries["marca"].get().strip(),
            self.editar_prenda_entries["talle"].get().strip(),
            self.editar_prenda_entries["color"].get().strip(),
            self.editar_prenda_entries["precio"].get().strip(),
            self.editar_prenda_entries["obs_ingreso"].get().strip(),
        )

        if not resultado:
            self.label_estado_editar_prenda.configure(text=mensaje, text_color="#B00020")
            return

        self.label_estado_editar_prenda.configure(text=mensaje, text_color="#2E5E2E")
        self.buscar_prenda_para_editar()

        resumen_actualizado = "\n".join([
            f"Código prenda: {codigo_prenda}",
            f"Artículo: {self.editar_prenda_entries['articulo'].get().strip()}",
            f"Marca: {self.editar_prenda_entries['marca'].get().strip()}",
            f"Talle: {self.editar_prenda_entries['talle'].get().strip()}",
            f"Color: {self.editar_prenda_entries['color'].get().strip()}",
            f"Precio: {self.editar_prenda_entries['precio'].get().strip()}",
            f"Obs ingreso: {self.editar_prenda_entries['obs_ingreso'].get().strip() or '-'}",
        ])
        messagebox.showinfo("Prenda editada", f"{mensaje}\n\n{resumen_actualizado}")

    def eliminar_prendas_desde_pantalla(self):
        codigos_texto = self.textbox_eliminar_prendas.get("1.0", "end").strip()
        if not codigos_texto:
            self.label_estado_eliminar_prendas.configure(
                text="ERROR: Ingresá al menos un código.",
                text_color="#B00020"
            )
            return

        confirmar = messagebox.askyesno(
            "Confirmar eliminación",
            "¿Querés eliminar las prendas ingresadas? Esta acción no se puede deshacer."
        )
        if not confirmar:
            return

        resultado, datos = eliminar_ingresos_desde_gui(codigos_texto)
        if not resultado:
            self.label_estado_eliminar_prendas.configure(text="No se eliminaron prendas.", text_color="#B00020")
            self._mostrar_resultado_eliminar_prendas(datos)
            return

        self.label_estado_eliminar_prendas.configure(
            text=f"Se eliminaron {datos['cantidad_eliminadas']} prendas.",
            text_color="#2E5E2E"
        )
        self._mostrar_resultado_eliminar_prendas(self._construir_resumen_eliminar_prendas(datos))

    def reversar_ventas_desde_pantalla(self):
        codigos_texto = self.textbox_reversar_ventas.get("1.0", "end").strip()
        if not codigos_texto:
            self.label_estado_reversar_ventas.configure(
                text="ERROR: Ingresá al menos un código.",
                text_color="#B00020"
            )
            return

        confirmar = messagebox.askyesno(
            "Confirmar reversa",
            "¿Querés reversar las ventas ingresadas?"
        )
        if not confirmar:
            return

        resultado, datos = reversar_ventas_desde_gui(codigos_texto)
        if not resultado:
            self.label_estado_reversar_ventas.configure(text="No se reversaron ventas.", text_color="#B00020")
            self._mostrar_resultado_reversar_ventas(datos)
            return

        self.label_estado_reversar_ventas.configure(
            text=f"Se reversaron {datos['cantidad_reversadas']} ventas.",
            text_color="#2E5E2E"
        )
        self._mostrar_resultado_reversar_ventas(self._construir_resumen_reversar_ventas(datos))

    def buscar_lote_remarque(self):
        codigo_proveedora = self.entry_remarque_proveedora.get().strip()
        resultado, datos = obtener_lote_remarque_proveedora_desde_gui(codigo_proveedora)

        if not resultado:
            self.label_estado_remarque.configure(text=datos, text_color="#B00020")
            self._limpiar_formulario_remarque()
            return

        self.lote_remarque_actual = datos
        self.remarque_vars["nombre_proveedora"].set(
            datos["nombre_proveedora"] or datos["codigo_proveedora"]
        )
        self.remarque_vars["cantidad_prendas"].set(str(datos["cantidad_prendas"]))
        self.label_estado_remarque.configure(
            text=f"Lote cargado para {datos['codigo_proveedora']}.",
            text_color="#2E5E2E"
        )
        self._cargar_filas_remarque(datos["prendas"])

    def guardar_lote_remarque(self):
        codigo_proveedora = self.entry_remarque_proveedora.get().strip()

        if not self.filas_remarque:
            self.label_estado_remarque.configure(
                text="ERROR: Primero cargá un lote de remarque.",
                text_color="#B00020"
            )
            return

        prendas_remarque = []
        for fila in self.filas_remarque:
            prendas_remarque.append({
                "codigo_prenda": fila["codigo_prenda"],
                "precio_remarcado": fila["entry_precio_remarcado"].get().strip(),
            })

        resultado, datos = guardar_lote_remarque_desde_gui(codigo_proveedora, prendas_remarque)
        if not resultado:
            self.label_estado_remarque.configure(text=datos, text_color="#B00020")
            return

        self.label_estado_remarque.configure(
            text=f"Se guardaron {datos['cantidad_actualizadas']} remarques para {datos['codigo_proveedora']}.",
            text_color="#2E5E2E"
        )

        self.buscar_lote_remarque()

        resumen = "\n".join(
            [
                f"{fila['codigo_prenda']}: {self._format_moneda(fila['precio_remarcado'])}"
                for fila in datos["prendas_actualizadas"][:10]
            ]
        )
        if len(datos["prendas_actualizadas"]) > 10:
            resumen += "\n..."

        messagebox.showinfo(
            "Lote de remarque guardado",
            "\n".join([
                f"Proveedora: {datos['codigo_proveedora']}",
                f"Cantidad actualizada: {datos['cantidad_actualizadas']}",
                f"Fecha lote: {datos['fecha_remarque']}",
                resumen
            ])
        )

    def exportar_lote_remarque_actual(self):
        codigo_proveedora = self.entry_remarque_proveedora.get().strip()

        if not self.filas_remarque or not self.lote_remarque_actual:
            self.label_estado_remarque.configure(
                text="ERROR: Primero cargá un lote de remarque.",
                text_color="#B00020"
            )
            return

        prendas_remarque = []
        for fila in self.filas_remarque:
            prendas_remarque.append({
                "codigo_prenda": fila["codigo_prenda"],
                "articulo": fila["articulo"],
                "marca": fila["marca"],
                "talle": fila["talle"],
                "color": fila["color"],
                "precio_actual": fila["precio_actual"],
                "precio_remarcado": fila["entry_precio_remarcado"].get().strip(),
            })

        resultado, mensaje = exportar_lote_remarque_excel(
            codigo_proveedora,
            self.lote_remarque_actual.get("nombre_proveedora", ""),
            prendas_remarque
        )
        if not resultado:
            self.label_estado_remarque.configure(text=mensaje, text_color="#B00020")
            return

        self.label_estado_remarque.configure(
            text=f"Lote exportado para {codigo_proveedora}.",
            text_color="#2E5E2E"
        )
        messagebox.showinfo("Exportar remarque", f"Archivo generado: {mensaje}")

    def _cargar_filas_remarque(self, prendas):
        for widget in self.frame_tabla_remarque.winfo_children():
            info = widget.grid_info()
            if int(info.get("row", 0)) > 0:
                widget.destroy()

        self.filas_remarque = []
        self.fila_actual_remarque = 1

        for prenda in prendas:
            self._agregar_fila_remarque(prenda)

    def _agregar_fila_remarque(self, prenda):
        fila_grid = self.fila_actual_remarque

        valores = [
            (prenda["codigo_prenda"], 0, 90),
            (prenda["articulo"], 1, 140),
            (prenda["marca"], 2, 100),
            (prenda["talle"], 3, 70),
            (prenda["color"], 4, 90),
            (self._format_moneda(prenda["precio_actual"]), 5, 100),
            (str(prenda["dias_en_inventario"]), 6, 60),
            (
                self._format_moneda(prenda["precio_remarcado"])
                if prenda["precio_remarcado"] not in ("", None)
                else "-",
                7,
                110,
            ),
        ]

        for texto, columna, ancho in valores:
            ctk.CTkLabel(
                self.frame_tabla_remarque,
                text=str(texto),
                text_color=COLOR_TEXTO,
                width=ancho,
                anchor="w"
            ).grid(row=fila_grid, column=columna, padx=4, pady=4, sticky="w")

        entry_precio_remarcado = ctk.CTkEntry(self.frame_tabla_remarque, width=120)
        entry_precio_remarcado.grid(row=fila_grid, column=8, padx=4, pady=4, sticky="w")
        if prenda["precio_remarcado"] not in ("", None):
            entry_precio_remarcado.insert(0, self._format_precio(prenda["precio_remarcado"]))
        entry_precio_remarcado.bind("<FocusOut>", lambda event, e=entry_precio_remarcado: self._formatear_precio_entry_remarque(e))

        self.filas_remarque.append({
            "codigo_prenda": prenda["codigo_prenda"],
            "articulo": prenda["articulo"],
            "marca": prenda["marca"],
            "talle": prenda["talle"],
            "color": prenda["color"],
            "precio_actual": prenda["precio_actual"],
            "entry_precio_remarcado": entry_precio_remarcado,
        })
        self.fila_actual_remarque += 1

    def cargar_prendas_vencidas(self):
        resultado, datos = calcular_prendas_vencidas()

        if not resultado:
            self.label_estado_prendas_vencidas.configure(text=datos, text_color="#B00020")
            if hasattr(self, "prendas_vencidas_vars"):
                self.prendas_vencidas_vars["cantidad_prendas"].set("-")
            self._mostrar_prendas_vencidas("No se encontraron resultados para mostrar.")
            return

        self.label_estado_prendas_vencidas.configure(
            text="Prendas vencidas cargadas correctamente.",
            text_color="#2E5E2E"
        )
        self.prendas_vencidas_vars["cantidad_prendas"].set(str(datos["cantidad_prendas"]))
        self._mostrar_prendas_vencidas(
            self._construir_tabla_prendas_vencidas(datos["prendas_vencidas"])
        )

    def _mostrar_modulo_devolucion_pendiente(self):
        messagebox.showinfo(
            "Devolución",
            "Vamos a construir esta parte del proceso en el próximo paso."
        )

    def buscar_lote_aprobacion_remarque(self):
        codigo_proveedora = self.entry_aprobacion_remarque_proveedora.get().strip()
        resultado, datos = obtener_lote_decision_remarque_desde_gui(codigo_proveedora)

        if not resultado:
            self.label_estado_aprobacion_remarque.configure(text=datos, text_color="#B00020")
            self._limpiar_formulario_aprobacion_remarque()
            return

        self.lote_aprobacion_remarque_actual = datos
        self.aprobacion_remarque_vars["nombre_proveedora"].set(
            datos["nombre_proveedora"] or datos["codigo_proveedora"]
        )
        self.aprobacion_remarque_vars["cantidad_prendas"].set(str(datos["cantidad_prendas"]))
        self.label_estado_aprobacion_remarque.configure(
            text=f"Pendientes cargados para {datos['codigo_proveedora']}.",
            text_color="#2E5E2E"
        )
        self._cargar_filas_aprobacion_remarque(datos["prendas"])

    def guardar_decisiones_remarque(self):
        codigo_proveedora = self.entry_aprobacion_remarque_proveedora.get().strip()

        if not self.filas_aprobacion_remarque:
            self.label_estado_aprobacion_remarque.configure(
                text="ERROR: Primero cargá un lote pendiente.",
                text_color="#B00020"
            )
            return

        decisiones = []
        for fila in self.filas_aprobacion_remarque:
            decisiones.append({
                "codigo_prenda": fila["codigo_prenda"],
                "decision": fila["decision_var"].get().strip(),
            })

        resultado, datos = guardar_decisiones_remarque_desde_gui(codigo_proveedora, decisiones)
        if not resultado:
            self.label_estado_aprobacion_remarque.configure(text=datos, text_color="#B00020")
            return

        self.label_estado_aprobacion_remarque.configure(
            text=(
                f"Aprobadas: {datos['cantidad_aprobadas']} | "
                f"Pendientes devolución: {datos['cantidad_devolucion']}"
            ),
            text_color="#2E5E2E"
        )

        resultado_recarga, datos_recarga = obtener_lote_decision_remarque_desde_gui(codigo_proveedora)
        if resultado_recarga:
            self.lote_aprobacion_remarque_actual = datos_recarga
            self.aprobacion_remarque_vars["nombre_proveedora"].set(
                datos_recarga["nombre_proveedora"] or datos_recarga["codigo_proveedora"]
            )
            self.aprobacion_remarque_vars["cantidad_prendas"].set(str(datos_recarga["cantidad_prendas"]))
            self._cargar_filas_aprobacion_remarque(datos_recarga["prendas"])
        else:
            self._limpiar_formulario_aprobacion_remarque()

        resumen = []
        if datos["prendas_aprobadas"]:
            resumen.append(f"Aprobadas: {datos['cantidad_aprobadas']}")
        if datos["prendas_devolucion"]:
            resumen.append(f"Pendientes devolución: {datos['cantidad_devolucion']}")

        messagebox.showinfo(
            "Decisiones guardadas",
            "\n".join([
                f"Proveedora: {datos['codigo_proveedora']}",
                *resumen,
            ])
        )

    def _cargar_filas_aprobacion_remarque(self, prendas):
        for widget in self.frame_tabla_aprobacion_remarque.winfo_children():
            info = widget.grid_info()
            if int(info.get("row", 0)) > 0:
                widget.destroy()

        self.filas_aprobacion_remarque = []
        self.fila_actual_aprobacion_remarque = 1

        for prenda in prendas:
            fila_grid = self.fila_actual_aprobacion_remarque
            valores = [
                (prenda["codigo_prenda"], 0, 90),
                (prenda["articulo"], 1, 130),
                (self._format_moneda(prenda["precio_actual"]), 2, 100),
                (self._format_moneda(prenda["precio_remarcado"]), 3, 120),
                (self._formatear_fecha(prenda["fecha_remarque"]), 4, 100),
            ]

            for texto, columna, ancho in valores:
                ctk.CTkLabel(
                    self.frame_tabla_aprobacion_remarque,
                    text=str(texto),
                    text_color=COLOR_TEXTO,
                    width=ancho,
                    anchor="w"
                ).grid(row=fila_grid, column=columna, padx=4, pady=4, sticky="w")

            decision_var = ctk.StringVar(value="PENDIENTE")
            option_decision = ctk.CTkOptionMenu(
                self.frame_tabla_aprobacion_remarque,
                values=self.decision_remarque_opciones,
                variable=decision_var,
                width=120,
                fg_color=COLOR_FONDO,
                button_color=COLOR_FONDO,
                button_hover_color=COLOR_HOVER,
                text_color=COLOR_TEXTO,
                dropdown_fg_color=COLOR_FONDO,
                dropdown_text_color=COLOR_TEXTO
            )
            option_decision.grid(row=fila_grid, column=5, padx=4, pady=4, sticky="w")

            self.filas_aprobacion_remarque.append({
                "codigo_prenda": prenda["codigo_prenda"],
                "decision_var": decision_var,
            })
            self.fila_actual_aprobacion_remarque += 1

    def _mostrar_prendas_vencidas(self, texto):
        self.textbox_prendas_vencidas.configure(state="normal")
        self.textbox_prendas_vencidas.delete("1.0", "end")
        self.textbox_prendas_vencidas.insert("1.0", texto)
        self.textbox_prendas_vencidas.configure(state="disabled")

    def _construir_tabla_prendas_vencidas(self, filas):
        if not filas:
            return "Todavía no hay resultados para mostrar."

        columnas = [
            ("FECHA ING", 12),
            ("PROV", 10),
            ("CODIGO", 14),
            ("ARTICULO", 22),
            ("MARCA", 16),
            ("TALLE", 8),
            ("COLOR", 12),
            ("PRECIO", 12),
            ("DIAS", 8),
        ]

        encabezado = " ".join(titulo.ljust(ancho) for titulo, ancho in columnas)
        separador = "-" * len(encabezado)
        lineas = [encabezado, separador]

        for fila in filas:
            valores = [
                self._formatear_fecha(fila["fecha_ingreso"]),
                fila["codigo_proveedora"],
                fila["codigo_prenda"],
                fila["articulo"],
                fila["marca"],
                fila["talle"],
                fila["color"],
                self._format_moneda(fila["precio"]),
                str(fila["dias_en_inventario"]),
            ]
            fila_texto = " ".join(
                str(valor)[:ancho].ljust(ancho)
                for valor, (_, ancho) in zip(valores, columnas)
            )
            lineas.append(fila_texto)

        return "\n".join(lineas)

    def buscar_prenda_para_devolucion(self):
        codigo_prenda = self.entry_devolucion_codigo.get().strip()
        resultado, datos = obtener_prenda_desde_gui(codigo_prenda)

        if not resultado:
            self.label_estado_devolucion.configure(text=datos, text_color="#B00020")
            self._limpiar_formulario_devolucion()
            return

        self.label_estado_devolucion.configure(
            text=f"Prenda {datos['codigo_prenda']} lista para revisar.",
            text_color="#2E5E2E"
        )
        self.label_detalle_devolucion.configure(
            text=(
                f"Proveedora: {datos['codigo_proveedora']}\n"
                f"Artículo: {datos['articulo']}\n"
                f"Marca: {datos['marca']} | Talle: {datos['talle']} | Color: {datos['color']}\n"
                f"Precio: {self._format_moneda(datos['precio'])}\n"
                f"Estado actual: {datos['estado'] or '-'}"
            ),
            text_color=COLOR_TEXTO
        )

    def registrar_devolucion_desde_pantalla(self):
        codigo_prenda = self.entry_devolucion_codigo.get().strip()
        if not codigo_prenda:
            self.label_estado_devolucion.configure(
                text="ERROR: CODIGO_PRENDA es obligatorio.",
                text_color="#B00020"
            )
            return

        confirmar = messagebox.askyesno(
            "Confirmar devolución",
            f"¿Querés registrar la devolución de la prenda {codigo_prenda}?"
        )
        if not confirmar:
            return

        resultado, datos = registrar_devolucion_desde_gui(codigo_prenda)
        if not resultado:
            self.label_estado_devolucion.configure(text=datos, text_color="#B00020")
            return

        self.label_estado_devolucion.configure(
            text=f"Devolución registrada para {datos['codigo_prenda']}.",
            text_color="#2E5E2E"
        )
        self.label_detalle_devolucion.configure(
            text=(
                f"Proveedora: {datos['codigo_proveedora']}\n"
                f"Artículo: {datos['articulo']}\n"
                f"Marca: {datos['marca']} | Talle: {datos['talle']} | Color: {datos['color']}\n"
                f"Precio: {self._format_moneda(datos['precio'])}\n"
                f"Estado actual: {datos['estado']}"
            ),
            text_color=COLOR_TEXTO
        )
        messagebox.showinfo(
            "Devolución registrada",
            "\n".join([
                f"Código prenda: {datos['codigo_prenda']}",
                f"Artículo: {datos['articulo']}",
                f"Estado actual: {datos['estado']}",
            ])
        )

    def buscar_proveedora_desde_pantalla(self):
        resultado, datos = obtener_proveedora_desde_gui(
            self.entry_buscar_proveedora.get().strip()
        )

        if not resultado:
            self.label_estado_ver_proveedora.configure(text=datos, text_color="#B00020")
            self._limpiar_datos_proveedora()
            self._mostrar_detalle_proveedora("Todavía no hay resultados para mostrar.")
            return

        self.label_estado_ver_proveedora.configure(
            text=f"Proveedora {datos['proveedora']['codigo_proveedora']} encontrada.",
            text_color="#2E5E2E"
        )

        proveedora = datos["proveedora"]
        resumen = datos["resumen_prendas"]
        self.detalle_proveedora_actual = datos["detalle_prendas"]
        self.datos_proveedora_vars["nombre_proveedora"].set(proveedora["nombre_proveedora"] or "-")
        self.datos_proveedora_vars["telefono"].set(proveedora["telefono"] or "-")
        self.datos_proveedora_vars["banco"].set(proveedora["banco"] or "-")
        self.datos_proveedora_vars["numero_cuenta"].set(proveedora["numero_cuenta"] or "-")
        self.datos_proveedora_vars["titular_cuenta"].set(proveedora["titular_cuenta"] or "-")
        self.datos_proveedora_vars["alias"].set(proveedora["alias"] or "-")
        self.datos_proveedora_vars["obs_ingreso"].set(proveedora["obs_ingreso"] or "-")
        self.datos_proveedora_vars["estado"].set(proveedora["estado"] or "-")
        self.datos_proveedora_vars["total_prendas"].set(str(resumen["total_prendas"]))
        self.datos_proveedora_vars["prendas_disponibles"].set(str(resumen["prendas_disponibles"]))
        self.datos_proveedora_vars["prendas_vendidas"].set(str(resumen["prendas_vendidas"]))
        self.datos_proveedora_vars["prendas_devueltas"].set(str(resumen["prendas_devueltas"]))
        self.actualizar_filtro_proveedora()

    def mostrar_todas_las_proveedoras(self):
        resultado, datos = obtener_todas_las_proveedoras_desde_gui()

        if not resultado:
            self.label_estado_ver_proveedora.configure(text=datos, text_color="#B00020")
            self._limpiar_datos_proveedora()
            self._mostrar_detalle_proveedora("Todavía no hay resultados para mostrar.")
            return

        self.label_estado_ver_proveedora.configure(
            text=f"Se encontraron {len(datos)} proveedoras.",
            text_color="#2E5E2E"
        )
        self._limpiar_datos_proveedora()
        self.detalle_proveedora_actual = []
        self._mostrar_detalle_proveedora(self._construir_tabla_todas_proveedoras(datos))


    def validar_venta_pendiente_desde_pantalla(self):
        codigo_prenda = self.entry_validar_pendiente.get().strip()
        resultado, mensaje = validar_venta_pendiente_por_codigo(codigo_prenda)

        if not resultado:
            self.label_pendientes_estado.configure(text=mensaje, text_color="#B00020")
            return

        self.entry_validar_pendiente.delete(0, "end")
        self.label_pendientes_estado.configure(text=mensaje, text_color="#2E5E2E")
        self.buscar_ventas_pendientes()

    def buscar_rendicion(self):
        resultado, datos = calcular_rendicion_proveedora(
            self.entry_rendicion_mes.get().strip(),
            self.entry_rendicion_anio.get().strip(),
            self.entry_rendicion_codigo.get().strip()
        )

        if not resultado:
            self.rendicion_actual = None
            self.boton_exportar_rendicion.configure(state="disabled")
            self.label_rendicion_estado.configure(text=datos, text_color="#B00020")
            self._limpiar_resumen_rendicion()
            self._mostrar_detalle_rendicion("No se encontraron resultados para mostrar.")
            return

        self.rendicion_actual = datos
        self.boton_exportar_rendicion.configure(state="normal")
        self.label_rendicion_estado.configure(
            text=(
                f"Rendición generada para {datos['codigo_proveedora']} "
                f"({datos['mes']:02d}/{datos['anio']})."
            ),
            text_color="#2E5E2E"
        )

        nombre = datos["nombre_proveedora"] or "-"
        self.resumen_vars["nombre_proveedora"].set(nombre)
        self.resumen_vars["cantidad_prendas"].set(str(datos["cantidad_prendas"]))
        self.resumen_vars["total_vendido"].set(self._format_moneda(datos["total_vendido"]))
        self.resumen_vars["comision_proveedora"].set(self._format_moneda(datos["comision_proveedora"]))
        self.resumen_vars["comision_fashion_reset"].set(self._format_moneda(datos["comision_fashion_reset"]))
        self.resumen_vars["total_descuentos"].set(self._format_moneda(datos["total_descuentos"]))
        self.resumen_vars["saldo_final"].set(self._format_moneda(datos["saldo_final"]))

        self._mostrar_detalle_rendicion_tabla(datos["ventas_rendicion"])

    def exportar_rendicion_actual(self):
        if not self.rendicion_actual:
            messagebox.showwarning("Rendición", "Primero generá una rendición.")
            return

        ventas_exportacion = [
            [
                venta["codigo_prenda"],
                venta["articulo"],
                venta["precio_venta"],
            ]
            for venta in self.rendicion_actual["ventas_rendicion"]
        ]

        resultado, mensaje = exportar_rendicion_excel(
            self.rendicion_actual["mes"],
            self.rendicion_actual["anio"],
            self.rendicion_actual["codigo_proveedora"],
            ventas_exportacion,
            self.rendicion_actual["cantidad_prendas"],
            self.rendicion_actual["total_vendido"],
            self.rendicion_actual["comision_proveedora"],
            self.rendicion_actual["comision_fashion_reset"],
            self.rendicion_actual["total_descuentos"],
            self.rendicion_actual["saldo_final"]
        )
        if not resultado:
            messagebox.showerror("Exportar rendición", mensaje)
            return

        messagebox.showinfo("Exportar rendición", f"Archivo generado: {mensaje}")

    def _limpiar_resumen_rendicion(self):
        for variable in self.resumen_vars.values():
            variable.set("-")

    def _limpiar_resumen_general(self):
        for variable in self.resumen_general_vars.values():
            variable.set("-")

    def _limpiar_pendientes_validacion(self):
        for variable in self.pendientes_vars.values():
            variable.set("-")

    def _limpiar_datos_proveedora(self):
        for variable in self.datos_proveedora_vars.values():
            variable.set("-")

    def _limpiar_formulario_editar_venta(self):
        if not hasattr(self, "editar_venta_entries"):
            return
        for entry in self.editar_venta_entries.values():
            entry.delete(0, "end")
        if hasattr(self, "editar_venta_tipo_pago"):
            self.editar_venta_tipo_pago.set("EFECTIVO")
        if hasattr(self, "editar_venta_validacion"):
            self.editar_venta_validacion.set("PENDIENTE")
        if hasattr(self, "label_editar_venta_info"):
            self.label_editar_venta_info.configure(
                text="Todavía no hay una venta cargada para editar.",
                text_color="#666666"
            )

    def _limpiar_formulario_editar_prenda(self):
        if not hasattr(self, "editar_prenda_entries"):
            return
        for entry in self.editar_prenda_entries.values():
            entry.delete(0, "end")
        if hasattr(self, "label_editar_prenda_info"):
            self.label_editar_prenda_info.configure(
                text="Todavía no hay una prenda cargada para editar.",
                text_color="#666666"
            )

    def _limpiar_formulario_devolucion(self):
        if hasattr(self, "label_detalle_devolucion"):
            self.label_detalle_devolucion.configure(
                text="Todavía no hay una prenda cargada para devolución.",
                text_color="#666666"
            )

    def _limpiar_formulario_remarque(self):
        self.lote_remarque_actual = None
        if hasattr(self, "remarque_vars"):
            for variable in self.remarque_vars.values():
                variable.set("-")
        if hasattr(self, "frame_tabla_remarque"):
            for widget in self.frame_tabla_remarque.winfo_children():
                info = widget.grid_info()
                if int(info.get("row", 0)) > 0:
                    widget.destroy()
        self.filas_remarque = []
        self.fila_actual_remarque = 1

    def _limpiar_formulario_aprobacion_remarque(self):
        self.lote_aprobacion_remarque_actual = None
        if hasattr(self, "aprobacion_remarque_vars"):
            for variable in self.aprobacion_remarque_vars.values():
                variable.set("-")
        if hasattr(self, "frame_tabla_aprobacion_remarque"):
            for widget in self.frame_tabla_aprobacion_remarque.winfo_children():
                info = widget.grid_info()
                if int(info.get("row", 0)) > 0:
                    widget.destroy()
        self.filas_aprobacion_remarque = []
        self.fila_actual_aprobacion_remarque = 1

    def _mostrar_desglose_resumen_general(self, texto):
        self.textbox_resumen_general.configure(state="normal")
        self.textbox_resumen_general.delete("1.0", "end")
        self.textbox_resumen_general.insert("1.0", texto)
        self.textbox_resumen_general.configure(state="disabled")

    def _mostrar_ventas_pendientes(self, texto):
        self.textbox_pendientes.configure(state="normal")
        self.textbox_pendientes.delete("1.0", "end")
        self.textbox_pendientes.insert("1.0", texto)
        self.textbox_pendientes.configure(state="disabled")

    def _mostrar_detalle_proveedora(self, texto):
        self.textbox_detalle_proveedora.configure(state="normal")
        self.textbox_detalle_proveedora.delete("1.0", "end")
        self.textbox_detalle_proveedora.insert("1.0", texto)
        self.textbox_detalle_proveedora.configure(state="disabled")

    def actualizar_filtro_proveedora(self):
        if not hasattr(self, "detalle_proveedora_actual"):
            return

        filtro = self.filtro_estado_proveedora.get().strip().upper()
        filas = self.detalle_proveedora_actual

        if filtro == "DISPONIBLES":
            filas = [
                fila for fila in filas
                if str(fila.get("estado") or "").strip().upper() == "DISPONIBLE"
            ]
        elif filtro == "VENDIDAS":
            filas = [
                fila for fila in filas
                if str(fila.get("estado") or "").strip().upper() == "VENDIDO"
            ]

        self._mostrar_detalle_proveedora(self._construir_tabla_proveedora(filas))

    def _construir_tabla_resumen_general(self, filas):
        if not filas:
            return "Todavía no hay resultados para mostrar."

        columnas = [
            ("CODIGO", 12),
            ("CANTIDAD", 12),
            ("TOTAL VENDIDO", 14),
            ("COMISION", 14),
            ("DESCUENTOS", 14),
            ("SALDO FINAL", 14),
        ]

        encabezado = " | ".join(titulo.ljust(ancho) for titulo, ancho in columnas)
        separador = "-+-".join("-" * ancho for _, ancho in columnas)
        tabla = [encabezado, separador]

        for fila in filas:
            valores = [
                str(fila["codigo_proveedora"] or ""),
                str(fila["cantidad_prendas"] or ""),
                self._format_moneda(fila["total_vendido"]),
                self._format_moneda(fila["comision_proveedora"]),
                self._format_moneda(fila["descuentos"]),
                self._format_moneda(fila["saldo_final"]),
            ]

            celdas = []
            for indice, valor in enumerate(valores):
                ancho = columnas[indice][1]
                texto = str(valor)
                if len(texto) > ancho:
                    texto = texto[: ancho - 1] + "."
                celdas.append(texto.ljust(ancho))

            tabla.append(" | ".join(celdas))

        return "\n".join(tabla)

    def _construir_tabla_pendientes(self, filas):
        if not filas:
            return "Todavía no hay resultados para mostrar."

        columnas = [
            ("FECHA", 12),
            ("PROVEEDORA", 12),
            ("CODIGO", 12),
            ("ARTICULO", 18),
            ("PRECIO", 12),
            ("CLIENTE", 18),
            ("TIPO PAGO", 18),
            ("VALIDACION", 12),
        ]

        encabezado = " | ".join(titulo.ljust(ancho) for titulo, ancho in columnas)
        separador = "-+-".join("-" * ancho for _, ancho in columnas)
        tabla = [encabezado, separador]

        for fila in filas:
            valores = [
                self._formatear_fecha(fila["fecha_venta"]),
                str(fila["codigo_proveedora"] or ""),
                str(fila["codigo_prenda"] or ""),
                str(fila["articulo"] or ""),
                self._format_moneda(fila["precio_venta"]),
                str(fila["cliente"] or ""),
                str(fila["tipo_pago"] or ""),
                str(fila["validacion"] or ""),
            ]

            celdas = []
            for indice, valor in enumerate(valores):
                ancho = columnas[indice][1]
                texto = str(valor)
                if len(texto) > ancho:
                    texto = texto[: ancho - 1] + "."
                celdas.append(texto.ljust(ancho))

            tabla.append(" | ".join(celdas))

        return "\n".join(tabla)

    def _construir_tabla_proveedora(self, filas):
        if not filas:
            return "Todavía no hay resultados para mostrar."

        columnas = [
            ("CODIGO", 12),
            ("ARTICULO", 16),
            ("MARCA", 14),
            ("TALLE", 10),
            ("COLOR", 12),
            ("PRECIO", 12),
            ("ESTADO", 12),
        ]

        encabezado = " | ".join(titulo.ljust(ancho) for titulo, ancho in columnas)
        separador = "-+-".join("-" * ancho for _, ancho in columnas)
        tabla = [encabezado, separador]

        for fila in filas:
            precio = fila["precio"]
            precio_texto = self._format_moneda(precio) if isinstance(precio, (int, float)) else str(precio or "")
            valores = [
                str(fila["codigo_prenda"] or ""),
                str(fila["articulo"] or ""),
                str(fila["marca"] or ""),
                str(fila["talle"] or ""),
                str(fila["color"] or ""),
                precio_texto,
                str(fila["estado"] or ""),
            ]

            celdas = []
            for indice, valor in enumerate(valores):
                ancho = columnas[indice][1]
                texto = str(valor)
                if len(texto) > ancho:
                    texto = texto[: ancho - 1] + "."
                celdas.append(texto.ljust(ancho))

            tabla.append(" | ".join(celdas))

        return "\n".join(tabla)

    def _construir_tabla_todas_proveedoras(self, filas):
        if not filas:
            return "Todavía no hay resultados para mostrar."

        columnas = [
            ("CODIGO", 10),
            ("NOMBRE", 18),
            ("TELEFONO", 14),
            ("BANCO", 14),
            ("NUMERO CTA", 16),
            ("TITULAR", 18),
            ("ALIAS", 16),
            ("ESTADO", 10),
        ]

        encabezado = " | ".join(titulo.ljust(ancho) for titulo, ancho in columnas)
        separador = "-+-".join("-" * ancho for _, ancho in columnas)
        tabla = [encabezado, separador]

        for fila in filas:
            valores = [
                str(fila["codigo_proveedora"] or ""),
                str(fila["nombre_proveedora"] or ""),
                str(fila["telefono"] or ""),
                str(fila["banco"] or ""),
                str(fila["numero_cuenta"] or ""),
                str(fila["titular_cuenta"] or ""),
                str(fila["alias"] or ""),
                str(fila["estado"] or ""),
            ]

            celdas = []
            for indice, valor in enumerate(valores):
                ancho = columnas[indice][1]
                texto = str(valor)
                if len(texto) > ancho:
                    texto = texto[: ancho - 1] + "."
                celdas.append(texto.ljust(ancho))

            tabla.append(" | ".join(celdas))

        return "\n".join(tabla)

    def _mostrar_detalle_rendicion(self, texto):
        self.textbox_rendicion.configure(state="normal")
        self.textbox_rendicion.delete("1.0", "end")
        self.textbox_rendicion.insert("1.0", texto)
        self.textbox_rendicion.configure(state="disabled")

    def _mostrar_resultado_eliminar_prendas(self, texto):
        self.textbox_resultado_eliminar_prendas.configure(state="normal")
        self.textbox_resultado_eliminar_prendas.delete("1.0", "end")
        self.textbox_resultado_eliminar_prendas.insert("1.0", texto)
        self.textbox_resultado_eliminar_prendas.configure(state="disabled")

    def _mostrar_resultado_reversar_ventas(self, texto):
        self.textbox_resultado_reversar_ventas.configure(state="normal")
        self.textbox_resultado_reversar_ventas.delete("1.0", "end")
        self.textbox_resultado_reversar_ventas.insert("1.0", texto)
        self.textbox_resultado_reversar_ventas.configure(state="disabled")

    def _mostrar_detalle_rendicion_tabla(self, ventas):
        self._mostrar_detalle_rendicion(self._construir_tabla_rendicion(ventas))

    def _construir_tabla_rendicion(self, ventas):
        if not ventas:
            return "Todavía no hay resultados para mostrar."

        columnas = [
            ("FECHA", 12),
            ("CODIGO", 12),
            ("ARTICULO", 18),
            ("PRECIO", 12),
            ("CLIENTE", 18),
            ("TIPO PAGO", 22),
            ("VALIDACION", 12),
        ]

        encabezado = " | ".join(titulo.ljust(ancho) for titulo, ancho in columnas)
        separador = "-+-".join("-" * ancho for _, ancho in columnas)
        filas = [encabezado, separador]

        for venta in ventas:
            valores = [
                self._formatear_fecha(venta["fecha_venta"]),
                str(venta["codigo_prenda"] or ""),
                str(venta["articulo"] or ""),
                self._format_moneda(venta["precio_venta"]),
                str(venta["cliente"] or ""),
                str(venta["tipo_pago"] or ""),
                str(venta["validacion"] or ""),
            ]

            celdas = []
            for indice, valor in enumerate(valores):
                ancho = columnas[indice][1]
                texto = str(valor)
                if len(texto) > ancho:
                    texto = texto[: ancho - 1] + "."
                celdas.append(texto.ljust(ancho))

            filas.append(" | ".join(celdas))

        return "\n".join(filas)

    def _construir_resumen_eliminar_prendas(self, datos):
        lineas = [f"ELIMINADAS: {datos['cantidad_eliminadas']}"]

        if datos["eliminadas"]:
            lineas.append("")
            lineas.append("PRENDAS ELIMINADAS:")
            for fila in datos["eliminadas"]:
                lineas.append(
                    f"{fila['codigo_prenda']} | {fila['articulo'] or '-'} | {fila['estado'] or '-'}"
                )

        if datos["errores"]:
            lineas.append("")
            lineas.append("ERRORES:")
            lineas.extend(datos["errores"])

        return "\n".join(lineas)

    def _construir_resumen_reversar_ventas(self, datos):
        lineas = [f"REVERSADAS: {datos['cantidad_reversadas']}"]

        if datos["reversadas"]:
            lineas.append("")
            lineas.append("VENTAS REVERSADAS:")
            for fila in datos["reversadas"]:
                lineas.append(f"{fila['codigo_prenda']} | {fila['articulo'] or '-'}")

        if datos["errores"]:
            lineas.append("")
            lineas.append("ERRORES:")
            lineas.extend(datos["errores"])

        return "\n".join(lineas)

    def _mover_foco(self, filas_tabla, fila_indice, col_indice):
        if fila_indice < 0 or fila_indice >= len(filas_tabla):
            return
        if col_indice < 0 or col_indice >= len(filas_tabla[fila_indice]):
            return
        filas_tabla[fila_indice][col_indice].focus_set()

    def _mover_foco_siguiente(self, filas_tabla, fila_indice, col_indice):
        if col_indice + 1 < len(filas_tabla[fila_indice]):
            self._mover_foco(filas_tabla, fila_indice, col_indice + 1)
        elif fila_indice + 1 < len(filas_tabla):
            self._mover_foco(filas_tabla, fila_indice + 1, 0)

    def _mover_foco_anterior(self, filas_tabla, fila_indice, col_indice):
        if col_indice - 1 >= 0:
            self._mover_foco(filas_tabla, fila_indice, col_indice - 1)
        elif fila_indice - 1 >= 0:
            self._mover_foco(filas_tabla, fila_indice - 1, len(filas_tabla[fila_indice - 1]) - 1)

    def _navegar_entrada(self, event, filas_tabla, fila_indice, col_indice):
        if event.keysym in ("Return", "KP_Enter", "Tab") and not (event.keysym == "Tab" and event.state & 0x0001):
            self._mover_foco_siguiente(filas_tabla, fila_indice, col_indice)
        elif event.keysym in ("ISO_Left_Tab", "Tab") and event.state & 0x0001:
            self._mover_foco_anterior(filas_tabla, fila_indice, col_indice)
        elif event.keysym == "Right":
            self._mover_foco(filas_tabla, fila_indice, col_indice + 1)
        elif event.keysym == "Left":
            self._mover_foco(filas_tabla, fila_indice, col_indice - 1)
        elif event.keysym == "Down":
            self._mover_foco(filas_tabla, fila_indice + 1, col_indice)
        elif event.keysym == "Up":
            self._mover_foco(filas_tabla, fila_indice - 1, col_indice)
        return "break"

    def _autocompletar_venta(self, event, fila_indice):
        if fila_indice < 0 or fila_indice >= len(self.filas_venta):
            return

        fila = self.filas_venta[fila_indice]
        entry_codigo = fila[0]
        codigo_prenda = entry_codigo.get().strip()

        for entrada in fila:
            try:
                entrada.configure(border_color=COLOR_TEXTO)
            except Exception:
                pass

        if not codigo_prenda:
            for entrada in fila[1:6]:
                entrada.configure(state="normal")
                entrada.delete(0, "end")
                entrada.configure(state="disabled")
            return

        resultado, datos = buscar_ingreso_por_codigo(codigo_prenda)
        if not resultado:
            entry_codigo.configure(border_color="red")
            for entrada in fila[1:6]:
                entrada.configure(state="normal")
                entrada.delete(0, "end")
                entrada.configure(state="disabled")
            return

        entry_codigo.configure(border_color=COLOR_TEXTO)
        fila[1].configure(state="normal")
        fila[1].delete(0, "end")
        fila[1].insert(0, datos["articulo"])
        fila[1].configure(state="disabled")

        fila[2].configure(state="normal")
        fila[2].delete(0, "end")
        fila[2].insert(0, datos["marca"])
        fila[2].configure(state="disabled")

        fila[3].configure(state="normal")
        fila[3].delete(0, "end")
        fila[3].insert(0, datos["talle"])
        fila[3].configure(state="disabled")

        fila[4].configure(state="normal")
        fila[4].delete(0, "end")
        fila[4].insert(0, datos["color"])
        fila[4].configure(state="disabled")

        fila[5].configure(state="normal")
        fila[5].delete(0, "end")
        fila[5].insert(0, self._format_precio(datos["precio_lista"]))
        fila[5].configure(state="disabled")

        self._actualizar_total_venta()

    def _autocompletar_y_navegar_venta(self, event, fila_indice, col_indice):
        self._autocompletar_venta(event, fila_indice)
        return self._navegar_entrada(event, self.filas_venta, fila_indice, col_indice)

    def _parse_precio(self, texto):
        texto = str(texto or "").strip()
        if not texto:
            return None
        texto = texto.replace(".", "").replace(",", "")
        if not texto.isdigit():
            return None
        return int(texto)

    def _format_precio(self, texto):
        valor = self._parse_precio(texto)
        if valor is None:
            return str(texto)
        return f"{valor:,}".replace(",", ".")

    def _format_moneda(self, valor):
        return f"${self._format_precio(valor)}"

    def _formatear_fecha(self, valor):
        if isinstance(valor, datetime):
            return valor.strftime("%d/%m/%Y")
        texto = str(valor or "").strip()
        if not texto:
            return ""
        for formato in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(texto, formato).strftime("%d/%m/%Y")
            except ValueError:
                pass
        return texto

    def _formatear_precio_entry(self, entry):
        texto = entry.get()
        if not texto:
            return
        entry.delete(0, "end")
        entry.insert(0, self._format_precio(texto))
        self._actualizar_total_venta()

    def _formatear_precio_entry_remarque(self, entry):
        texto = entry.get()
        if not texto:
            return
        entry.delete(0, "end")
        entry.insert(0, self._format_precio(texto))

    def _actualizar_tipo_pago_venta(self, tipo_pago, entry_obs, validacion_widget):
        if tipo_pago in {"EFECTIVO", "DESCUENTO A PROVEEDORA"}:
            validacion_widget.set("PAGADO")
        else:
            validacion_widget.set("PENDIENTE")

        if tipo_pago == "DESCUENTO A PROVEEDORA":
            entry_obs.configure(
                placeholder_text="Ingresar código",
                placeholder_text_color="#8A8A8A",
                text_color=COLOR_TEXTO
            )
        else:
            entry_obs.configure(
                placeholder_text="",
                placeholder_text_color="#8A8A8A",
                text_color=COLOR_TEXTO
            )

    def volver_menu_principal(self):
        self._ocultar_frames_secundarios()
        self.frame_botones.pack(pady=30, padx=30, fill="both", expand=True)

    def agregar_filas_ingreso(self, cantidad=10):
        for _ in range(cantidad):
            entry_numero = ctk.CTkEntry(self.frame_tabla_ingreso, width=80)
            entry_numero.grid(row=self.fila_actual_ingreso, column=0, padx=5, pady=4)

            entry_articulo = ctk.CTkEntry(self.frame_tabla_ingreso, width=140)
            entry_articulo.grid(row=self.fila_actual_ingreso, column=1, padx=5, pady=4)

            entry_marca = ctk.CTkEntry(self.frame_tabla_ingreso, width=120)
            entry_marca.grid(row=self.fila_actual_ingreso, column=2, padx=5, pady=4)

            entry_talle = ctk.CTkEntry(self.frame_tabla_ingreso, width=80)
            entry_talle.grid(row=self.fila_actual_ingreso, column=3, padx=5, pady=4)

            entry_color = ctk.CTkEntry(self.frame_tabla_ingreso, width=100)
            entry_color.grid(row=self.fila_actual_ingreso, column=4, padx=5, pady=4)

            entry_precio = ctk.CTkEntry(self.frame_tabla_ingreso, width=100)
            entry_precio.grid(row=self.fila_actual_ingreso, column=5, padx=5, pady=4)

            entry_obs = ctk.CTkEntry(self.frame_tabla_ingreso, width=160)
            entry_obs.grid(row=self.fila_actual_ingreso, column=6, padx=5, pady=4)

            fila_indice = len(self.filas_ingreso)
            fila_entradas = [
                entry_numero,
                entry_articulo,
                entry_marca,
                entry_talle,
                entry_color,
                entry_precio,
                entry_obs
            ]

            for col_indice, entrada in enumerate(fila_entradas):
                entrada.bind("<Return>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<KP_Enter>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<Tab>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<Shift-Tab>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<Right>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<Left>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<Down>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))
                entrada.bind("<Up>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_ingreso, r, c))

            self.filas_ingreso.append(fila_entradas)
            self.fila_actual_ingreso += 1

    def agregar_filas_venta(self, cantidad=10):
        for _ in range(cantidad):
            entry_codigo = ctk.CTkEntry(self.frame_tabla_venta, width=60)
            entry_codigo.grid(row=self.fila_actual_venta, column=0, padx=5, pady=4)

            entry_articulo = ctk.CTkEntry(self.frame_tabla_venta, width=75, state="disabled")
            entry_articulo.grid(row=self.fila_actual_venta, column=1, padx=3, pady=4)

            entry_marca = ctk.CTkEntry(self.frame_tabla_venta, width=70, state="disabled")
            entry_marca.grid(row=self.fila_actual_venta, column=2, padx=3, pady=4)

            entry_talle = ctk.CTkEntry(self.frame_tabla_venta, width=55, state="disabled")
            entry_talle.grid(row=self.fila_actual_venta, column=3, padx=3, pady=4)

            entry_color = ctk.CTkEntry(self.frame_tabla_venta, width=55, state="disabled")
            entry_color.grid(row=self.fila_actual_venta, column=4, padx=3, pady=4)

            entry_precio = ctk.CTkEntry(self.frame_tabla_venta, width=70, state="disabled")
            entry_precio.grid(row=self.fila_actual_venta, column=5, padx=3, pady=4)
            entry_precio.bind("<KeyRelease>", lambda event, e=entry_precio: self._actualizar_total_venta())
            entry_precio.bind("<FocusOut>", lambda event, e=entry_precio: self._formatear_precio_entry(e))

            entry_obs_venta = ctk.CTkEntry(
                self.frame_tabla_venta,
                width=85,
                placeholder_text="",
                placeholder_text_color="#8A8A8A"
            )
            entry_obs_venta.grid(row=self.fila_actual_venta, column=8, padx=3, pady=4)

            tipo_pago_var = ctk.StringVar(value="EFECTIVO")
            validacion_var = ctk.StringVar(value="PAGADO")
            entry_tipo_pago = ctk.CTkOptionMenu(
                self.frame_tabla_venta,
                values=self.tipo_pago_opciones,
                variable=tipo_pago_var,
                width=110,
                fg_color=COLOR_FONDO,
                button_color=COLOR_FONDO,
                button_hover_color=COLOR_HOVER,
                text_color=COLOR_TEXTO,
                dropdown_fg_color=COLOR_FONDO,
                dropdown_text_color=COLOR_TEXTO,
                command=lambda value, obs=entry_obs_venta, validacion=validacion_var: self._actualizar_tipo_pago_venta(
                    value,
                    obs,
                    validacion
                )
            )
            entry_tipo_pago.grid(row=self.fila_actual_venta, column=6, padx=3, pady=4)

            entry_validacion = ctk.CTkOptionMenu(
                self.frame_tabla_venta,
                values=self.validacion_opciones,
                variable=validacion_var,
                width=65,
                fg_color=COLOR_FONDO,
                button_color=COLOR_FONDO,
                button_hover_color=COLOR_HOVER,
                text_color=COLOR_TEXTO,
                dropdown_fg_color=COLOR_FONDO,
                dropdown_text_color=COLOR_TEXTO
            )
            entry_validacion.grid(row=self.fila_actual_venta, column=7, padx=3, pady=4)
            self._actualizar_tipo_pago_venta("EFECTIVO", entry_obs_venta, validacion_var)

            fila_indice = len(self.filas_venta)
            fila_entradas = [
                entry_codigo,
                entry_articulo,
                entry_marca,
                entry_talle,
                entry_color,
                entry_precio,
                entry_tipo_pago,
                entry_validacion,
                entry_obs_venta
            ]

            for col_indice, entrada in enumerate(fila_entradas):
                entrada.bind("<Return>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<KP_Enter>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<Tab>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<Shift-Tab>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<Right>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<Left>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<Down>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))
                entrada.bind("<Up>", lambda event, r=fila_indice, c=col_indice: self._navegar_entrada(event, self.filas_venta, r, c))

            entry_codigo.bind("<Return>", lambda event, r=fila_indice, c=0: self._autocompletar_y_navegar_venta(event, r, c))
            entry_codigo.bind("<KP_Enter>", lambda event, r=fila_indice, c=0: self._autocompletar_y_navegar_venta(event, r, c))
            entry_codigo.bind("<Tab>", lambda event, r=fila_indice, c=0: self._autocompletar_y_navegar_venta(event, r, c))
            entry_codigo.bind("<FocusOut>", lambda event, r=fila_indice: self._autocompletar_venta(event, r))

            self.filas_venta.append(fila_entradas)
            self.fila_actual_venta += 1

        self._actualizar_total_venta()

    def _actualizar_total_venta(self):
        total = 0
        for fila in self.filas_venta:
            texto_precio = fila[5].get().strip()
            if not texto_precio:
                continue
            valor = self._parse_precio(texto_precio)
            if valor is None:
                continue
            total += valor

        if hasattr(self, "label_total_venta"):
            self.label_total_venta.configure(text=f"TOTAL DEL LOTE: ${total:,}".replace(",", "."))


if __name__ == "__main__":
    app = AppFashionReset()
    app.mainloop()
