import openpyxl
import qrcode
from PIL import Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import simpledialog, messagebox

def generar_qr_with_code(datos, codigo):
    qr_data = '\n'.join([f"{key}: {value}" for key, value in datos.items()])
    qr = qrcode.make(qr_data)

    # Tamaño del código QR
    qr_width, qr_height = qr.size

    # Tamaño fijo del texto en píxeles, reducido a la mitad
    text_width = 500  # Ajusta según sea necesario
    text_height = 100  # Ajusta según sea necesario

    # Crear una nueva imagen lo suficientemente grande como para contener el código QR y el texto
    img_width = qr_width  # El ancho de la imagen será igual al ancho del código QR
    img_height = qr_height + text_height + 20  # Ajustar según sea necesario para el texto
    img = Image.new('RGB', (img_width, img_height), color='white')

    # Pegar el código QR en la nueva imagen
    img.paste(qr, (0, 0))

    # Agregar el texto (incluyendo 'cod:')
    texto_codigo = f'cod: {codigo}'
    font_size = 50  # Tamaño de fuente
    font = ImageFont.truetype("arial.ttf", font_size)  # Usar la fuente Arial con el tamaño determinado
    draw = ImageDraw.Draw(img)
    draw.text(((img_width - text_width) / 2, qr_height + 10), texto_codigo, font=font, fill='black')

    return img

def generar_codigos_seleccionados(codigos_a_generar):
    # Carga el archivo Excel
    try:
        workbook = openpyxl.load_workbook('C:/QR-Script/sillas-relevamiento.xlsx')
        sheet = workbook.active

        # Lista para almacenar los códigos generados
        codigos_generados = []

        # Recorre cada fila del archivo Excel
        for row in sheet.iter_rows(min_row=2, values_only=False):
            codigo = str(row[0].value)  # Se asume que el código está en la primera columna
            if len(codigo.split('.')) == 3:  # Verifica si el código tiene tres pares de números
                datos = {
                    "Codigo": codigo,
                    "Modelo": row[1].value,
                    "Submodelo": row[2].value,
                    "Descripcion": row[3].value,
                    "Fecha de compra": row[4].value,
                    "Marca": row[5].value,
                    "Proveedor": row[6].value,
                    "Fecha de vencimiento por garantia": row[7].value
                }
                if codigos_a_generar == "Todos" or codigo in codigos_a_generar:
                    qr_with_code = generar_qr_with_code(datos, codigo)
                    filename = f'C:/QR-Script/QR-Generados/{codigo}.png'
                    qr_with_code.save(filename)
                    print(f"El código QR se guardó como {filename}")
                    codigos_generados.append(codigo)

        # Mostrar los códigos generados
        messagebox.showinfo("Códigos generados", "Se generaron los QR para los siguientes códigos:\n\n- " + '\n- '.join(codigos_generados))
    except FileNotFoundError:
        print("No se encontró el archivo Excel 'sillas-relevamiento.xlsx' en la ubicación especificada.")

def generar_todo():
    generar_codigos_seleccionados("Todos")

def ingresar_codigo():
    codigos = []
    while True:
        codigo = simpledialog.askstring("Ingresar código", "Por favor, ingrese un código:")
        if codigo is None:
            break  
        elif codigo:
            codigos.append(codigo)
        else:
            messagebox.showwarning("Error", "Por favor, ingrese un código válido.")

    generar_codigos_seleccionados(codigos)

def main():
    root = tk.Tk()
    root.withdraw()  

    # Crear la ventana principal
    ventana = tk.Toplevel()
    ventana.geometry("400x200")
    # Botones
    boton_generar_todo = tk.Button(ventana, text="Generar todo", command=generar_todo)
    boton_generar_todo.pack(pady=5)

    boton_ingresar_codigo = tk.Button(ventana, text="Ingresar código", command=ingresar_codigo)
    boton_ingresar_codigo.pack(pady=5)

    # Mostrar la ventana
    ventana.mainloop()

if __name__ == "__main__":
    main()

