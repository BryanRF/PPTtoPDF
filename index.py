import os
from comtypes.client import CreateObject

def convertir_pptx_a_pdf(ruta_presentacion, ruta_salida):
    """
    Convierte un archivo PowerPoint (.pptx) a PDF.

    :param ruta_presentacion: Ruta completa del archivo .pptx.
    :param ruta_salida: Ruta completa del archivo PDF de salida.
    """
    if not ruta_salida.endswith(".pdf"):
        raise ValueError("La ruta de salida debe terminar en '.pdf'")

    powerpoint = CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1  # Hacer visible PowerPoint (opcional)

    try:
        presentacion = powerpoint.Presentations.Open(ruta_presentacion)
        presentacion.SaveAs(ruta_salida, 32)  # 32 es el formato para PDF
        print(f"Archivo convertido exitosamente: {ruta_salida}")
    except Exception as e:
        print(f"Error al convertir: {e}")
    finally:
        presentacion.Close()
        powerpoint.Quit()

# Ejemplo de uso
if __name__ == "__main__":
    # Obtener la ruta del directorio raíz del proyecto
    directorio_actual = os.path.dirname(os.path.abspath(__file__))

    # Rutas relativas a la raíz del proyecto
    ruta_presentacion = os.path.join(directorio_actual, "test.pptx")
    ruta_salida = os.path.join(directorio_actual, "test.pdf")

    convertir_pptx_a_pdf(ruta_presentacion, ruta_salida)
