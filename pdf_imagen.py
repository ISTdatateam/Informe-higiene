import os
import re
import requests
import pandas as pd
from pdf2image import convert_from_bytes


def sanitize_filename(name):
    """Elimina caracteres inválidos para nombres de archivo"""
    return re.sub(r'[\\/*?:"<>|;]', "", str(name))


def download_pdf_from_gdrive(url):
    try:
        file_id = url.split('/d/')[1].split('/')[0]
        download_url = f"https://drive.google.com/uc?id={file_id}"
        response = requests.get(download_url, timeout=10)
        response.raise_for_status()
        return response.content
    except Exception as e:
        print(f"Error descargando PDF: {e}")
        return None


def pdf_to_images(pdf_bytes, output_dir):
    try:
        os.makedirs(output_dir, exist_ok=True)
        images = convert_from_bytes(pdf_bytes)

        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"{i + 1}.png")
            image.save(image_path, 'PNG')

        return True
    except Exception as e:
        print(f"Error convirtiendo PDF: {e}")
        return False


def process_certificates(csv_file, output_base_dir):
    # Leer CSV considerando posibles formatos
    try:
        df = pd.read_csv(csv_file, delimiter=';', header=None, names=['id', 'url'])
    except:
        df = pd.read_csv(csv_file)

    # Limpieza de datos
    df['id'] = df['id'].apply(sanitize_filename)
    df['url'] = df['url'].str.strip()

    print("Datos a procesar:\n", df.head())

    for index, row in df.iterrows():
        equipo_id = str(row['id']).strip()
        pdf_url = row['url'].strip()

        print(f"\nProcesando ID: {equipo_id}")

        # Validar URL
        if not pdf_url.startswith('http'):
            print(f"URL inválida: {pdf_url}")
            continue

        pdf_content = download_pdf_from_gdrive(pdf_url)

        if pdf_content:
            output_dir = os.path.join(output_base_dir, equipo_id)
            success = pdf_to_images(pdf_content, output_dir)

            if success:
                print(f"✓ Conversión exitosa en {output_dir}")
            else:
                print(f"✗ Falló conversión para {equipo_id}")
        else:
            print("No se pudo descargar el PDF")


if __name__ == "__main__":
    # Configuración
    CSV_FILE = "urlspdfs.csv"  # Cambiar por tu ruta real
    OUTPUT_DIR = "imagenes_pdf"

    # Crear directorio principal
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    try:
        process_certificates(CSV_FILE, OUTPUT_DIR)
    except KeyboardInterrupt:
        print("\nProceso detenido por el usuario")
    finally:
        print("Proceso finalizado")
