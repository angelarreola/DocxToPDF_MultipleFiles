import os, sys
from docx2pdf import convert

rute = input("Ingrese la dirección donde se encuentran los '.docx' : ")
newDir = input("Ingrese el nombre de la carpeta en donde se guardaran los PDFs: ")
try:
    os.chdir(rute)
    os.mkdir(f"{rute}\{newDir}")
except FileExistsError:
    print(f"El directorio '{newDir}' ya existe, borrelo e intente correr el programa de nuevo.")
    sys.exit(1)
except FileNotFoundError:
    print("La ruta indicada no es valida.")
    sys.exit(1)
except Exception as e:
    print(type(e).__name__)
    sys.exit(1)

for file in os.listdir():
    name, ext = os.path.splitext(file)
    if ext == '.docx':
        convert(f"{rute}\{file}", f"{rute}\{newDir}\{name}.pdf")
        print(f"Se ha convertido el archivo: {name} a PDF de manera exitosa!")
        
print("Proceso de conversion finalizado con exito!")
