import os
import docx
import sys
import getpass

# ---------- Declaraciones ----------
user = getpass.getuser()
ruta_de_proyectos = "C:\\Users\\{}\\AndroidStudioProjects".format(user)
escritorio = "C:\\Users\\{}\\Desktop".format(user)
# Usuarios de Mac:
# ruta_de_proyectos = "//Users//{}//AndroidStudioProjects".format(user)
# escritorio = "//Users//{}//Desktop".format(user)

def filtro_por_extension(nombre):
    extensiones_a_incluir = [".java", ".xml", ".png", ".jpg", ".mp3", ".mp4"]
    nombre_extension = os.path.splitext(nombre)[-1].lower()
    return nombre_extension in extensiones_a_incluir or nombre == "build.gradle"

def filtro_por_archivo(nombre):
    excluir = ["ic_launcher_background.xml", "ic_launcher_foreground.xml", "ExampleUnitTest",
               "ExampleInstrumentedTest.java", "BuildConfig.java", "ExampleUnitTest.java"]
    return nombre not in excluir

def filtro_por_carpeta(root, proyecto_i):
    return (root.endswith("\layout") or root.endswith("\main")
            or root.endswith(proyecto_i) or root.endswith("drawable")
            or root.endswith("drawable-v24") or root.endswith("anim")
            or root.endswith("raw") or "\main\java" in root)

def es_imagen(nombre):
    extensiones = [".png", ".jpg"]
    nombre_extension = os.path.splitext(nombre)[-1].lower()
    return nombre_extension in extensiones

def es_mp3(nombre):
    extensiones = [".mp3"]
    nombre_extension = os.path.splitext(nombre)[-1].lower()
    return nombre_extension in extensiones

def es_mp4(nombre):
    extensiones = [".mp4"]
    nombre_extension = os.path.splitext(nombre)[-1].lower()
    return nombre_extension in extensiones

# ---------- Ejecución ----------

try:
    os.mkdir(os.path.join(escritorio, "Android a Word"))
except:
    pass

# Obtener lista de proyectos de Android -->
try:
    os.chdir(ruta_de_proyectos)
except:
    print("Esa ruta no existe. Verifica el nombre de usuario.")
    sys.exit()
proyectos = os.listdir()

# Cambiar a folder nuevo en escritorio -->
try:
    os.chdir(os.path.join(escritorio, "Android a Word"))
except:
    print("Esa ruta no existe. Verifica el nombre de usuario.")
    sys.exit()

# Crear documentos por cada proyecto -->
for proyecto_i in proyectos:
    document = docx.Document()
    
    ruta_i = os.path.join(ruta_de_proyectos, proyecto_i)
    archivos_1 = {}
    for root, dirs, files in os.walk(ruta_i):
        if filtro_por_carpeta(root, proyecto_i):
            for name in files:
                if filtro_por_extension(name) and filtro_por_archivo(name):
                    archivos_1[name] = os.path.join(root, name)

    # Copiar contenido de archivos al documento -->
    document.add_heading(proyecto_i, level=1)
    for i in archivos_1:
        document.add_heading(i, level=2)
        if es_imagen(i):
            document.add_picture(archivos_1[i], width=docx.shared.Inches(1.25))
        elif es_mp3(i):
            document.add_paragraph("Archivo de sonido.")
        elif es_mp4(i):
            document.add_paragraph("Archivo de video.")
        else:
            f = open(archivos_1[i], "r")
            document.add_paragraph(f.read())
      
    document.save(proyecto_i + '.docx')
