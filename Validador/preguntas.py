from validacion.utils import escribir

def validar_requisitos_consolidacion():
    while True:
            while True:
                    opcion = input("Antes de consolidar los archivos, revisemos que los requisitos minimos se cumplan:\n"
                                "¿Ya estan cargados los archivos en la carpeta 'carpetas_OPS'?\n (S/N): ")
                    if opcion == "S" or opcion == "s":
                        break
                    elif opcion == "N" or opcion == "n":
                        escribir("Por favor, carga los archivos en la carpeta 'carpetas_OPS' antes de continuar con la consolidación y validación.")
                        escribir("Gracias por utilizar el programa🗝️VALIDADOR🗝️ de archivos para OPS. ¡Hasta luego!")
                        exit()  # Salir del programa si no se han cargado los archivos
                    else:
                        escribir("Selecciona una opcion valida.")
                        continue
            while True:
                opcion = input("¿Para el proceso de CORRESPONSALES la carpeta se llama 'CBS_L60'?\n (S/N): ")
                if opcion == "S" or opcion == "s":
                    break
                elif opcion == "N" or opcion == "n":
                    escribir("Por favor, asegúrate de que la carpeta de corresponsales se llame 'CBS_L60' para que el validador pueda identificarla correctamente durante la consolidación y validación.")
                    escribir("Gracias por utilizar el programa🗝️VALIDADOR🗝️ de archivos para OPS. ¡Hasta luego!")
                    exit()  # Salir del programa si la carpeta no se llama 'CBS_L60'
                else:
                    escribir("Selecciona una opcion valida.")
                    continue
            break