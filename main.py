import tkinter as tk
from Modelo.model import DataModel
from Vista.view import DataView
from Controlador.controller import DataController

def main():
    """Inicia la aplicación."""
    try:
        root = tk.Tk()
        model = DataModel()
        view = DataView(root)
        controller = DataController(model, view)
        root.mainloop()
    except KeyboardInterrupt:
        print("Programa detenido por el usuario.")
    except Exception as e:
        print(f"Error inesperado: {e}")

if __name__ == "__main__":
    main()