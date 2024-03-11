from tkinter import Tk
from tkinter.filedialog import askopenfilenames, askdirectory
import excelToJson

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filenames = askopenfilenames(filetypes=[("Excel files", ".xlsx .xlsm")], title="Seleccione las planillas que desea convertir")

if len(filenames) != 0:
    outputFolder = askdirectory(title="Seleccione la carpeta donde se guardará")

    for file in filenames:
        filename = file.split("/")
        filename.reverse()
        filename = filename[0]
        filename = filename.split(".")[0]

        data = excelToJson.readFile(file)
        json = excelToJson.exportJson(data)

        f = open( str(outputFolder + "/" + filename + ".json") , "w")
        f.write(json.replace("'", '"'))
        f.close()
