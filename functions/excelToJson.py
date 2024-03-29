import pandas
import time
import base64
import unicodedata

def strip_accent(text):
    return ''.join((c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn'))


def contains(collection, this_item):
    return this_item in collection

def separate_alpha_numeric(input_string):
    alpha_chars = "".join([char for char in input_string if char.isalpha()])
    numeric_chars = int(
        "".join([char for char in input_string if char.isdigit()])) - 1
    
    return [column_to_number(alpha_chars), numeric_chars]


def column_to_number(column_string):
    column_string = column_string.upper()
    result = -1

    # TODO: 
    for char in column_string:
        result = ((result + 1) * 26) + ((ord(char) - ord('A')))
    
    return result

def range(dataframe: pandas.DataFrame, start: str, end: str):
    start = separate_alpha_numeric(start)
    end = separate_alpha_numeric(end)

    return pandas.DataFrame(dataframe.iloc[start[1]-1:end[1], start[0]:end[0]+1].values)


def cell(dataframe: pandas.DataFrame, start: str):
    start = separate_alpha_numeric(start)

    return str(dataframe.iloc[start[1], start[0]])


def cellOfDF(dataframe: pandas.DataFrame, start: str):
    start = separate_alpha_numeric(start)

    return str(dataframe.iloc[start[1] - 1, start[0]])

def intCellOfDF(dataframe: pandas.DataFrame, start: str):
    start = separate_alpha_numeric(start)
    value = str(dataframe.iloc[start[1] - 1, start[0]])

    try:
        return int(value)
    except:
        return 0

def hasPointsOnMistic(dataframe: pandas.DataFrame):
    return (int(intCellOfDF(dataframe, "M101")) > 0)

def hasPointsOnPsichiq(dataframe: pandas.DataFrame):
    return (int(intCellOfDF(dataframe, "M117")) > 0)


def range_to_json(sheet: pandas.DataFrame, start: str, end: str, keys: str, values: str, name: str):
    items = {}
    keys_values = set()

    for row in range(sheet, start, end).values:

        row = pandas.DataFrame(row).transpose()
        key = cell(row, keys)

        if key not in keys_values and key != "" and key != "nan":
            items[strip_accent(key)] = strip_accent(cell(row, values))
            keys_values.add(key)

    result = " '" + name + "': " + str(items)
    return result


def getRangedCombatData(rangeBlock: pandas.DataFrame):
    tipo = "A1"
    nombre = "C1"
    municion = "C2"
    conocimiento = "A3"
    tamanio = "D3"
    critPrincipal = "A5"
    critSecundario = "B5"
    entereza = "C5"
    rotura = "D5"
    presencia = "E5"
    turno = "F2"
    ataque = "G2"
    defensa = "H2"
    defensaTipo = "I2"
    danio = "J2"
    calidad = "H5"
    especial = "H6"
    caracteristica = "A6"
    advertencia = "I6"

    items = ""

    if not cell(rangeBlock, nombre) == "nan":
        items = ", {"
        items += "'nombre': '" + cell(rangeBlock, nombre) + "',"
        items += "'tipo': '" + cell(rangeBlock, tipo) + "',"
        items += "'conocimiento': '" + cell(rangeBlock, conocimiento) + "',"
        items += "'tamanio': '" + cell(rangeBlock, tamanio) + "',"
        items += "'critPrincipal': '" + cell(rangeBlock, critPrincipal) + "',"
        items += "'critSecundario': '" + \
            cell(rangeBlock, critSecundario) + "',"
        items += "'entereza': '" + cell(rangeBlock, entereza) + "',"
        items += "'rotura': '" + cell(rangeBlock, rotura) + "',"
        items += "'presencia': '" + cell(rangeBlock, presencia) + "',"
        items += "'turno': '" + cell(rangeBlock, turno) + "',"
        items += "'ataque': '" + cell(rangeBlock, ataque) + "',"
        items += "'defensa': '" + cell(rangeBlock, defensa) + "',"
        items += "'defensaTipo': '" + cell(rangeBlock, defensaTipo) + "',"
        items += "'danio': '" + cell(rangeBlock, danio) + "',"
        items += "'calidad': '" + cell(rangeBlock, calidad) + "',"
        items += "'caracteristica': '" + \
            cell(rangeBlock, caracteristica) + "',"
        items += "'advertencia': '" + cell(rangeBlock, advertencia) + "',"
        items += "'municion': '" + cell(rangeBlock, municion) + "',"
        items += "'especial': '" + cell(rangeBlock, especial) + "'}"
    return items


def getBaseCombatData(rangeBlock: pandas.DataFrame):
    tipo = "A2"
    nombre = "B1"
    conocimiento = "A3"
    tamanio = "D3"
    critPrincipal = "A5"
    critSecundario = "B5"
    entereza = "C5"
    rotura = "D5"
    presencia = "E5"
    turno = "F3"
    ataque = "G3"
    defensa = "H3"
    defensaTipo = "I3"
    danio = "J3"
    calidad = "H5"
    especial = "H6"
    caracteristica = "A6"
    advertencia = "I6"

    items = ""

    if not cell(rangeBlock, nombre) == "nan":
        items = ", {"
        items += "'nombre': '" + cell(rangeBlock, nombre) + "',"
        items += "'tipo': '" + cell(rangeBlock, tipo) + "',"
        items += "'conocimiento': '" + cell(rangeBlock, conocimiento) + "',"
        items += "'tamanio': '" + cell(rangeBlock, tamanio) + "',"
        items += "'critPrincipal': '" + cell(rangeBlock, critPrincipal) + "',"
        items += "'critSecundario': '" + \
            cell(rangeBlock, critSecundario) + "',"
        items += "'entereza': '" + cell(rangeBlock, entereza) + "',"
        items += "'rotura': '" + cell(rangeBlock, rotura) + "',"
        items += "'presencia': '" + cell(rangeBlock, presencia) + "',"
        items += "'turno': '" + cell(rangeBlock, turno) + "',"
        items += "'ataque': '" + cell(rangeBlock, ataque) + "',"
        items += "'defensa': '" + cell(rangeBlock, defensa) + "',"
        items += "'defensaTipo': '" + cell(rangeBlock, defensaTipo) + "',"
        items += "'danio': '" + cell(rangeBlock, danio) + "',"
        items += "'calidad': '" + cell(rangeBlock, calidad) + "',"
        items += "'caracteristica': '" + \
            cell(rangeBlock, caracteristica) + "',"
        items += "'advertencia': '" + cell(rangeBlock, advertencia) + "',"
        items += "'municion': '-',"
        items += "'especial': '" + cell(rangeBlock, especial) + "'}"
    return items


def getUnarmedCombatData(sheet: pandas.DataFrame):
    unarmedBlock = range(sheet, "C20", "L25")

    nombre = "A1"
    conocimiento = "A2"
    critPrincipal = "A4"
    critSecundario = "B4"
    entereza = "C4"
    rotura = "D4"
    presencia = "E4"
    turno = "F2"
    ataque = "G2"
    defensa = "H2"
    defensaTipo = "I2"
    danio = "J2"

    items = "{"
    items += "'nombre': '" + cell(unarmedBlock, nombre) + "',"
    items += "'tipo': 'desarmado',"
    items += "'conocimiento': '" + cell(unarmedBlock, conocimiento) + "',"
    items += "'tamanio': 'Normal',"
    items += "'critPrincipal': '" + cell(unarmedBlock, critPrincipal) + "',"
    items += "'critSecundario': '" + cell(unarmedBlock, critSecundario) + "',"
    items += "'entereza': '" + cell(unarmedBlock, entereza) + "',"
    items += "'rotura': '" + cell(unarmedBlock, rotura) + "',"
    items += "'presencia': '" + cell(unarmedBlock, presencia) + "',"
    items += "'turno': '" + cell(unarmedBlock, turno) + "',"
    items += "'ataque': '" + cell(unarmedBlock, ataque) + "',"
    items += "'defensa': '" + cell(unarmedBlock, defensa) + "',"
    items += "'defensaTipo': '" + cell(unarmedBlock, defensaTipo) + "',"
    items += "'danio': '" + cell(unarmedBlock, danio) + "',"
    items += "'calidad': '-',"
    items += "'caracteristica': '-',"
    items += "'advertencia': '-',"
    items += "'municion': '-',"
    items += "'especial': '-'}"

    return items


def getPsychicProjectionAsWeapon(sheet: pandas.DataFrame):
    weaponBlock = range(sheet, "O12", "Q13")

    turno = "A1"
    ataque = "B1"
    defensa = "C1"

    items = ",{"
    items += "'nombre': 'Proyección Psiquica',"
    items += "'tipo': 'Psiquica',"
    items += "'conocimiento': 'Conocida',"
    items += "'tamanio': 'Normal',"
    items += "'critPrincipal': 'Ene',"
    items += "'critSecundario': 'Pen',"
    items += "'entereza': '999',"
    items += "'rotura': '0',"
    items += "'presencia': '0',"
    items += "'turno': '" + cell(weaponBlock, turno) + "',"
    items += "'ataque': '" + cell(weaponBlock, ataque) + "',"
    items += "'defensa': '" + cell(weaponBlock, defensa) + "',"
    items += "'defensaTipo': 'Par',"
    items += "'danio': '0',"
    items += "'calidad': '-',"
    items += "'caracteristica': '-',"
    items += "'advertencia': '-',"
    items += "'municion': '-',"
    items += "'variable': 'true',"
    items += "'especial': '-'}"
    return items


def getMagicProjectionAsWeapon(sheet: pandas.DataFrame):
    weaponBlock = range(sheet, "O12", "Q13")

    turno = "A1"
    ataque = "B1"
    defensa = "C1"

    items = ",{"
    items += "'nombre': 'Proyección Magica',"
    items += "'tipo': 'Místico',"
    items += "'conocimiento': 'Conocida',"
    items += "'tamanio': 'Normal',"
    items += "'critPrincipal': 'Ene',"
    items += "'critSecundario': 'Pen',"
    items += "'entereza': '999',"
    items += "'rotura': '0',"
    items += "'presencia': '0',"
    items += "'turno': '" + cell(weaponBlock, turno) + "',"
    items += "'ataque': '" + cell(weaponBlock, ataque) + "',"
    items += "'defensa': '" + cell(weaponBlock, defensa) + "',"
    items += "'defensaTipo': 'Par',"
    items += "'danio': '0',"
    items += "'calidad': '-',"
    items += "'caracteristica': '-',"
    items += "'advertencia': '-',"
    items += "'municion': '-',"
    items += "'variable': 'true',"
    items += "'especial': '-'}"
    return items


def getArmours(sheet: pandas.DataFrame):
    items = ", 'armaduras': ["
    rng = range(sheet, "C12", "S15")

    firstComma = ""

    for row in rng.values:
        row = pandas.DataFrame(row).transpose()
        if not cell(row, "A1") == "nan":
            items += firstComma + "{ 'nombre': '" + cell(row, "A1") + "',"
            items += "'Localizacion': '" + cell(row, "D1") + "',"
            items += "'calidad': '" + cell(row, "F1") + "',"
            items += "'FIL': '" + cell(row, "G1") + "',"
            items += "'CON': '" + cell(row, "H1") + "',"
            items += "'PEN': '" + cell(row, "I1") + "',"
            items += "'CAL': '" + cell(row, "J1") + "',"
            items += "'ELE': '" + cell(row, "K1") + "',"
            items += "'FRI': '" + cell(row, "L1") + "',"
            items += "'ENE': '" + cell(row, "M1") + "',"
            items += "'Entereza': '" + cell(row, "N1") + "',"
            items += "'Presencia': '" + cell(row, "O1") + "',"
            items += "'RestMov': '" + cell(row, "P1") + "',"
            items += "'Enc': '" + cell(row, "Q1") + "'}"

            firstComma = ", "
    items += "]"
    return items


def getArmourData(combatSheet: pandas.DataFrame):
    items = ", 'armadura': {"
    items += "'restriccionMov': '" + cellOfDF(combatSheet, "E16") + "',"
    items += "'penNatural': '" + cellOfDF(combatSheet, "H17") + "',"
    items += "'requisito': '" + cellOfDF(combatSheet, "H16") + "',"
    items += "'penAccionFisica': '" + cellOfDF(combatSheet, "S16") + "',"
    items += "'penNaturalFinal': '" + cellOfDF(combatSheet, "S17") + "'"
    items += ", 'armaduraTotal': {"
    items += "'FIL': '" + cellOfDF(combatSheet, "I16") + "',"
    items += "'CON': '" + cellOfDF(combatSheet, "J16") + "',"
    items += "'PEN': '" + cellOfDF(combatSheet, "K16") + "',"
    items += "'CAL': '" + cellOfDF(combatSheet, "L16") + "',"
    items += "'ELE': '" + cellOfDF(combatSheet, "M16") + "',"
    items += "'FRI': '" + cellOfDF(combatSheet, "N16") + "',"
    items += "'ENE': '" + cellOfDF(combatSheet, "O16") + "'}"
    items += getArmours(combatSheet) + "}"
    return items


def getAllCombatData(combatSheet: pandas.DataFrame, pointsSheet: pandas.DataFrame, psychicSheet: pandas.DataFrame, mistycSheet: pandas.DataFrame):
    items = "'armas': ["
    items += getUnarmedCombatData(combatSheet)
    if hasPointsOnMistic(pointsSheet):
        items += getMagicProjectionAsWeapon(mistycSheet)
    if hasPointsOnPsichiq(pointsSheet):
        items += getPsychicProjectionAsWeapon(psychicSheet)
    blocksWeaponsBase = [["C27", "L32"], ["C34", "L39"], [
        "C41", "L46"], ["N27", "W32"], ["N34", "W39"], ["N41", "W46"]]
    blocksWeaponsRanged = [["C49", "L55"], [
        "C58", "L64"], ["N49", "W55"], ["N58", "W64"]]
    for block in blocksWeaponsBase:
        items += getBaseCombatData(range(combatSheet, block[0], block[1]))
    for block in blocksWeaponsRanged:
        items += getRangedCombatData(range(combatSheet, block[0], block[1]))

    items += "]"
    items += getArmourData(combatSheet)
    return items


def getBasicData(principalSheet: pandas.DataFrame):
    cansancio = "N16"
    puntosDeVida = "N11"
    regeneracion = "J11"
    movimiento = "J16"
    nombre = "K4"
    categoria = "K5"
    nivel = "O6"
    clase = "K7"
    acumDanio = "Y13"
    creadoConMagia = "Y14"
    gnosis = "AB13"
    natura = "AB14"

    items = "'datosElementales': {"
    items += "'cansancio': '" + cellOfDF(principalSheet, cansancio) + "',"
    items += "'puntosDeVida': '" + cellOfDF(principalSheet, puntosDeVida) + "',"
    items += "'regeneracion': '" + cellOfDF(principalSheet, regeneracion) + "',"
    items += "'nombre': '" + cellOfDF(principalSheet, nombre) + "',"
    items += "'categoria': '" + cellOfDF(principalSheet, categoria) + "',"
    items += "'nivel': '" + cellOfDF(principalSheet, nivel) + "',"
    items += "'clase': '" + cellOfDF(principalSheet, clase) + "',"
    items += "'acumDanio': '" + cellOfDF(principalSheet, acumDanio) + "',"
    items += "'creadoConMagia': '" + \
        cellOfDF(principalSheet, creadoConMagia) + "',"
    items += "'gnosis': '" + cellOfDF(principalSheet, gnosis) + "',"
    items += "'natura': '" + cellOfDF(principalSheet, natura) + "',"
    items += "'movimiento': '" + cellOfDF(principalSheet, movimiento) + "' }"
    return items


def getMisticBasicData(mistycSheet: pandas.DataFrame):
    regen = "J12"
    act = "L12"
    zeon = "K18"

    items = "'regen': '" + cellOfDF(mistycSheet, regen) + "',"
    items += "'act': '" + cellOfDF(mistycSheet, act) + "',"
    items += "'zeon': '" + cellOfDF(mistycSheet, zeon) + "'"
    return items


def basicKiData(kiSheet: pandas.DataFrame):
    maxAcu = "F24"
    genericAcu = "D24"

    items = "'acumulacionMax': '" + cellOfDF(kiSheet, maxAcu) + "',"
    items += "'acumulacionGenerica': '" + cellOfDF(kiSheet, genericAcu) + "'"
    return items


def readFile(name: str):
    t = time.process_time()

    # Read file
    dataframe = pandas.read_excel(
        name, ["Principal", "Combate", "Psíquicos", "Místicos", "PDs", "Elan", "Ki"])

    elapsed_time = time.process_time() - t

    print("all done at %.2f seconds" % elapsed_time)

    return dataframe


def readFileFromBase64(base64Excel: str):
    t = time.process_time()
    
    excel = base64.b64decode(base64Excel)

    # Read file
    dataframe = pandas.read_excel(
        excel, ["Principal", "Combate", "Psíquicos", "Místicos", "PDs", "Elan", "Ki"])

    elapsed_time = time.process_time() - t

    print("all done at %.2f seconds" % elapsed_time)

    return dataframe

def readFile(file):
    t = time.process_time()
    
    # Read file
    dataframe = pandas.read_excel(
        file, ["Principal", "Combate", "Psíquicos", "Místicos", "PDs", "Elan", "Ki"])

    elapsed_time = time.process_time() - t

    print("all done at %.2f seconds" % elapsed_time)

    return dataframe


def exportJson(dataframe: pandas.DataFrame):
    principalSheet = dataframe["Principal"]
    combatSheet = dataframe["Combate"]
    psychicSheet = dataframe["Psíquicos"]
    mistycSheet = dataframe["Místicos"]
    pointsSheet = dataframe["PDs"]
    kiSheet = dataframe["Ki"]
    elanSheet = dataframe["Elan"]

    json = "{"
    # Set Attributes Block
    json += range_to_json(principalSheet, "D11", "H18",
                          "A1", "D1", "Atributos")
    # Set Skills Block
    json += "," + range_to_json(principalSheet, "M22",
                                "Q77", "A1", "E1", "Habilidades")
    # Set BasicData Block
    json += "," + getBasicData(principalSheet)
    # Set Resistances Block
    json += "," + range_to_json(principalSheet, "D57",
                                "J62", "A1", "G1", "Resistencias")
    # Set Creature Powers Block
    json += "," + range_to_json(principalSheet, "AB12",
                                "AH23", "A1", "E1", "PoderesDeCriatura")
    # Set Elemental Habilities Block
    json += "," + range_to_json(principalSheet, "AD12",
                                "AH69", "A1", "D1", "HabilidadesElementales")
    # START Combat Section
    json += ", 'Combate': {"
    # Set Weapons Tables Block
    json += range_to_json(combatSheet, "AH11", "AL16",
                          "A1", "E1", "TablasDeArmas")
    # Set Combat Styles Block
    json += "," + range_to_json(combatSheet, "AB19",
                                "AQ28", "A1", "G1", "EstilosDeCombate")
    # Set ArsMagnus Block
    json += "," + range_to_json(combatSheet, "AB55",
                                "AQ64", "A1", "G1", "ArsMagnus")
    # Set Combat Styles Block
    json += "," + range_to_json(combatSheet, "AB31",
                                "AQ49", "A1", "E1", "ArtesMarciales")
    # Set Combat Block
    json += ", " + getAllCombatData(combatSheet,
                                    pointsSheet, psychicSheet, mistycSheet)
    # END Combat Section
    json += "}"
    # Set Ki Block
    json += ", 'Ki': {"
    json += range_to_json(kiSheet, "C12", "D23", "A1", "B1", "Acumulaciones")
    json += ", " + range_to_json(kiSheet, "C35",
                                 "F36", "A1", "D1", "Habilidades")
    json += ", " + range_to_json(kiSheet, "C12", "F23", "A1", "D1", "Maximos")
    json += ", " + basicKiData(kiSheet)
    json += "}"
    # Set Elan Block
    json += ", " + range_to_json(elanSheet, "C29", "G49", "D1", "E1", "Elan")
    # START Mistic Section
    if hasPointsOnMistic(pointsSheet):
        json += ", 'Misticos': {"
        json += getMisticBasicData(mistycSheet)
        json += ", " + range_to_json(mistycSheet,
                                     "C15", "H25", "A1", "F1", "Vias")
        json += ", " + range_to_json(mistycSheet,
                                     "C15", "H25", "A1", "C1", "SubVias")
        json += ", " + range_to_json(mistycSheet,
                                     "W53", "AB73", "A1", "F1", "Metamagia")
        json += ", " + range_to_json(mistycSheet,
                                     "Y12", "AC50", "E1", "A1", "Conjuros")
        json += ", " + range_to_json(mistycSheet,
                                     "AG12", "AK50", "E1", "A1", "Libres")
        json += "}"
    # END Mistic Section
    # START PSY Section
    if hasPointsOnPsichiq(pointsSheet):
        json += ", 'Psiquicos': {"
        json += range_to_json(psychicSheet, "C25", "Q36",
                              "A1", "D1", "Disciplinas")
        json += ", " + range_to_json(psychicSheet,
                                     "C39", "Q50", "A1", "D1", "Patrones")
        json += ", " + range_to_json(psychicSheet,
                                     "V11", "AB64", "A1", "G1", "Poderes")
        json += ", " + range_to_json(psychicSheet,
                                     "AD17", "AK62", "A1", "H1", "Innatos")
        json += "}"
    # END PSY Section
    # START PSY Section
    if hasPointsOnPsichiq(pointsSheet):
        json += ", 'Psiquicos': {"
        json += range_to_json(psychicSheet, "C25", "Q36",
                              "A1", "D1", "Disciplinas")
        json += ", " + range_to_json(psychicSheet,
                                     "C39", "Q50", "A1", "D1", "Patrones")
        json += ", " + range_to_json(psychicSheet,
                                     "V11", "AB64", "A1", "G1", "Poderes")
        json += ", " + range_to_json(psychicSheet,
                                     "AD17", "AK62", "A1", "H1", "Innatos")
        json += "}"
    # END PSY Section
    json += "}"

    return json