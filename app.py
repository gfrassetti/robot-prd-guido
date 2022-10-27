from gc import collect
import time
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import logging
import os
from tkinter import messagebox

path = "./FT_a_procesar"
contenido = os.listdir(path)
fichas = []

for ficha in contenido:
    if os.path.isfile(os.path.join(path, ficha)) and ficha.endswith(".xlsm"):
        fichas.append(ficha)
print(fichas)

logging.basicConfig(
    filename="app.txt",
    level=logging.INFO,
    format="%(asctime)s:%(levelname)s:%(message)s",
)

# read/write Excel

""" Download chromedriver """
# chrome://version/
# https://chromedriver.storage.googleapis.com/index.html
# pip install openpyxl


logging.info("...Iniciando...")

#'openpyxl.worksheet.merge.MergedCellRange', MergedCellRange


# driver_service = Service(executable_path="./selenium-driver/chromedriver.exe")
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window()
# URL
driver.get("http://grisinobftestfinal:8001/maer/")  # cambiar url aca


class Login:
    def __init__(self, user, password):
        self.user = user
        self.password = password

    def login(self):
        try:
            logging.info("Iniciando sesion..")
            login_user = WebDriverWait(driver, 10).until(
                expected_conditions.presence_of_element_located(
                    (By.ID, "ext-comp-1002")
                )
            )

            login_user.send_keys(self.user)

            password_user = WebDriverWait(driver, 10).until(
                expected_conditions.presence_of_element_located(
                    (By.ID, "ext-comp-1004")
                )
            )
            password_user.send_keys(self.password)

            login_btn = driver.find_element(By.ID, "ext-gen31")
            login_btn.click()
            logging.info("Entrando...")
        except (Exception) as error_excepction:
            logging.warning("Error: ", error_excepction)


class LoadFile:
    def __init__(self, fichas):
        self.fichas = fichas
        logging.info("Cargando file...")

    def loop(self, rango, lista):
        # Iterar por cod de color
        for cod in rango:
            for i in cod:
                if i.value != None:
                    lista.append(i.value)

    def split_cod_color(self, cod_color):
        return cod_color.split("-", 1)[1]

    def loop_cod_color(self, rango_cod_color, lista_cod_color, celda, rango_str, ws):
        # comprobar si la celda esta mergeada para elegir un color o varios
        merged_cell_ranges = ws.merged_cells.ranges
        merged_cell_ranges_list = list(map(str, merged_cell_ranges))
        merged_cells = []

        for merged_range in merged_cell_ranges_list:
            merged_cells.append(merged_range)

        if rango_str in merged_cells:
            print("Cell is Merged!")
            merged_cell_cod_color = self.split_cod_color(celda.value)
            lista_cod_color.append(merged_cell_cod_color)
            print("109: ", lista_cod_color)
            return True
        else:
            print("Cell is not merged")
            for cod in rango_cod_color:
                for i in cod:
                    if i.value != None:
                        single_cod_color = self.split_cod_color(i.value)
                        lista_cod_color.append(single_cod_color)
                        print("116: ", lista_cod_color)
            return False

    def comprobar_y_cargar(
        self,
        actions,
        descripcion_validacion,
        talles,
        lista_cod_color,
        cantidad_insumo_confeccion,
        insumo_confeccion,
        lista_colores,
        isCombined,
        agregar_insumo,
    ):
        if "TALLE" in descripcion_validacion:
            logging.info("Cargnado insumo por talle...")
            # Por cada talle cargar...}
            for talle in talles:
                time.sleep(1)
                logging.info("talles disponibles: ", talle)
                for idx, i in enumerate(lista_cod_color):
                    self.load_insumo_por_talle(
                        actions,
                        insumo_confeccion,
                        i,
                        cantidad_insumo_confeccion,
                        talle,
                        lista_colores[idx],
                        isCombined,
                        agregar_insumo,
                    )
        else:
            for idx, i in enumerate(lista_cod_color):
                self.load_insumo2(
                    actions,
                    insumo_confeccion,
                    i,
                    cantidad_insumo_confeccion,
                    lista_colores[idx],
                    isCombined,
                    agregar_insumo,
                )
            logging.info("Carga de insumo por talle finalizada...")

    def load_insumo(self, actions, insumo, cod_color_insumo, cantidad):
        if insumo != None:

            logging.info(f"Cargando el insumo {insumo}")
            time.sleep(2)
            actions.send_keys(insumo + "." + cod_color_insumo)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(2)
            actions.send_keys(cantidad)
            time.sleep(2)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            time.sleep(2)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
        else:
            actions.send_keys(Keys.ESCAPE)
            actions.perform()
            actions.send_keys(Keys.ESCAPE)
            actions.perform()
            logging.info(f"Carga de insumo {insumo} finalizada")

    def load_insumo2(
        self, actions, insumo, i, cantidad, color, isCombined, agregar_insumo
    ):
        if insumo != None:
            logging.info(f"Cargando inusmo por talle: {insumo}")
            agregar_insumo.click()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(1)
            print(f"Cargando cod de color: {i}")
            actions.send_keys(insumo + "." + i)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(2)
            actions.send_keys(cantidad)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            if isCombined:
                actions.send_keys(Keys.ENTER)
                print(f"Cargando color Todos")
                actions.perform()
            else:
                print("Color a cargar: ", color)
                actions.send_keys(color)
                actions.perform()
                time.sleep(2)
                actions.send_keys(Keys.ARROW_DOWN)
                actions.perform()
                time.sleep(2)
                actions.send_keys(Keys.ENTER)
                actions.perform()
            time.sleep(2)
            actions.send_keys("Todos")
            actions.perform()
            time.sleep(1)
            actions.send_keys(Keys.ESCAPE)
            actions.perform()

    def load_insumo_por_talle(
        self,
        actions,
        insumo,
        cod_color_insumo,
        cantidad,
        talle,
        color,
        isCombined,
        agregar_insumo,
    ):
        if insumo != None:
            logging.info(f"Cargando inusmo por talle: {insumo}")
            time.sleep(2)
            agregar_insumo.click()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(1)
            actions.send_keys(
                insumo
                + "."
                + cod_color_insumo
                + "."
                + talle
                + " - Etiqueta GRISINO Marca y Talle"
            )
            actions.perform()
            time.sleep(3)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(2)
            actions.send_keys(cantidad)
            actions.perform()
            time.sleep(3)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(3)
            if isCombined:
                actions.send_keys(Keys.ENTER)
                print(f"Cargando color Todos")
                actions.perform()
            else:
                print("Color a cargar: ", color)
                actions.send_keys(color)
                time.sleep(1)
                actions.perform()
                time.sleep(1)
                actions.send_keys(Keys.ENTER)
            time.sleep(2)
            if (
                talle == "6"
                or talle == "8"
                or talle == "10"
                or talle == "12"
                or talle == "14"
            ):
                actions.send_keys(talle + " (GRISINO C3-C4 NUMEROS)")
                actions.perform()
            else:
                actions.send_keys(talle + " (GRISINO C1-C2 NUMEROS)")
                actions.perform()
                time.sleep(1)
            time.sleep(2)

    def load_new(self):
        try:
            time.sleep(2)
            btn_produccion = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div[1]/div[2]/div/div/div/div[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/em/button",
                    )
                )
            )
            btn_produccion.click()
            time.sleep(2)
            btn_ficha_tecnica = driver.find_element(
                By.XPATH,
                "//span[contains(text(),'Fichas Técnicas')]",
            )
            btn_ficha_tecnica.click()

            btn_ficha_tecnica2 = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (By.ID, "menuPrincipalProducciónFichas TécnicasFichas Técnicas")
                )
            )
            btn_ficha_tecnica2.click()

            btn_maxim_ft = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (
                        By.CLASS_NAME,
                        "x-tool-maximize",
                    )
                )
            )
            btn_maxim_ft.click()
            btn_add_new = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (
                        By.XPATH,
                        "//button[contains(text(),'Agregar')]",
                    )
                )
            )
            time.sleep(2)
            btn_add_new.click()
            logging.info("Nueva ficha tecnica")
            logging.info("reading excel..")

            for index, ficha in enumerate(self.fichas):
                logging.info(f"Cargando ficha: {ficha}")
                wb = load_workbook(f"./FT_a_procesar/{ficha}", data_only=True)
                ws = wb.active

                time.sleep(60)

                input_coleccion = driver.find_element(
                    By.XPATH,
                    "//input[@id='ext-comp-1254']",
                )
                time.sleep(10)
                input_coleccion.click()
                actions = ActionChains(driver)
                coleccion_parte1 = ws["B1"].value
                coleccion_parte2 = ws["B2"].value
                coleccion_parte3 = ws["L5"].value
                time.sleep(2)
                coleccion = (
                    f"{coleccion_parte1} {coleccion_parte2[0]} - {coleccion_parte3}"
                )
                print("Coleccion: ", coleccion)
                time.sleep(2)
                actions.send_keys(coleccion)
                time.sleep(3)
                actions.send_keys(Keys.ENTER)
                time.sleep(2)
                actions.perform()
                time.sleep(3)
                actions.send_keys(Keys.TAB)
                actions.perform()
                time.sleep(3)
                producto = ws["B2"].value
                time.sleep(3)
                actions.send_keys(producto)
                time.sleep(2)
                actions.perform()
                actions.send_keys(Keys.ARROW_DOWN)
                actions.perform()
                actions.send_keys(Keys.ENTER)
                actions.perform()
                time.sleep(3)
                actions.send_keys(Keys.TAB)
                actions.perform()
                actions.send_keys(Keys.TAB)
                actions.perform()
                molde = ws["T2"].value
                actions.send_keys(molde)
                actions.perform()
                time.sleep(4)

                btn_add_rule = driver.find_element(
                    By.XPATH,
                    "//table[@id='ext-comp-1281']/tbody/tr[2]/td[2]/em/button",
                )
                time.sleep(5)
                logging.info("Agregando regla - telas corte")
                btn_add_rule.click()
                actions.send_keys("100 - CORTE")
                actions.perform()
                time.sleep(1)
                actions.send_keys(Keys.ENTER)
                actions.perform()
                time.sleep(1)
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                time.sleep(3)

                nueva_entrada = driver.find_element(
                    By.XPATH,
                    "//div[@id='ext-comp-1276']/div/div[2]/div/div/div[2]/div/div/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button",
                )
                time.sleep(5)
                nueva_entrada.click()
                time.sleep(2)

                if index == 0:
                    agregar_insumo = driver.find_element(
                        By.XPATH,
                        "//table[@id='ext-comp-1324']/tbody/tr[2]/td[2]/em/button",
                    )
                elif index == 1:
                    agregar_insumo = driver.find_element(
                        By.XPATH,
                        "//table[@id='ext-comp-1705']/tbody/tr[2]/td[2]/em/button",
                    )

                time.sleep(2)
                agregar_insumo.click()
                logging.info("Agregando insumos telas")
                time.sleep(1)
                actions.send_keys(Keys.TAB)
                actions.perform()

                insumo_1 = ws["I6"].value

                cod_color_inusmo = ws["L7"].value

                cod_color_insumo2 = ws["L9"].value

                # Cantidad
                cantidad_insumo_1 = str(ws["J6"].value)
                cantidad_insumo_2 = ws["J8"].value

                time.sleep(2)
                self.load_insumo(
                    actions,
                    insumo_1,
                    self.split_cod_color(cod_color_inusmo),
                    cantidad_insumo_1,
                )
                time.sleep(2)
                insumo_2 = ws["I8"].value
                insumo_3 = ws["I10"].value

                cod_color_insumo3 = ws["L11"].value

                cantidad_insumo_3 = ws["J10"].value
                insumo_4 = ws["I12"].value
                insumo_5 = ws["I14"].value
                insumo_6 = ws["I16"].value
                cod_color_insumo4 = ws["L13"].value

                cod_color_insumo5 = ws["L15"].value
                cod_color_insumo6 = ws["L17"].value

                cantidad_insumo_4 = str(ws["J12"].value)
                cantidad_insumo_5 = str(ws["J14"].value)
                cantidad_insumo_6 = str(ws["J16"].value)

                # Si insumo existe.. agregar otro
                # Se puede hacer una fx decoradora -----------------------------------------------------------
                if insumo_2 != None:
                    agregar_insumo.click()
                    actions.send_keys(Keys.TAB)
                    time.sleep(2)
                    actions.perform()
                    time.sleep(2)
                    self.load_insumo(
                        actions,
                        insumo_2,
                        self.split_cod_color(cod_color_insumo2),
                        cantidad_insumo_2,
                    )
                else:
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    logging.info("Carga de insumos terminada")

                time.sleep(2)

                if insumo_3 != None:
                    agregar_insumo.click()
                    actions.send_keys(Keys.TAB)
                    time.sleep(2)
                    actions.perform()
                    time.sleep(2)
                    self.load_insumo(
                        actions,
                        insumo_3,
                        self.split_cod_color(cod_color_insumo3),
                        cantidad_insumo_3,
                    )
                else:
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    logging.info("Carga de insumos terminada")

                time.sleep(2)

                if insumo_4 != None:
                    agregar_insumo.click()
                    actions.send_keys(Keys.TAB)
                    time.sleep(2)
                    actions.perform()
                    time.sleep(2)
                    self.load_insumo(
                        actions,
                        insumo_4,
                        self.split_cod_color(cod_color_insumo4),
                        cantidad_insumo_4,
                    )

                else:
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    logging.info("Carga de insumos terminada")

                time.sleep(2)

                if insumo_5 != None:
                    agregar_insumo.click()
                    actions.send_keys(Keys.TAB)
                    time.sleep(2)
                    actions.perform()
                    time.sleep(2)
                    self.load_insumo(
                        actions,
                        insumo_5,
                        self.split_cod_color(cod_color_insumo5),
                        cantidad_insumo_5,
                    )

                else:
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    logging.info("Carga de insumos terminada")

                time.sleep(2)

                if insumo_6 != None:
                    agregar_insumo.click()
                    actions.send_keys(Keys.TAB)
                    time.sleep(2)
                    actions.perform()
                    time.sleep(2)
                    self.load_insumo(
                        actions,
                        insumo_6,
                        self.split_cod_color(cod_color_insumo6),
                        cantidad_insumo_6,
                    )
                else:
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    logging.info("Carga de insumos terminada")

                # ----------------------------------------------     CONFECCION -----------------------------------------------------------------------------------------------
                logging.info("Comenzando carga de confeccion...")
                time.sleep(1)
                btn_add_rule.click()
                time.sleep(1)
                proceso_corte = WebDriverWait(driver, 35).until(
                    expected_conditions.presence_of_element_located(
                        (
                            By.XPATH,
                            "/html/body/div[4]/div[2]/div/div/div/div[1]/div[8]/div[2]/div[1]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div/input",
                        )
                    )
                )
                estampas_o_bordados = ws["B22"].value
                proceso_corte.send_keys("600 - PREPARACION P/ TALLER")
                time.sleep(1)
                actions.send_keys(Keys.ARROW_DOWN)
                actions.perform()
                actions.send_keys(Keys.ARROW_DOWN)
                actions.perform()
                actions.send_keys(Keys.ENTER)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

                time.sleep(3)

                # /html/body/div[4]/div[2]/div/div/div/div[1]/div[8]/div[2]/div[1]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button
                nueva_entrada_bordado = WebDriverWait(driver, 35).until(
                    expected_conditions.presence_of_element_located(
                        (
                            By.XPATH,
                            "/html/body/div[4]/div[2]/div/div/div/div[1]/div[8]/div[2]/div[1]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button",
                        )
                    )
                )
                time.sleep(3)
                nueva_entrada_bordado.click()
                time.sleep(1)
                # ------------------------------------------------------  confeccion ----------------------------------------------------------------------------
                insumo_confeccion_1 = ws["I22"].value
                insumo_confeccion_2 = ws["I26"].value
                insumo_confeccion_3 = ws["I28"].value
                insumo_confeccion_4 = ws["I30"].value

                isCombined1 = ws["L22"]
                isCombined2 = ws["L26"]
                isCombined3 = ws["L28"]

                cantidad_insumo_confeccion_1 = ws["J22"].value
                cantidad_insumo_confeccion_2 = ws["J26"].value
                cantidad_insumo_confeccion_3 = ws["J28"].value

                descripcion_validacion_1 = ws["B22"].value
                descripcion_validacion_2 = ws["B26"].value
                descripcion_validacion_3 = ws["B28"].value

                rango_cod_color_str = "L22:U22"
                rango_cod_color2_str = "L26:U26"
                rango_cod_color3_str = "L28:U28"

                rango_cod_color = ws["L22":"U22"]
                rango_cod_color2 = ws["L26":"U26"]
                rango_cod_color3 = ws["L28":"U28"]

                lista_cod_color_1 = []
                lista_cod_color_2 = []
                lista_cod_color_3 = []

                lista_colores = []
                rango_colores = ws["L4":"T4"]

                # siempre van a estar en estas celdas??

                talles_value = ws["P2"].value
                talles = [str(x) for x in talles_value.split(" - ")]
                talles2 = []

                talles_value_2 = ws["P3"].value
                """ Talles segundo cod de prod """
                if talles_value_2 != None:
                    talles2 = [str(x) for x in talles_value_2.split(" - ")]
                else:
                    pass

                # LOOPS
                self.loop(rango_colores, lista_colores)

                # MERGED CELLS
                combinado_1 = self.loop_cod_color(
                    rango_cod_color,
                    lista_cod_color_1,
                    isCombined1,
                    rango_cod_color_str,
                    ws,
                )

                combinado_2 = self.loop_cod_color(
                    rango_cod_color2,
                    lista_cod_color_2,
                    isCombined2,
                    rango_cod_color2_str,
                    ws,
                )

                combinado_3 = self.loop_cod_color(
                    rango_cod_color3,
                    lista_cod_color_3,
                    isCombined2,
                    rango_cod_color3_str,
                    ws,
                )

                # FX

                time.sleep(2)
                if insumo_confeccion_1 != None:
                    self.comprobar_y_cargar(
                        actions,
                        descripcion_validacion_1,
                        talles,
                        lista_cod_color_1,
                        cantidad_insumo_confeccion_1,
                        insumo_confeccion_1,
                        lista_colores,
                        combinado_1,
                        agregar_insumo,
                    )
                    logging.info(f"Carga de inusmo: {insumo_confeccion_1} terminada")

                time.sleep(2)

                if insumo_confeccion_2 != None:
                    self.comprobar_y_cargar(
                        actions,
                        descripcion_validacion_2,
                        talles,
                        lista_cod_color_2,
                        cantidad_insumo_confeccion_2,
                        insumo_confeccion_2,
                        lista_colores,
                        combinado_2,
                        agregar_insumo,
                    )
                    logging.info(f"Carga de inusmo: {insumo_confeccion_2} terminada")

                time.sleep(2)

                if insumo_confeccion_3 != None:
                    self.comprobar_y_cargar(
                        actions,
                        descripcion_validacion_3,
                        talles,
                        lista_cod_color_3,
                        cantidad_insumo_confeccion_3,
                        insumo_confeccion_3,
                        lista_colores,
                        combinado_3,
                        agregar_insumo,
                    )
                    logging.info(f"Carga de inusmo: {insumo_confeccion_3} terminada")

                time.sleep(2)
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

                # --------------------------------------------------- GUARDA PRIMERA PARTE -----------------------------------------------------
                time.sleep(2)
                btn_guardar = driver.find_element(
                    By.XPATH,
                    "/html/body/div[4]/div[2]/div/div/div/div[1]/div[8]/div[2]/div[1]/div/div/div/div/div/div[2]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/em/button",
                )
                time.sleep(2)
                btn_guardar.click()
                time.sleep(5)
                btn_si = driver.find_element(
                    By.XPATH, "//button[contains(text(),'Sí')]"
                )
                time.sleep(5)
                btn_si.click()
                btn_ok = driver.find_element(
                    By.XPATH, "//button[contains(text(),'OK')]"
                )
                time.sleep(5)
                btn_ok.click()
                logging.info("Ficha Guardada")
                time.sleep(5)
                btn_close = driver.find_element(
                    By.XPATH,
                    "//div[@id='ext-comp-1485']/div/div/div/div/div",
                )
                time.sleep(2)
                btn_close.click()
                time.sleep(2)
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

                """ ------------------------------------------------------------  SEGUNDO COD DE PRODUCTO  ---------------------------------------------------- """

                cod_art_2 = ws["B3"].value

                if cod_art_2 != None:
                    logging.info(f"Cargando segundo codigo de producto: {cod_art_2}")
                    input_coleccion2 = driver.find_element(
                        By.XPATH,
                        "//*[@id='ext-comp-1254']",
                    )
                    actions = ActionChains(driver)
                    coleccion = ws["G1"].value
                    time.sleep(3)
                    input_coleccion2.send_keys(coleccion)
                    time.sleep(3)
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    time.sleep(3)
                    actions.send_keys(Keys.TAB)
                    actions.perform()
                    time.sleep(3)
                    actions.send_keys(cod_art_2)
                    actions.send_keys(Keys.ARROW_DOWN)
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    time.sleep(3)
                    actions.send_keys(Keys.TAB)
                    actions.perform()
                    time.sleep(3)
                    actions.send_keys(Keys.TAB)
                    actions.perform()
                    molde = ws["T2"].value
                    actions.send_keys(molde)
                    actions.perform()

                    time.sleep(2)
                    btn_add_rule.click()
                    time.sleep(1)
                    actions.send_keys("100 - CORTE")
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()

                    time.sleep(2)
                    logging.info("Agregando entrada")
                    nueva_entrada2 = driver.find_element(
                        By.XPATH,
                        "//div[@id='ext-comp-1276']/div/div[2]/div/div/div[2]/div/div/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button",
                    )
                    nueva_entrada2.click()
                    time.sleep(2)

                    agregar_insumo2 = driver.find_element(
                        By.XPATH,
                        "//table[@id='ext-comp-1527']/tbody/tr[2]/td[2]/em/button",
                    )
                    agregar_insumo2.click()
                    logging.info("Cargando insumos...")
                    time.sleep(4)
                    actions.send_keys(Keys.TAB)
                    actions.perform()

                    insumo_1 = ws["I6"].value
                    color_inusmo = ws["L7"].value
                    color_insumo2 = ws["L9"].value
                    cantidad_insumo_1 = str(ws["K6"].value)
                    cantidad_insumo_2 = str(ws["K8"].value)

                    time.sleep(2)
                    self.load_insumo(actions, insumo_1, color_inusmo, cantidad_insumo_1)
                    time.sleep(2)
                    insumo_2 = ws["I8"].value
                    insumo_3 = ws["I10"].value
                    color_insumo3 = ws["L11"].value
                    cantidad_insumo_3 = ws["K10"].value
                    insumo_4 = ws["I12"].value
                    insumo_5 = ws["I14"].value
                    insumo_6 = ws["I16"].value
                    color_insumo4 = ws["N5"].value
                    # XTA004
                    color_insumo5 = ws["N5"].value
                    # XTD001
                    color_insumo6 = ws["N5"].value
                    cantidad_insumo_4 = str(ws["K12"].value)
                    cantidad_insumo_5 = str(ws["K14"].value)
                    cantidad_insumo_6 = str(ws["K16"].value)

                    # Si insumo existe.. agregar otro
                    # Se puede hacer una fx decoradora -----------------------------------------------------------
                    if insumo_2 != None:
                        agregar_insumo2.click()
                        actions.send_keys(Keys.TAB)
                        time.sleep(2)
                        actions.perform()
                        time.sleep(2)
                        self.load_insumo(
                            actions, insumo_2, cod_color_insumo2, cantidad_insumo_2
                        )
                    else:
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        logging.info("Carga de insumos terminada")

                    time.sleep(2)

                    if insumo_3 != None:
                        agregar_insumo2.click()
                        actions.send_keys(Keys.TAB)
                        time.sleep(2)
                        actions.perform()
                        time.sleep(2)
                        self.load_insumo(
                            actions, insumo_3, color_insumo3, cantidad_insumo_3
                        )
                    else:
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        logging.info(f"Carga de insumo {insumo_3} finalizada")

                    time.sleep(2)

                    if insumo_4 != None:
                        agregar_insumo2.click()
                        actions.send_keys(Keys.TAB)
                        time.sleep(2)
                        actions.perform()
                        time.sleep(2)
                        self.load_insumo(
                            actions, insumo_4, color_insumo4, cantidad_insumo_4
                        )

                    else:
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        logging.info("Carga de insumos terminada")

                    time.sleep(2)

                    if insumo_5 != None:
                        agregar_insumo2.click()
                        actions.send_keys(Keys.TAB)
                        time.sleep(2)
                        actions.perform()
                        time.sleep(2)
                        self.load_insumo(
                            actions, insumo_5, color_insumo5, cantidad_insumo_5
                        )

                    else:
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        logging.info("Carga de insumos terminada")

                    time.sleep(2)

                    if insumo_6 != None:
                        agregar_insumo2.click()
                        actions.send_keys(Keys.TAB)
                        time.sleep(2)
                        actions.perform()
                        time.sleep(2)
                        self.load_insumo(
                            actions, insumo_6, color_insumo6, cantidad_insumo_6
                        )
                    else:
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        actions.send_keys(Keys.ESCAPE)
                        actions.perform()
                        logging.info("Carga de insumos terminada")

                    time.sleep(4)

                    # ------------------------------------------------------  confeccion ----------------------------------------------------------------------------
                    logging.info("Comenzando carga de confeccion...")
                    time.sleep(1)
                    btn_add_rule.click()
                    time.sleep(1)
                    actions.send_keys(estampas_o_bordados)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.ARROW_DOWN)
                    actions.perform()
                    actions.send_keys(Keys.ARROW_DOWN)
                    actions.perform()
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()

                    time.sleep(3)

                    # /html/body/div[4]/div[2]/div/div/div/div[1]/div[9]/div[2]/div[1]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button
                    nueva_entrada_bordado2 = driver.find_element(
                        By.XPATH,
                        "/html/body/div[4]/div[2]/div/div/div/div[1]/div[8]/div[2]/div[1]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button",
                    )
                    time.sleep(5)
                    nueva_entrada_bordado2.click()
                    time.sleep(1)

                    insumo_confeccion_1 = ws["I26"].value
                    insumo_confeccion_2 = ws["I27"].value
                    insumo_confeccion_3 = ws["I28"].value

                    isCombined1 = ws["L26"]
                    isCombined2 = ws["L27"]
                    isCombined3 = ws["L28"]

                    cantidad_insumo_confeccion_1 = ws["J26"].value
                    cantidad_insumo_confeccion_2 = ws["J27"].value
                    cantidad_insumo_confeccion_3 = ws["J28"].value

                    descripcion_validacion_1 = ws["B26"].value
                    descripcion_validacion_2 = ws["B27"].value
                    descripcion_validacion_3 = ws["B28"].value

                    rango_cod_color_str = "L26:U26"
                    rango_cod_color2_str = "L27:U27"
                    rango_cod_color3_str = "L28:U28"

                    rango_cod_color = ws["L26":"U26"]
                    rango_cod_color2 = ws["L27":"U27"]
                    rango_cod_color3 = ws["L28":"U28"]

                    lista_cod_color_1 = []
                    lista_cod_color_2 = []
                    lista_cod_color_3 = []

                    lista_colores = []
                    rango_colores = ws["L4":"T4"]

                    # LOOPS
                    self.loop(rango_colores, lista_colores)

                    # MERGED CELLS
                    combinado_1 = self.loop_cod_color(
                        rango_cod_color,
                        lista_cod_color_1,
                        isCombined1,
                        rango_cod_color_str,
                        ws,
                    )

                    combinado_2 = self.loop_cod_color(
                        rango_cod_color2,
                        lista_cod_color_2,
                        isCombined2,
                        rango_cod_color2_str,
                        ws,
                    )

                    combinado_3 = self.loop_cod_color(
                        rango_cod_color3,
                        lista_cod_color_3,
                        isCombined3,
                        rango_cod_color3_str,
                        ws,
                    )

                    # FX
                    time.sleep(1)
                    if insumo_confeccion_1 != None:
                        self.comprobar_y_cargar(
                            actions,
                            descripcion_validacion_1,
                            talles2,
                            lista_cod_color_1,
                            cantidad_insumo_confeccion_1,
                            insumo_confeccion_1,
                            lista_colores,
                            combinado_1,
                            agregar_insumo2,
                        )
                        logging.info(
                            f"Carga de inusmo: {insumo_confeccion_1} terminada"
                        )

                    time.sleep(2)

                    if insumo_confeccion_2 != None:
                        self.comprobar_y_cargar(
                            actions,
                            descripcion_validacion_2,
                            talles2,
                            lista_cod_color_2,
                            cantidad_insumo_confeccion_2,
                            insumo_confeccion_2,
                            lista_colores,
                            combinado_2,
                            agregar_insumo2,
                        )
                        logging.info(
                            f"Carga de inusmo: {insumo_confeccion_2} terminada"
                        )

                    time.sleep(2)

                    if insumo_confeccion_3 != None:
                        self.comprobar_y_cargar(
                            actions,
                            descripcion_validacion_3,
                            talles2,
                            lista_cod_color_3,
                            cantidad_insumo_confeccion_3,
                            insumo_confeccion_3,
                            lista_colores,
                            combinado_3,
                            agregar_insumo2,
                        )
                        logging.info(
                            f"Carga de inusmo: {insumo_confeccion_3} terminada"
                        )

                    time.sleep(2)
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()
                    actions.send_keys(Keys.ESCAPE)
                    actions.perform()

                    # --------------------------------------------------- GUARDA SEGUNDA PARTE -----------------------------------------------------
                    time.sleep(3)
                    btn_guardar2 = driver.find_element(
                        By.XPATH,
                        "//table[@id='ext-comp-1304']/tbody/tr[2]/td[2]/em/button",
                    )
                    time.sleep(2)
                    btn_guardar2.click()
                    time.sleep(2)
                    btn_si2 = driver.find_element(
                        By.XPATH,
                        "//table[@id='ext-comp-1480']/tbody/tr[2]/td[2]/em/button",
                    )
                    time.sleep(3)
                    btn_si2.click()
                    logging.info("Ficha Guardada")
                    time.sleep(2)
                    btn_ok2 = driver.find_element(
                        By.XPATH,
                        "//button[contains(text(),'OK')]",
                    )
                    time.sleep(2)
                    btn_ok2.click()
                    time.sleep(2)
                    btn_close2 = driver.find_element(
                        By.XPATH, "//div[@id='ext-comp-1663']/div/div/div/div/div"
                    )
                    btn_close2.click()
                    time.sleep(2)
                    logging.info(f"Ficha: ${ficha} cargada exitosamente")
                else:
                    pass
            # -------------------------------------------------- --------------------------------------------------------------------------

        except (Exception) as error_excepction:
            logging.warning("Error: ", error_excepction)
            print(error_excepction)


# Credenciales
log = Login("RobotPRD", "Robot123")
log.login()


# Cargar Nueva ficha
load = LoadFile(fichas)
load.load_new()
