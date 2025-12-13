

library(shiny)
library(RSelenium)
library(readxl)

# --- Utilidades ---
selenium_ok <- function(host = "127.0.0.1", port = 4444) {
  ok <- FALSE
  try({
    drv <- remoteDriver(remoteServerAddr = host, port = port, browserName = "chrome", path = "/wd/hub")
    drv$open(); drv$close(); ok <- TRUE
  }, silent = TRUE)
  ok
}

wait_for <- function(remDr, by = "css selector", value, timeout_secs = 30, poll = 0.2) {
  t0 <- Sys.time()
  repeat {
    el <- try(remDr$findElement(using = by, value = value), silent = TRUE)
    if (!inherits(el, "try-error")) return(el)
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_secs) {
      stop(sprintf("No se encontró el elemento '%s' (%s) en %ds", value, by, timeout_secs))
    }
    Sys.sleep(poll)
  }
}

safe_click <- function(remDr, by, value) {
  el <- wait_for(remDr, by, value)
  remDr$executeScript("arguments[0].scrollIntoView(true);", list(el))
  el$clickElement(); Sys.sleep(0.8)
}

# --- Login ---
login_mytron <- function(remDr, user, pass) {
  remDr$navigate("http://mytron.mapfre.com.pe/login.xhtml"); Sys.sleep(1.0)
  wait_for(remDr, "id", "frmBody:usernameLoginValue")$sendKeysToElement(list(user))
  wait_for(remDr, "id", "frmBody:password")$sendKeysToElement(list(pass))
  safe_click(remDr, "xpath", "//span[normalize-space(.)='Iniciar Sesión']/ancestor::button[1]")
  Sys.sleep(2.0); TRUE
}

# --- Navegación del árbol ---
expandir_toggler <- function(remDr, label_text) {
  xp_label <- sprintf("//span[contains(@class,'ui-treenode-label') and normalize-space(.)='%s']", label_text)
  xp_toggler <- sprintf("%s/preceding-sibling::span[contains(@class,'ui-tree-toggler')]", xp_label)
  lab <- wait_for(remDr, "xpath", xp_label)
  tog <- try(remDr$findElement("xpath", xp_toggler), silent = TRUE)
  if (!inherits(tog, "try-error")) tog$clickElement() else lab$clickElement()
  Sys.sleep(0.8)
}
abrir_mantenimiento <- function(remDr) expandir_toggler(remDr, "MANTENIMIENTOS")
abrir_planeamiento_control <- function(remDr) expandir_toggler(remDr, "PLANEAMIENTO Y CONTROL")

abrir_mantenedor_umbral <- function(remDr) {
  enlace <- wait_for(remDr, "xpath", "//a[contains(@href,'ListarUmbralAlerta.xhtml')]")
  remDr$executeScript("arguments[0].scrollIntoView(true);", list(enlace))
  enlace$clickElement(); Sys.sleep(2.0)
}

abrir_alerta <- function(remDr, i) {
  id <- sprintf("frmBody:j_idt76:%d:modificarlinkUmbralAlerta", i - 1)
  safe_click(remDr, "id", id)
}

# --- Interacción de modal y tabla de grupos ---
click_nuevo_grupo <- function(remDr) {
  safe_click(remDr, "id", "frmBody:j_idt104")
}

abrir_menu <- function(remDr, base_id) {
  trigger_xpath <- sprintf("//select[@id,'%s_input']/following-sibling::div[contains(@class,'ui-selectonemenu-trigger')]", base_id)
  trigger <- try(wait_for(remDr, "xpath", trigger_xpath, timeout_secs = 4), silent = TRUE)
  if (!inherits(trigger, "try-error")) {
    trigger$clickElement()
  } else {
    label_id <- paste0(base_id, "_label")
    label <- try(wait_for(remDr, "id", label_id, timeout_secs = 3), silent = TRUE)
    if (!inherits(label, "try-error")) label$clickElement()
  }
  Sys.sleep(0.5)
}

seleccionar_opcion <- function(remDr, base_id, valor) {
  abrir_menu(remDr, base_id)
  xpath_opcion <- sprintf("//li[contains(@class,'ui-selectonemenu-item') and normalize-space(.)='%s']", valor)
  safe_click(remDr, "xpath", xpath_opcion)
}

llenar_fila <- function(remDr, fila) {
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputcampoValue", fila$Campo)
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputoperadorValue", fila$Operador)
  val <- wait_for(remDr, "id", "frmBodyModalEditAlertaControl:inputvalorUmbralValue")
  val$clearElement(); val$sendKeysToElement(list(as.character(fila$`Valor Umbral`)))
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputoperadorConexionValue", fila$`Operador conexión`)
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputestadoValue", fila$`Marca inhabilitado`)
}

guardar_condicion <- function(remDr) {
  safe_click(remDr, "id", "frmBodyModalEditAlertaControl:guardarUmbralAlertaItemButton")
}

guardar_grupo <- function(remDr) {
  safe_click(remDr, "xpath", "//span[normalize-space(.)='Guardar']/ancestor::button[1]")
}

leer_codigo_grupo <- function(remDr) {
  campo <- wait_for(remDr, "id", "frmBodyModalEditAlertaControl:inputgrupoValue")
  valor <- campo$getElementAttribute("value")[[1]]
  as.integer(valor)
}

# Usar SIEMPRE índice interno (grupo_codigo - 1) para acciones en la tabla
editar_grupo <- function(remDr, grupo_codigo) {
  grupo_index <- grupo_codigo - 1
  xpath <- sprintf("//img[contains(@src,'edit.png') and contains(@id,':%d:')]", grupo_index)
  safe_click(remDr, "xpath", xpath)
}

click_agregar_en_grupo <- function(remDr, grupo_codigo) {
  grupo_index <- grupo_codigo - 1
  xpath <- sprintf("//img[contains(@src,'add.png') and contains(@id,':%d:')]", grupo_index)
  btn <- wait_for(remDr, "xpath", xpath, timeout_secs = 30)
  remDr$executeScript("arguments[0].scrollIntoView(true);", list(btn))
  btn$clickElement(); Sys.sleep(1.0)
}

# --- UI ---
ui <- fluidPage(
  titlePanel("Carga de grupos desde Excel (MyTron)"),
  sidebarLayout(
    sidebarPanel(
      textInput("host", "Host Selenium", "127.0.0.1"),
      numericInput("port", "Puerto", 4444),
      textInput("user", "Usuario", Sys.getenv("MYTRON_USER")),
      passwordInput("pass", "Contraseña", Sys.getenv("MYTRON_PASS")),
      actionButton("check", "Probar Selenium"),
      actionButton("login", "Iniciar Sesión"),
      numericInput("alerta_num", "Número de alerta a abrir", 1),
      fileInput("archivo_excel", "Selecciona archivo Excel"),
      actionButton("cargar_excel", "Cargar grupos en alerta"),
      hr(),
      verbatimTextOutput("status")
    ),
    mainPanel(
      p("Carga cada grupo con todas sus condiciones desde el Excel, usando edición y botón ➕ con índice DOM (grupo_codigo - 1).")
    )
  )
)
llenar_grupos_desde_excel <- function(remDr, path_excel) {
  datos <- read_excel(path_excel)
  grupos <- unique(datos$GRUPOS)
  total <- length(grupos)
  
  for (g in grupos) {
    filas <- datos[datos$GRUPOS == g, ]
    cat(sprintf("🆕 Creando grupo para GRUPOS = %s con %d condiciones\n", as.character(g), nrow(filas)))
    
    # Crear grupo y guardar primera condición
    click_nuevo_grupo(remDr)
    llenar_fila(remDr, filas[1, ])
    guardar_condicion(remDr)
    guardar_grupo(remDr)
    
    Sys.sleep(2.0)
    grupo_codigo <- leer_codigo_grupo(remDr)
    grupo_index <- grupo_codigo - 1
    cat(sprintf("✅ Grupo visible con código %d (índice DOM %d)\n", grupo_codigo, grupo_index))
    
    # Validar aparición mediante el id de los botones con índice interno
    xpath_validacion <- sprintf("//img[contains(@id,':%d:')]", grupo_index)
    wait_for(remDr, "xpath", xpath_validacion, timeout_secs = 30)
    
    # Agregar condiciones restantes
    if (nrow(filas) > 1) {
      cat(sprintf("✏️ Editando grupo %d (índice DOM %d)\n", grupo_codigo, grupo_index))
      editar_grupo(remDr, grupo_codigo)
      for (i in 2:nrow(filas)) {
        cat(sprintf("➕ Usando índice DOM %d para grupo %d\n", grupo_index, grupo_codigo))
        click_agregar_en_grupo(remDr, grupo_codigo)
        cat(sprintf("➕ Agregando condición %d: %s %s %s\n",
                    i,
                    as.character(filas[i, ]$Campo),
                    as.character(filas[i, ]$Operador),
                    as.character(filas[i, ]$`Valor Umbral`)))
        llenar_fila(remDr, filas[i, ])
        guardar_condicion(remDr)
      }
    }
  }
  paste("✅ Grupos procesados:", total)
}

server <- function(input, output, session) {
  rv <- reactiveValues(driver = NULL)
  
  # Verificar conexión con Selenium
  observeEvent(input$check, {
    ok <- selenium_ok(input$host, input$port)
    output$status <- renderText(if (ok) "✅ Selenium OK." else "❌ Selenium NO responde.")
  })
  
  # Login en MyTron
  observeEvent(input$login, {
    req(input$host, input$port, input$user, input$pass)
    try(if (!is.null(rv$driver)) rv$driver$close(), silent = TRUE)
    msg <- tryCatch({
      remDr <- remoteDriver(remoteServerAddr = input$host, port = input$port,
                            browserName = "chrome", path = "/wd/hub")
      remDr$open(); rv$driver <- remDr
      ok <- login_mytron(remDr, input$user, input$pass)
      if (ok) "✅ Login exitoso." else "❌ Login falló."
    }, error = function(e) paste("❌ Error en login:", conditionMessage(e)))
    output$status <- renderText(msg)
  })
  
  # Cargar grupos desde Excel en alerta seleccionada
  observeEvent(input$cargar_excel, {
    req(rv$driver, input$archivo_excel, input$alerta_num)
    msg <- tryCatch({
      abrir_mantenimiento(rv$driver)
      abrir_planeamiento_control(rv$driver)
      abrir_mantenedor_umbral(rv$driver)
      abrir_alerta(rv$driver, input$alerta_num)
      llenar_grupos_desde_excel(rv$driver, input$archivo_excel$datapath)
    }, error = function(e) paste("❌ Error al cargar Excel:", conditionMessage(e)))
    output$status <- renderText(msg)
  })
  
  # Cerrar sesión y navegador al terminar
  session$onSessionEnded(function() {
    try(if (!is.null(rv$driver)) rv$driver$close(), silent = TRUE)
  })
}

# Lanzar la aplicación
shinyApp(ui, server)
