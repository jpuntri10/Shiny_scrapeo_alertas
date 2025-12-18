



#
# Carga de grupos desde Excel (MyTron) - RSelenium + Shiny
# Correcciones:
#  - Selección flexible del Campo (alias/contains)
#  - Rescate del ítem 1 si no quedó persistido tras guardar grupo
#  - Soporte para "+" dentro del modal (si existe); si no, usa flujo original
#  - Tipado de Valor Umbral: numérico solo para RAMO e IMP_PRIMA; texto para otros campos
#  - Reintentos anti-Stale y esperas dirigidas (jQuery.active)
# ------------------------------------------------------------

library(shiny)
library(RSelenium)
library(readxl)

# ========== Utilidades de espera y robustez ==========

wait_for <- function(remDr, by = "css selector", value, timeout_secs = 20, poll = 0.15) {
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

wait_ajax <- function(remDr, timeout_secs = 8, poll = 0.05) {
  t0 <- Sys.time()
  repeat {
    jq <- try(remDr$executeScript("return (window.jQuery ? jQuery.active : 0);"), silent = TRUE)
    active <- if (!inherits(jq, "try-error")) as.numeric(jq[[1]]) else 0
    if (active == 0) break
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_secs) break
    Sys.sleep(poll)
  }
}

try_js <- function(remDr, script, args = list()) {
  try(remDr$executeScript(script, args), silent = TRUE)
}

safe_click_retry <- function(remDr, using, value, tries = 3, post_ajax_wait = TRUE) {
  for (k in 1:tries) {
    el <- wait_for(remDr, using, value)
    ok <- try({
      remDr$executeScript("arguments[0].scrollIntoView({block:'center'});", list(el))
      el$clickElement()
      TRUE
    }, silent = TRUE)
    if (!inherits(ok, "try-error")) {
      if (post_ajax_wait) wait_ajax(remDr)
      return(invisible(TRUE))
    }
    if (k == tries) stop(sprintf("Fallo al hacer click en '%s' (%s) tras %d intentos", value, using, tries))
    Sys.sleep(0.15)
  }
}

safe_type_retry <- function(remDr, using, value, texto, tries = 3) {
  for (k in 1:tries) {
    el <- wait_for(remDr, using, value)
    ok <- try({
      el$clearElement()
      el$sendKeysToElement(list(texto))
      TRUE
    }, silent = TRUE)
    if (!inherits(ok, "try-error")) {
      try_js(remDr, "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", list(el))
      try_js(remDr, "arguments[0].blur();", list(el))
      wait_ajax(remDr, 3)
      return(invisible(TRUE))
    }
    if (k == tries) stop(sprintf("Fallo al escribir en '%s' (%s) tras %d intentos", value, using, tries))
    Sys.sleep(0.15)
  }
}

element_exists <- function(remDr, using, value, timeout_short = 2) {
  out <- try({
    el <- wait_for(remDr, using, value, timeout_secs = timeout_short)
    !is.null(el)
  }, silent = TRUE)
  !inherits(out, "try-error") && isTRUE(out)
}

# ========== Login y navegación ==========

login_mytron <- function(remDr, user, pass) {
  remDr$navigate("http://mytron.mapfre.com.pe/login.xhtml")
  Sys.sleep(0.4)
  wait_for(remDr, "id", "frmBody:usernameLoginValue")$sendKeysToElement(list(user))
  wait_for(remDr, "id", "frmBody:password")$sendKeysToElement(list(pass))
  safe_click_retry(remDr, "xpath", "//span[normalize-space(.)='Iniciar Sesión']/ancestor::button[1]")
  wait_ajax(remDr, 12)
  TRUE
}

expandir_toggler <- function(remDr, label_text) {
  xp_label   <- sprintf("//span[contains(@class,'ui-treenode-label') and normalize-space(.)='%s']", label_text)
  xp_toggler <- sprintf("%s/preceding-sibling::span[contains(@class,'ui-tree-toggler')]", xp_label)
  lab <- wait_for(remDr, "xpath", xp_label)
  tog <- try(remDr$findElement("xpath", xp_toggler), silent = TRUE)
  if (!inherits(tog, "try-error")) tog$clickElement() else lab$clickElement()
  wait_ajax(remDr)
}
abrir_mantenimiento        <- function(remDr) expandir_toggler(remDr, "MANTENIMIENTOS")
abrir_planeamiento_control <- function(remDr) expandir_toggler(remDr, "PLANEAMIENTO Y CONTROL")

abrir_mantenedor_umbral <- function(remDr) {
  enlace <- wait_for(remDr, "xpath", "//a[contains(@href,'ListarUmbralAlerta.xhtml')]")
  remDr$executeScript("arguments[0].scrollIntoView({block:'center'});", list(enlace))
  enlace$clickElement()
  wait_ajax(remDr, 10)
}

abrir_alerta <- function(remDr, i) {
  id <- sprintf("frmBody:j_idt76:%d:modificarlinkUmbralAlerta", i - 1)
  safe_click_retry(remDr, "id", id)
}

# ========== SelectOneMenu helpers (flexibles) ==========

abrir_menu <- function(remDr, base_id) {
  trigger_xpath <- sprintf("//div[@id='%s']/span[contains(@class,'ui-selectonemenu-trigger')]", base_id)
  trigger <- try(wait_for(remDr, "xpath", trigger_xpath, timeout_secs = 8), silent = TRUE)
  if (!inherits(trigger, "try-error")) {
    try(trigger$clickElement(), silent = TRUE)
  } else {
    label_id <- paste0(base_id, "_label")
    label <- try(wait_for(remDr, "id", label_id, timeout_secs = 8), silent = TRUE)
    if (!inherits(label, "try-error")) try(label$clickElement(), silent = TRUE)
  }
  wait_ajax(remDr, 6)
}

# alias para campos con label "abreviado" en la UI
campo_alias <- function(valor) {
  v <- toupper(trimws(as.character(valor)))
  # Ajusta/añade alias reales de tu UI aquí:
  # He visto "TIPO_PERSC" en tus capturas, para Excel "TIPO_PERSONA_BENEF"
  if (v %in% c("TIPO_PERSONA_BENEF", "TIPO_PERSONA_BENEFICIARIO", "TIPO_PERS_BENEF", "TIPO_PERSC")) {
    return("TIPO_PERSC")  # <-- cambia por el label EXACTO que muestra tu UI si es distinto
  }
  v
}

# selección flexible: exacto -> alias -> contains
seleccionar_opcion_flexible <- function(remDr, base_id, valor) {
  # 1) intento exacto con el valor dado
  abrir_menu(remDr, base_id)
  xp_exact <- sprintf("//ul[contains(@class,'ui-selectonemenu-items')]//li[contains(@class,'ui-selectonemenu-item') and normalize-space(.)='%s']", valor)
  ok <- try({ safe_click_retry(remDr, "xpath", xp_exact); TRUE }, silent = TRUE)
  if (!inherits(ok, "try-error")) return(invisible(TRUE))
  
  # 2) intento exacto con alias (si lo hay)
  val_alias <- campo_alias(valor)
  if (val_alias != toupper(trimws(as.character(valor)))) {
    abrir_menu(remDr, base_id)
    xp_alias <- sprintf("//ul[contains(@class,'ui-selectonemenu-items')]//li[contains(@class,'ui-selectonemenu-item') and normalize-space(.)='%s']", val_alias)
    ok2 <- try({ safe_click_retry(remDr, "xpath", xp_alias); TRUE }, silent = TRUE)
    if (!inherits(ok2, "try-error")) return(invisible(TRUE))
  }
  
  # 3) contains por fragmento significativo (primeros 8-12 chars)
  frag <- substr(toupper(trimws(as.character(valor))), 1, max(8, min(nchar(valor), 12)))
  abrir_menu(remDr, base_id)
  xp_contains <- sprintf("//ul[contains(@class,'ui-selectonemenu-items')]//li[contains(@class,'ui-selectonemenu-item') and contains(normalize-space(.), '%s')]", frag)
  safe_click_retry(remDr, "xpath", xp_contains)
}

seleccionar_opcion <- function(remDr, base_id, valor) {
  abrir_menu(remDr, base_id)
  xpath_opcion <- sprintf("//ul[contains(@class,'ui-selectonemenu-items')]//li[contains(@class,'ui-selectonemenu-item') and normalize-space(.)='%s']", valor)
  safe_click_retry(remDr, "xpath", xpath_opcion)
}

# ========== Normalizadores / validadores ==========

norm_si_no <- function(x) {
  x <- toupper(trimws(as.character(x)))
  if (x %in% c("SI","S","Y")) "SI" else "NO"
}
norm_operador_conexion <- function(x) {
  x <- toupper(trimws(as.character(x)))
  if (x %in% c("Y","FIN")) x else "Y"
}
norm_operador <- function(x) toupper(trimws(as.character(x)))
is_campo_numerico <- function(campo) toupper(trimws(as.character(campo))) %in% c("RAMO","IMP_PRIMA")
coerce_valor_umbral_num <- function(x) {
  x_chr <- trimws(as.character(x))
  x_chr <- gsub(",", ".", x_chr)
  x_chr <- gsub("[^0-9.\\-]", "", x_chr)
  if (nchar(x_chr) == 0) "" else x_chr
}

# ========== Modal: confirmar render de filas ==========

count_modal_items <- function(remDr) {
  js <- "
    var m = document.querySelector(\"div[id*='frmBodyModalEditAlertaControl']\");
    if(!m) return -1;
    var rows = m.querySelectorAll(\"table.ui-datatable-data > tr:not(.ui-datatable-empty-message)\");
    return rows ? rows.length : 0;
  "
  out <- try(remDr$executeScript(js), silent = TRUE)
  if (inherits(out, "try-error") || is.null(out)) -1 else as.integer(out[[1]])
}

wait_modal_items_at_least <- function(remDr, n_min = 1, timeout_secs = 10, poll = 0.1) {
  t0 <- Sys.time()
  repeat {
    k <- count_modal_items(remDr)
    if (!is.na(k) && k >= n_min) return(TRUE)
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_secs) break
    Sys.sleep(poll)
  }
  FALSE
}

# Detecta si existe un "+" dentro del modal actual
modal_tiene_boton_agregar <- function(remDr) {
  element_exists(remDr, "xpath", "//div[contains(@id,'frmBodyModalEditAlertaControl')]//img[contains(@src,'add.png')]", timeout_short = 2)
}

modal_click_agregar <- function(remDr) {
  safe_click_retry(remDr, "xpath", "//div[contains(@id,'frmBodyModalEditAlertaControl')]//img[contains(@src,'add.png')]")
  wait_ajax(remDr, 6)
}

# ========== Interacciones del modal y la tabla ==========

click_nuevo_grupo <- function(remDr) {
  safe_click_retry(remDr, "id", "frmBody:j_idt104")
}

llenar_fila <- function(remDr, fila) {
  campo_val <- toupper(trimws(as.character(fila$Campo)))
  # Campo con selección flexible (alias/contains)
  seleccionar_opcion_flexible(remDr, "frmBodyModalEditAlertaControl:inputcampoValue", campo_val)
  
  # Operador (exacto)
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputoperadorValue", norm_operador(fila$Operador))
  
  # Valor Umbral (numérico solo para RAMO/IMP_PRIMA)
  valor_escrito <- if (is_campo_numerico(campo_val)) coerce_valor_umbral_num(fila$`Valor Umbral`) else trimws(as.character(fila$`Valor Umbral`))
  safe_type_retry(remDr, "id", "frmBodyModalEditAlertaControl:inputvalorUmbralValue", valor_escrito)
  
  # Operador conexión y estado
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputoperadorConexionValue", norm_operador_conexion(fila$`Operador conexión`))
  seleccionar_opcion(remDr, "frmBodyModalEditAlertaControl:inputestadoValue",           norm_si_no(fila$`Marca inhabilitado`))
}

guardar_condicion <- function(remDr) {
  safe_click_retry(remDr, "id", "frmBodyModalEditAlertaControl:guardarUmbralAlertaItemButton")
  wait_ajax(remDr, 8)
}

# Guardar grupo con múltiples fallbacks
guardar_grupo <- function(remDr) {
  # 1) Por ID típico
  btn <- try(remDr$findElement("id", "frmBodyModalEditAlertaControl:j_idt183"), silent = TRUE)
  if (!inherits(btn, "try-error")) {
    try_js(remDr, "arguments[0].scrollIntoView({block:'center'});", list(btn))
    try(btn$clickElement(), silent = TRUE)
    wait_ajax(remDr, 8)
    return(invisible(TRUE))
  }
  # 2) Por texto dentro del modal, excluyendo guardar item
  xp <- "//div[contains(@id,'frmBodyModalEditAlertaControl')]//button[.//span[normalize-space(.)='Guardar'] and not(contains(@id,'guardarUmbralAlertaItemButton'))]"
  ok <- try({ safe_click_retry(remDr, "xpath", xp); TRUE }, silent = TRUE)
  if (!inherits(ok, "try-error")) { wait_ajax(remDr, 8); return(invisible(TRUE)) }
  # 3) Botón anterior al 'Cerrar'
  xp2 <- "//div[contains(@id,'frmBodyModalEditAlertaControl')]//span[normalize-space(.)='Cerrar']/ancestor::button[1]/preceding-sibling::button[1]"
  ok2 <- try({ safe_click_retry(remDr, "xpath", xp2); TRUE }, silent = TRUE)
  if (!inherits(ok2, "try-error")) { wait_ajax(remDr, 8); return(invisible(TRUE)) }
  # 4) Fallback JS
  try_js(remDr, "
    var m=document.querySelector(\"div[id*='frmBodyModalEditAlertaControl']\");
    if(m){
      var btns=m.querySelectorAll('button');
      for(var i=0;i<btns.length;i++){
        var t=(btns[i].textContent||'').trim().toUpperCase();
        if(t==='GUARDAR'){ btns[i].click(); break; }
      }
    }
  ")
  wait_ajax(remDr, 6)
  invisible(TRUE)
}

click_cerrar_modal_grupo <- function(remDr) {
  # 1) Por ID
  btn <- try(remDr$findElement("id", "frmBodyModalEditAlertaControl:j_idt184"), silent = TRUE)
  if (!inherits(btn, "try-error")) {
    try_js(remDr, "arguments[0].scrollIntoView({block:'center'});", list(btn))
    try(btn$clickElement(), silent = TRUE)
    wait_ajax(remDr, 4)
    return(invisible(TRUE))
  }
  # 2) Por texto
  xp <- "//div[contains(@id,'frmBodyModalEditAlertaControl')]//span[normalize-space(.)='Cerrar']/ancestor::button[1]"
  ok <- try({ safe_click_retry(remDr, "xpath", xp); TRUE }, silent = TRUE)
  if (!inherits(ok, "try-error")) { wait_ajax(remDr, 4); return(invisible(TRUE)) }
  # 3) Fallback JS
  try_js(remDr, "try{PF('ModalEditAlertaControl').hide();}catch(e){}")
  wait_ajax(remDr, 3)
  invisible(TRUE)
}

leer_codigo_grupo <- function(remDr) {
  campo <- wait_for(remDr, "id", "frmBodyModalEditAlertaControl:inputgrupoValue")
  valor <- campo$getElementAttribute("value")[[1]]
  as.integer(valor)
}

editar_grupo <- function(remDr, grupo_codigo) {
  grupo_index <- grupo_codigo - 1
  xpath <- sprintf("//img[contains(@src,'edit.png') and contains(@id,':%d:')]", grupo_index)
  safe_click_retry(remDr, "xpath", xpath)
}

click_agregar_en_grupo <- function(remDr, grupo_codigo) {
  grupo_index <- grupo_codigo - 1
  xpath <- sprintf("//img[contains(@src,'add.png') and contains(@id,':%d:')]", grupo_index)
  safe_click_retry(remDr, "xpath", xpath)
  wait_ajax(remDr, 6)
}

# ========== Proceso principal: Excel -> UI ==========

llenar_grupos_desde_excel <- function(remDr, path_excel) {
  datos <- read_excel(path_excel)
  names(datos) <- trimws(names(datos))
  requeridas <- c("GRUPOS","Campo","Operador","Valor Umbral","Operador conexión","Marca inhabilitado")
  if (!all(requeridas %in% names(datos))) {
    stop(sprintf("El Excel debe tener las columnas: %s", paste(requeridas, collapse = ", ")))
  }
  
  grupos <- unique(datos$GRUPOS)
  total  <- length(grupos)
  
  for (g in grupos) {
    filas <- datos[datos$GRUPOS == g, , drop = FALSE]
    cat(sprintf("🆕 Grupo '%s' con %d condición(es)\n", as.character(g), nrow(filas)))
    
    # 1) Nuevo grupo
    click_nuevo_grupo(remDr)
    
    # 2) PRIMERA condición
    llenar_fila(remDr, filas[1, ])
    guardar_condicion(remDr)
    
    # Confirmar visual en el modal
    ok_item1 <- wait_modal_items_at_least(remDr, n_min = 1, timeout_secs = 10)
    if (!ok_item1) {
      cat("⚠️ La 1ra fila aún no aparece en el modal; esperando un poco más…\n")
      wait_ajax(remDr, 6)
    }
    
    # ¿Podemos añadir más desde el propio modal?
    if (nrow(filas) > 1 && modal_tiene_boton_agregar(remDr)) {
      # Añadir filas 2..n SIN cerrar el modal ni guardar grupo
      for (i in 2:nrow(filas)) {
        modal_click_agregar(remDr)
        llenar_fila(remDr, filas[i, ])
        guardar_condicion(remDr)
        wait_modal_items_at_least(remDr, n_min = i, timeout_secs = 8)
      }
      # Cerrar modal (flujo del usuario)
      click_cerrar_modal_grupo(remDr)
      cat("🔒 Editor cerrado (vía modal '+').\n")
    } else {
      # Flujo original: guardar grupo -> editar -> '+'
      guardar_grupo(remDr)
      grupo_codigo <- leer_codigo_grupo(remDr)
      grupo_index  <- grupo_codigo - 1
      cat(sprintf("✅ Grupo creado: código %d (índice DOM %d)\n", grupo_codigo, grupo_index))
      
      # Validar que el grupo tenga al menos 1 ítem al reabrir
      editar_grupo(remDr, grupo_codigo)
      k <- count_modal_items(remDr)
      if (is.na(k) || k < 1) {
        # REPARACIÓN: volver a cargar la primera fila aquí
        cat("🛠  Reparando: primer ítem no quedó persistido; reingresando fila 1...\n")
        llenar_fila(remDr, filas[1, ])
        guardar_condicion(remDr)
        wait_modal_items_at_least(remDr, n_min = 1, timeout_secs = 8)
      }
      
      # Agregar condiciones restantes con '+'
      if (nrow(filas) > 1) {
        for (i in 2:nrow(filas)) {
          click_agregar_en_grupo(remDr, grupo_codigo)
          llenar_fila(remDr, filas[i, ])
          guardar_condicion(remDr)
          wait_modal_items_at_least(remDr, n_min = i, timeout_secs = 8)
        }
      }
      click_cerrar_modal_grupo(remDr)
      cat("🔒 Editor cerrado (flujo original).\n")
    }
  }
  
  paste("✅ Grupos procesados:", total)
}

# ========== Shiny UI ==========

ui <- fluidPage(
  titlePanel("Carga de grupos desde Excel (MyTron)"),
  sidebarLayout(
    sidebarPanel(
      textInput("host", "Host Selenium", "127.0.0.1"),
      numericInput("port", "Puerto", 4444),
      textInput("user", "Usuario", Sys.getenv("MYTRON_USER")),
      passwordInput("pass", "Contraseña", Sys.getenv("MYTRON_PASS")),
      actionButton("login", "Iniciar Sesión"),
      numericInput("alerta_num", "Número de alerta a abrir", 1),
      fileInput("archivo_excel", "Selecciona archivo Excel"),
      actionButton("cargar_excel", "Cargar grupos en alerta"),
      hr(),
      verbatimTextOutput("status")
    ),
    mainPanel(
      p("Excel con columnas: GRUPOS, Campo, Operador, Valor Umbral, Operador conexión, Marca inhabilitado."),
      p("Reglas: RAMO e IMP_PRIMA -> Valor Umbral numérico; otros campos -> texto (p. ej., 'S').")
    )
  )
)

# ========== Shiny Server ==========

server <- function(input, output, session) {
  rv <- reactiveValues(driver = NULL)
  
  observeEvent(input$login, {
    req(input$host, input$port, input$user, input$pass)
    try(if (!is.null(rv$driver)) rv$driver$close(), silent = TRUE)
    
    msg <- tryCatch({
      remDr <- remoteDriver(remoteServerAddr = input$host, port = input$port,
                            browserName = "chrome", path = "/wd/hub")
      remDr$open()
      rv$driver <- remDr
      ok <- login_mytron(remDr, input$user, input$pass)
      if (ok) "✅ Login exitoso." else "❌ Login falló."
    }, error = function(e) paste("❌ Error en login:", conditionMessage(e)))
    
    output$status <- renderText(msg)
  })
  
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
  
  session$onSessionEnded(function() {
    try(if (!is.null(rv$driver)) rv$driver$close(), silent = TRUE)
  })
}

# ========== Lanzar app ==========

# Lanzar la aplicación
shinyApp(ui, server)