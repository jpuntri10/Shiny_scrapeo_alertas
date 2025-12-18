


# ==== LIMPIA ENTORNO ====
rm(list = ls())
gc()

# ==== LIBRERÍAS ====
# install.packages("readxl")   # descomenta si no lo tienes
# install.packages("openxlsx") # descomenta si no lo tienes
library(readxl)
library(openxlsx)

# ==== 1) LECTURA DEL EXCEL ====
# OPCIÓN A: archivo en el working directory (ajusta nombre si hace falta):
archivo <- "valor_primas.xlsx"
hoja    <- 1   # o el nombre exacto de tu hoja, p.ej. "Hoja1"

# OPCIÓN B: si prefieres seleccionar el archivo manualmente:
# archivo <- file.choose()
# hoja    <- 1

# Leer la hoja
df <- read_excel(archivo, sheet = hoja)

# Tomar estrictamente las 4 primeras columnas: A-D
if (ncol(df) < 4) stop("La hoja necesita al menos 4 columnas (A:D).")
df <- df[, 1:4]

# Poner nombres SIN espacios para trabajar sin problemas
names(df) <- c("TIPO_CLIENT", "ramo", "Dolares", "Valor_a_cambiar")

# ==== 2) TIPOS BÁSICOS ====
# 'ramo' como texto; los montos como numéricos (quitando comas si hubiera)
df$ramo             <- as.character(df$ramo)
df$Dolares          <- suppressWarnings(as.numeric(gsub(",", "", as.character(df$Dolares))))
df$Valor_a_cambiar  <- suppressWarnings(as.numeric(gsub(",", "", as.character(df$Valor_a_cambiar))))

# Quitar filas sin ramo
df <- df[!is.na(df$ramo) & df$ramo != "", ]

# ==== 3) UMBRAL POR FILA ====
# Usamos D (Valor_a_cambiar); si es NA en alguna fila, usamos round(C) de ESA fila.
umbral <- df$Valor_a_cambiar
na_idx <- is.na(umbral)
if (any(na_idx)) umbral[na_idx] <- round(df$Dolares[na_idx], 0)

# Verificación rápida en consola: DEBEN salir distintos (894, 1540, 1483, 2756, …)
cat("Primeros 12 pares (ramo, umbral):\n")
print(data.frame(ramo = head(df$ramo, 12), umbral = head(umbral, 12)))

# ==== 4) CONSTRUCCIÓN DE REGLAS (BASE R, TODO TEXTO) ====
n <- nrow(df)
if (n == 0) stop("No hay filas válidas para procesar.")

GRUPOS            <- rep(seq_len(n), each = 4)
Campo             <- rep(c("TIPO_PERSONA_BENEF", "TIPO_PERSONA_BENEF", "RAMO", "IMP_PRIMA"), times = n)
Operador          <- rep(c("IGUAL", "IGUAL", "IGUAL", "MAYOR"), times = n)
Operador_conexion <- rep(c("Y", "Y", "Y", "FIN"), times = n)
Marca_inhabil     <- rep("NO", 4 * n)

# Valor_Umbral: dos "S", luego RAMO (texto), luego IMP_PRIMA (texto para evitar choques de tipo)
Valor_Umbral <- unlist(lapply(seq_len(n), function(i) {
  c("S", "S", df$ramo[i], as.character(umbral[i]))
}))

reglas <- data.frame(
  GRUPOS, Campo, Operador, Valor_Umbral, Operador_conexion, Marca_inhabil,
  stringsAsFactors = FALSE
)

# Verificación: los IMP_PRIMA deben variar por grupo
cat("\nPrimeros 12 IMP_PRIMA (GRUPOS, Valor_Umbral):\n")
imp <- reglas[reglas$Campo == "IMP_PRIMA", c("GRUPOS", "Valor_Umbral")]
print(head(imp, 12))

# ==== 5) EXPORTAR A EXCEL ====
wb <- createWorkbook()
addWorksheet(wb, "Reglas")
writeData(wb, "Reglas", reglas)
setColWidths(wb, "Reglas", cols = 1:ncol(reglas), widths = "auto")
saveWorkbook(wb, "reglas_generadas.xlsx", overwrite = TRUE)

## Mensaje final (con cierre correcto)
