library(shiny)
library(bslib)
library(dplyr)
library(ggplot2)
library(ggiraph)
library(arrow)
library(sf)
library(shinycssloaders)
library(shinydisconnect)
library(prettymapr)

# NUEVO
library(readxl)
library(tidyr)
library(openxlsx)

# cache en disco
shinyOptions(cache = cachem::cache_disk("cache"))

cut_comunas <- readRDS("cut_comunas_all.rds")

options(
  spinner.color = "#3C533C",
  spinner.size = 1,
  spinner.type = 8
)

ui <- page_fluid(
  theme = bs_theme(
    bg = "#181818",
    fg = "#FFFFFF",
    primary = "#3C533C",
    base_font = font_google("Inter"),
    font_scale = .9
  ),
  
  includeCSS("styles.css"),
  
  disconnectMessage(
    text = "La conexión se perdió. Por favor, recarga la página.",
    refresh = "Recargar",
    background = "#181818",
    colour = "#FFFFFF",
    refreshColour = "#3C533C"
  ),
  
  title = "Plataforma para LdB MH - CENSO 2024",
  
  div(
    style = "margin-top: 12px;",
    
    layout_columns(
      col_widths = c(4, 8),
      
      div(
        h3("Plataforma para LdB MH - CENSO 2024"),
        h4("Visor de Entidades y Manzanas"),
        p("El siguiente mapa muestra unidades territoriales por comuna. Haz clic en ellas para agregar el código MANZENT que será utilizado para conseguir los datos.")
      ),
      
      div(
        layout_columns(
          selectInput(
            "region", "Región",
            choices = cut_comunas |>
              distinct(REGION, COD_REGION) |>
              tibble::deframe(),
            selected = c("Metropolitana De Santiago" = 13)
          ),
          
          div(
            selectInput("comuna", "Comuna", choices = NULL, width = "100%"),
            div(
              class = "azar",
              actionLink("azar_comuna", "Comuna aleatoria")
            )
          )
        )
      )
    ),
    
    div(
      style = "max-width: 600px; margin: auto; margin-top: -6px;",
      
      layout_columns(
        col_widths = c(6, 6),
        div(
          h4(textOutput("titulo_comuna")),
          h5(textOutput("titulo_region")),
          div(
            style = "display: flex; gap: 4px;",
            span("Color:"),
            span("AREA_C (URBANO/RURAL)", class = "id_variable")
          )
        ),
        div(
          style = "display: flex; flex-direction: column; justify-content: flex-end; height: 100%;",
          em(
            class = "explicacion",
            "Para hacer zoom, presionar ícono de lupa",
            img(src = "lupa_a.png", height = "20px"),
            "y hacer scroll sobre el mapa, o presionar la segunda lupa",
            img(src = "lupa_b.png", height = "20px"),
            "y seleccionar el área."
          )
        )
      ),
      
      div(
        style = "position: relative;",
        
        girafeOutput("mapa_interactivo", height = "600px") |>
          withSpinner(proxy.height = 400),
        
        div(
          style = "position: relative; z-index: 9999; margin-top: 10px; padding-bottom: 8px;",
          actionButton(
            "limpiar_seleccion",
            "Limpiar selección",
            class = "btn btn-sm btn-custom-clear"
          )
        ),
        
        tableOutput("click_table")
      ),
      
      tags$hr(style = "margin: 16px 0; border-color: #FFF;"),
      
      h4("Tablas desde Excel"),
      
      fileInput(
        "excel_codigos",
        "Sube Excel con MANZENT (y GRUPO)",
        accept = c(".xlsx", ".xls")
      ),
      
      div(
        style = "display:flex; gap:12px; flex-wrap:wrap; margin-top:6px;",
        tags$a(
          href = "https://drive.google.com/uc?export=download&id=1XYvVLuc2b0pry15ThkUyJ-mgutGITojn",
          "Descargar plantilla de códigos",
          target = "_blank",
          style = "display:inline-block;"
        )
      ),
      
      div(
        style="margin-top:10px; display:flex; gap:10px; flex-wrap:wrap;",
        actionButton("generar_tablas", "Generar tablas", class = "btn btn-success btn-sm"),
        downloadButton("descargar_tablas", "Descargar Excel", class = "btn btn-sm btn-custom-download")
      ),
      
      uiOutput("estado_tablas")
    )
  )
)

server <- function(input, output, session) {
  
  
  # =========================
  # CONFIG: versión Entidades/Manzanas
  # =========================
  TERRITORIO_PARQUET <- "base/Censo2024_Entidades_Manzanas.parquet"
  ID_COL <- "MANZENT"  # reemplaza ID_LOCALIDAD
  ID_LABEL <- "MANZENT"
  
  # --------------------------
  # helpers TABLAS desde Excel
  # --------------------------
  build_tablas_from_excel <- function(path_excel, sheet = "tablas") {
    
    cfg <- readxl::read_excel(path_excel, sheet = sheet) |>
      rename_with(toupper) |>
      mutate(
        TABLA = as.character(TABLA),
        COL   = as.character(COL),
        LABEL = as.character(LABEL),
        ORDEN = as.integer(ORDEN)
      ) |>
      filter(!is.na(TABLA), !is.na(COL), !is.na(LABEL)) |>
      arrange(TABLA, ORDEN)
    
    validate(need(all(c("TABLA","ORDEN","COL","LABEL") %in% names(cfg)),
                  "Config TABLAS: faltan columnas TABLA/ORDEN/COL/LABEL."))
    
    split_cfg <- split(cfg, cfg$TABLA)
    
    tablas <- lapply(split_cfg, function(d) {
      list(
        cols = d$COL,
        labels = d$LABEL
      )
    })
    
    tablas
  }
  
  make_freq_table <- function(df, cols, labels, titulo) {
    x <- df[, cols, drop = FALSE]
    x[] <- lapply(x, function(v) suppressWarnings(as.numeric(v)))
    
    n <- colSums(x, na.rm = TRUE)
    total <- sum(n, na.rm = TRUE)
    pct <- if (total > 0) round(100 * n / total, 2) else rep(0, length(n))
    
    out <- data.frame(
      Categoria = labels,
      n = as.integer(n),
      pct = paste0(formatC(pct, format = "f", digits = 2), "%"),
      stringsAsFactors = FALSE
    )
    
    out <- rbind(out, data.frame(Categoria = "Total", n = as.integer(total), pct = "100.00%"))
    attr(out, "titulo") <- titulo
    out
  }
  
  generar_tablas_baseline <- function(df, TABLAS_obj) {
    
    grupos <- split(df, df$GRUPO)
    resultado <- list()
    
    for (g in names(grupos)) {
      
      df_g <- grupos[[g]]
      tablas_g <- list()
      
      for (nombre_tabla in names(TABLAS_obj)) {
        
        cfg <- TABLAS_obj[[nombre_tabla]]
        
        cols_ok <- intersect(cfg$cols, names(df_g))
        if (length(cols_ok) == 0) next
        
        labels_ok <- cfg$labels[match(cols_ok, cfg$cols)]
        
        tablas_g[[nombre_tabla]] <- make_freq_table(
          df_g,
          cols = cols_ok,
          labels = labels_ok,
          titulo = nombre_tabla
        )
      }
      
      resultado[[g]] <- tablas_g
    }
    
    resultado
  }
  
  # --------------------------
  # PROPORCIONES (definición)
  # --------------------------
  PROPORCIONES <- list(
    list(
      nombre = "Proporción de hogares compuestos solo por personas de 60 años o más",
      num_cols = c("n_hog_60"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Proporción de hogares con jefatura femenina",
      num_cols = c("n_jefatura_mujer"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Proporción de personas mayores de 15 años que tienen una religión o credo",
      num_cols = c("n_religion"),
      den_cols = c("n_per")
    ),
    list(
      nombre = "Proporción de Personas ocupadas en actividades del Sector Primario de la Economía (RRNN)",
      num_cols = c("n_caenes_A", "n_caenes_B"),
      den_cols = c("n_ocupado", "n_desocupado")
    ),
    list(
      nombre = "Proporción de fuerza de trabajo dependiente de Recursos Naturales",
      num_cols = c("n_ciuo_6"),
      den_cols = c("n_ocupado", "n_desocupado")
    ),
    list(
      nombre = "Proporción de viviendas hacinadas",
      num_cols = c("n_viv_hacinadas"),
      den_cols = c("n_vp_ocupada")
    ),
    list(
      nombre = "Hogares con teléfono móvil / celular / smartphone",
      num_cols = c("n_serv_tel_movil"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Hogares con computador (escritorio o portátil)",
      num_cols = c("n_serv_compu"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Hogares con tablet",
      num_cols = c("n_serv_tablet"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Hogares con internet fija",
      num_cols = c("n_serv_internet_fija"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Hogares con internet móvil (celular/tablet/BAM)",
      num_cols = c("n_serv_internet_movil"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Hogares con internet satelital",
      num_cols = c("n_serv_internet_satelital"),
      den_cols = c("n_hog")
    ),
    list(
      nombre = "Hogares con acceso a internet (total)",
      num_cols = c("n_internet"),
      den_cols = c("n_hog")
    )
  )
  
  make_proporciones_por_grupo <- function(df, props, dec = 2) {
    
    safe_sum <- function(x) sum(suppressWarnings(as.numeric(x)), na.rm = TRUE)
    
    grupos <- split(df, df$GRUPO)
    res <- list()
    
    for (g in names(grupos)) {
      
      df_g <- grupos[[g]]
      
      filas <- lapply(props, function(p) {
        
        faltan_num <- setdiff(p$num_cols, names(df_g))
        faltan_den <- setdiff(p$den_cols, names(df_g))
        
        if (length(faltan_num) > 0 || length(faltan_den) > 0) {
          return(data.frame(
            Indicador = p$nombre,
            Numerador = NA_real_,
            Denominador = NA_real_,
            Porcentaje = NA_real_,
            Valor = NA_character_,
            stringsAsFactors = FALSE
          ))
        }
        
        num <- sum(vapply(p$num_cols, function(cc) safe_sum(df_g[[cc]]), numeric(1)))
        den <- sum(vapply(p$den_cols, function(cc) safe_sum(df_g[[cc]]), numeric(1)))
        
        pct <- if (!is.na(den) && den > 0) round(100 * num / den, dec) else NA_real_
        
        valor_txt <- if (is.na(pct)) {
          as.character(num)
        } else {
          paste0(
            num, " (",
            formatC(pct, format = "f", digits = dec),
            "%)"
          )
        }
        
        data.frame(
          Indicador = p$nombre,
          Numerador = num,
          Denominador = den,
          Porcentaje = pct,
          Valor = valor_txt,
          stringsAsFactors = FALSE
        )
      })
      
      res[[g]] <- dplyr::bind_rows(filas)
    }
    
    res
  }
  
  # --------------------------
  # Export Excel (manteniendo “aire” entre tablas)
  # + PROPORCIONES arriba de todo
  # --------------------------
  exportar_tablas_excel_into_wb <- function(wb, tablas_por_grupo, proporciones_por_grupo) {
    
    bold_title <- openxlsx::createStyle(textDecoration = "bold", fontSize = 13)
    bold_head  <- openxlsx::createStyle(textDecoration = "bold")
    bold_total <- openxlsx::createStyle(textDecoration = "bold", border = "top")
    
    for (g in names(tablas_por_grupo)) {
      
      openxlsx::addWorksheet(wb, g)
      row <- 1
      
      # ---- PROPORCIONES ARRIBA ----
      prop_df <- proporciones_por_grupo[[g]]
      
      openxlsx::writeData(wb, g, "Proporciones", startRow = row, startCol = 1)
      openxlsx::addStyle(wb, g, bold_title, rows = row, cols = 1, gridExpand = TRUE)
      
      row <- row + 2
      
      openxlsx::writeData(wb, g, prop_df, startRow = row, startCol = 1)
      openxlsx::addStyle(wb, g, bold_head, rows = row, cols = 1:ncol(prop_df), gridExpand = TRUE)
      
      # ✅ ESPACIO después de proporciones
      row <- row + nrow(prop_df) + 4
      
      # ---- TABLAS (con “aire” entre cada una) ----
      for (nombre_tabla in names(tablas_por_grupo[[g]])) {
        
        tab <- tablas_por_grupo[[g]][[nombre_tabla]]
        
        openxlsx::writeData(wb, g, nombre_tabla, startRow = row, startCol = 1)
        openxlsx::addStyle(wb, g, bold_title, rows = row, cols = 1, gridExpand = TRUE)
        
        row <- row + 2
        
        openxlsx::writeData(wb, g, tab, startRow = row, startCol = 1)
        openxlsx::addStyle(wb, g, bold_head, rows = row, cols = 1:3, gridExpand = TRUE)
        
        openxlsx::addStyle(
          wb, g, bold_total,
          rows = row + nrow(tab), cols = 1:3,
          gridExpand = TRUE
        )
        
        # ✅ ESPACIO entre tablas
        row <- row + nrow(tab) + 4
      }
      
      # anchos
      openxlsx::setColWidths(wb, g, 1, 70)     # Categoria/Indicador
      openxlsx::setColWidths(wb, g, 2:4, 16)   # Numerador/Denominador/Porcentaje o n/pct
      openxlsx::setColWidths(wb, g, 5, 26)     # Valor (texto)
    }
  }
  
  # --------------------------
  # Map selection state
  # --------------------------
  seleccion <- reactiveVal(character())
  reset_key <- reactiveVal(0)
  observeEvent(input$limpiar_seleccion, {
    seleccion(character())
    reset_key(reset_key() + 1)
  })
  
  comunas_filtradas <- reactive({
    cut_comunas |>
      filter(COD_REGION == input$region)
  })
  
  observeEvent(input$region, {
    lista <- comunas_filtradas() |>
      select(COMUNA, CUT) |>
      tibble::deframe()
    updateSelectInput(session, "comuna", choices = lista)
  })
  
  observeEvent(input$azar_comuna, {
    req(input$region)
    comuna_aleatoria <- comunas_filtradas() |>
      slice_sample(n = 1) |>
      pull(CUT)
    updateSelectInput(session, "comuna", selected = comuna_aleatoria)
  })
  
  # --------------------------
  # Datos mapa (parquet ENTIDADES/MANZANAS)
  # --------------------------
  datos <- reactive({
    arrow::open_dataset(
      TERRITORIO_PARQUET,
      partitioning = c("COD_REGION", "CUT")
    )
  })
  
  datos_filtrados <- reactive({
    req(input$comuna)
    
    df <- datos() |>
      filter(
        COD_REGION == as.numeric(input$region),
        CUT        == as.numeric(input$comuna)
      ) |>
      collect()
    
    # asegurar que existe columna de nombre para tooltip/tabla
    if (!("LOCALIDAD" %in% names(df))) {
      if ("NOMBRE" %in% names(df)) {
        df <- df |> mutate(LOCALIDAD = as.character(.data[["NOMBRE"]]))
      } else if ("NOM" %in% names(df)) {
        df <- df |> mutate(LOCALIDAD = as.character(.data[["NOM"]]))
      } else {
        df <- df |> mutate(LOCALIDAD = as.character(.data[[ID_COL]]))
      }
    }
    
    # asegurar columnas mínimas para el mapa
    keep_cols <- intersect(c("AREA_C", ID_COL, "OBJECTID", "SHAPE", "LOCALIDAD"), names(df))
    validate(need("SHAPE" %in% keep_cols, "El parquet no trae columna SHAPE (geometría)."))
    
    df |>
      select(any_of(keep_cols)) |>
      st_as_sf(crs = 4326)
  })
  
  territorio <- reactive({
    cut_comunas |>
      filter(COD_REGION == input$region, CUT == input$comuna)
  })
  
  output$titulo_comuna <- renderText(territorio()$COMUNA)
  output$titulo_region <- renderText(territorio()$REGION)
  
  mapa <- reactive({
    req(input$comuna)
    validate(need(nrow(datos_filtrados()) >= 1, "No hay datos para esta selección."))
    
    datos_filtrados() |>
      ggplot() +
      aes(fill = AREA_C, data_id = .data[[ID_COL]]) +
      geom_sf_interactive(
        aes(tooltip = paste0(
          "<span class='id_variable'>", ID_LABEL, ":</span> ",
          .data[[ID_COL]], "<br>",
          "<span class='id_variable'>Nombre:</span> ", LOCALIDAD
        )),
        color = "#181818",
        linewidth = 0.1
      ) +
      scale_fill_manual(
        values = c("URBANO" = "#3C533C", "RURAL" = "#A9C272"),
        na.translate = FALSE
      ) +
      theme(
        axis.text.x = element_text(angle = 90, vjust = .5),
        axis.ticks = element_blank(),
        panel.background = element_blank(),
        plot.background = element_rect(fill = "#181818", color = NA),
        legend.background = element_rect(fill = "#181818", color = NA),
        panel.grid = element_line(color = "#333333"),
        axis.text = element_text(color = "#444444"),
        legend.key.size = unit(5, "mm"),
        legend.text = element_text(color = "#666666", margin = margin(l = 4, r = 6))
      ) +
      guides(fill = guide_legend(title = NULL, position = "top"))
  })
  
  output$mapa_interactivo <- renderGirafe({
    dummy <- reset_key()
    req(mapa())
    
    girafe(
      ggobj = mapa(),
      bg = "#181818",
      width_svg = 7,
      height_svg = 7,
      options = list(
        opts_sizing(rescale = TRUE),
        opts_toolbar(
          hidden = c("selection"),
          fixed = TRUE,
          tooltips = list(
            zoom_on = "activar zoom y desplazamiento",
            zoom_off = "desactivar zoom",
            zoom_rect = "zoom desde cuadro de selección",
            zoom_reset = "resetear zoom"
          ),
          saveaspng = FALSE
        ),
        opts_sizing(width = .7),
        opts_selection(type = "multiple", only_shiny = TRUE),
        opts_zoom(duration = 400, min = 1, max = 10),
        opts_hover(css = "stroke: #AE027E; stroke-width: 1;"),
        opts_tooltip(css = "background-color: #181818; color: #FFFFFF; font-size: 9pt; padding: 3px; border-radius: 3px;")
      )
    )
  }) |>
    bindCache(input$comuna, reset_key())
  
  observeEvent(input$mapa_interactivo_selected, {
    seleccion(input$mapa_interactivo_selected)
  }, ignoreNULL = TRUE)
  
  output$click_table <- renderTable({
    ids <- seleccion()
    req(length(ids) > 0)
    
    datos_filtrados() |>
      st_drop_geometry() |>
      transmute(
        MANZENT = as.character(.data[[ID_COL]]),
        LOCALIDAD = LOCALIDAD
      ) |>
      distinct() |>
      filter(MANZENT %in% ids) |>
      mutate(.ord = match(MANZENT, ids)) |>
      arrange(.ord) |>
      select(-.ord)
  }, striped = TRUE, bordered = TRUE, spacing = "xs")
  
  # --------------------------
  # Tablas desde Excel
  # --------------------------
  TABLAS <- reactive({
    build_tablas_from_excel("config_tablas_censo2024.xlsx", sheet = "tablas")
  })
  
  codigos_usuario <- reactive({
    req(input$excel_codigos)
    
    df <- readxl::read_excel(input$excel_codigos$datapath)
    names(df) <- toupper(names(df))
    
    validate(need("MANZENT" %in% names(df), "El Excel debe incluir una columna MANZENT."))
    
    if (!("GRUPO" %in% names(df))) df$GRUPO <- df$MANZENT
    
    df |>
      transmute(
        MANZENT = as.character(MANZENT),
        GRUPO = as.character(GRUPO)
      ) |>
      filter(!is.na(MANZENT), MANZENT != "") |>
      mutate(GRUPO = ifelse(is.na(GRUPO) | GRUPO == "", MANZENT, GRUPO)) |>
      distinct()
  })
  
  datos_full <- reactive({
    arrow::open_dataset(
      TERRITORIO_PARQUET,
      partitioning = c("COD_REGION", "CUT")
    )
  })
  
  datos_para_tablas <- eventReactive(input$generar_tablas, {
    
    cod <- codigos_usuario()
    tablas_cfg <- TABLAS()
    cols_tablas <- unique(unlist(lapply(tablas_cfg, function(x) x$cols)))
    cols_props  <- unique(unlist(lapply(PROPORCIONES, function(p) c(p$num_cols, p$den_cols))))
    
    # Columnas numéricas necesarias (no metas AREA_C aquí)
    cols_needed_num <- unique(c(cols_tablas, cols_props, "n_per"))
    
    cod <- cod |>
      mutate(
        MANZENT = as.character(MANZENT),
        GRUPO   = as.character(GRUPO)
      )
    
    df <- datos_full() |>
      mutate(
        MANZENT_KEY = arrow::cast(arrow::cast(MANZENT, arrow::int64()), arrow::utf8())
      ) |>
      filter(MANZENT_KEY %in% cod$MANZENT) |>
      # ✅ TRAE AREA_C explícitamente siempre
      select(MANZENT_KEY, AREA_C, any_of(cols_needed_num)) |>
      collect()
    
    df <- df |>
      rename(MANZENT = MANZENT_KEY) |>
      mutate(
        MANZENT = as.character(MANZENT),
        AREA_C  = as.character(AREA_C),
        # ✅ convierte a numérico solo lo numérico
        across(any_of(cols_needed_num), ~ suppressWarnings(as.numeric(.)))
      )
    
    # ✅ Derivadas (urban/rural)
    df <- df |>
      mutate(
        AREA_C = toupper(trimws(AREA_C)),
        n_per  = ifelse(is.na(n_per), 0, n_per),
        n_per_urbano = ifelse(AREA_C == "URBANO", n_per, 0),
        n_per_rural  = ifelse(AREA_C == "RURAL",  n_per, 0)
      )
    
    df |>
      left_join(cod, by = "MANZENT")
  })
  
  
  

  
  ids_no_encontrados <- eventReactive(input$generar_tablas, {
    ids <- codigos_usuario()$MANZENT
    encontrados <- unique(as.character(datos_para_tablas()$MANZENT))
    setdiff(ids, encontrados)
  })
  
  tablas_generadas <- eventReactive(input$generar_tablas, {
    generar_tablas_baseline(datos_para_tablas(), TABLAS())
  })
  
  proporciones_generadas <- eventReactive(input$generar_tablas, {
    make_proporciones_por_grupo(datos_para_tablas(), PROPORCIONES, dec = 2)
  })
  
  output$estado_tablas <- renderUI({
    req(input$generar_tablas)
    
    miss <- ids_no_encontrados()
    msg <- if (length(miss) == 0) {
      "Listo: tablas generadas."
    } else {
      paste0("Tablas generadas. Ojo: ", length(miss), " MANZENT no fueron encontrados en el parquet.")
    }
    
    div(style="margin-top:10px; color:#bbb;", msg)
  })
  
  output$descargar_tablas <- downloadHandler(
    filename = function() {
      paste0("tablas_linea_base_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      
      wb <- openxlsx::createWorkbook()
      
      openxlsx::addWorksheet(wb, "Codigos_usuario")
      openxlsx::writeData(wb, "Codigos_usuario", codigos_usuario())
      
      miss <- ids_no_encontrados()
      openxlsx::addWorksheet(wb, "IDs_no_encontrados")
      openxlsx::writeData(wb, "IDs_no_encontrados", data.frame(MANZENT = miss))
      
      # ✅ escribe: PROPORCIONES arriba + TABLAS con espacio
      exportar_tablas_excel_into_wb(wb, tablas_generadas(), proporciones_generadas())
      
      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
