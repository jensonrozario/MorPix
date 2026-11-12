# Library -----------------------------------------------------------------
library(shiny)
library(EBImage)
library(writexl)
library(readxl)


# SETTINGS ----------------------------------------------------------------
output_xlsx <- "Specimen_Measurements_Shiny.xlsx"

traits <- c(
  "Scale_Length",
  "Total_Length",
  "Carapace_Length",
  "Antero_Lateral_Margin",
  "Total_Width"
)
trait_colors <- c("cyan", "magenta", "yellow", "green", "orange")


# USER INTERFACE (UI) -----------------------------------------------------
ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      #plot-wrapper, #plot-wrapper * {
        cursor: none !important;
      }
      #plot-wrapper.no-loupe, #plot-wrapper.no-loupe * {
        cursor: crosshair !important;
      }
      #plot-wrapper {
        position: relative;
        display: block;
        width: 100%;
        line-height: 0;
        margin-bottom: 12px;
      }
      #plot-wrapper img {
        display: block;
      }
      #circular-loupe {
        position: absolute;
        width: 200px;
        height: 200px;
        border-radius: 50%;
        border: 3px solid #fff700;
        box-shadow: 0 4px 14px rgba(0,0,0,0.45);
        pointer-events: none;
        display: none;
        background-repeat: no-repeat;
        image-rendering: -webkit-optimize-contrast;
        z-index: 999;
      }
      /* Semi-transparent fluorescent lemon yellow center dot */
      #loupe-dot {
        position: absolute;
        top: 50%;
        left: 50%;
        width: 5px;
        height: 5px;
        transform: translate(-50%, -50%);
        background: rgba(255, 247, 0, 0.45);
        border: 1px solid rgba(255, 247, 0, 0.95);
        border-radius: 50%;
        box-shadow: 0 0 2px rgba(0, 0, 0, 0.85);
      }
      /* Fluorescent lemon yellow open-center cross arms around the dot */
      .loupe-arm {
        position: absolute;
        background: rgba(255, 247, 0, 0.92);
        box-shadow: 0 0 1px rgba(0, 0, 0, 0.85);
      }
      .loupe-top    { left: 50%; top: 22%;    width: 1.5px; height: 20%; transform: translateX(-50%); }
      .loupe-bottom { left: 50%; bottom: 22%; width: 1.5px; height: 20%; transform: translateX(-50%); }
      .loupe-left   { top: 50%;  left: 22%;   height: 1.5px; width: 20%; transform: translateY(-50%); }
      .loupe-right  { top: 50%;  right: 22%;  height: 1.5px; width: 20%; transform: translateY(-50%); }

      /* Compact sidebar styling */
      .compact-box {
        padding: 7px 10px;
        border-radius: 4px;
        margin-bottom: 8px;
      }
      .step-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 4px;
        padding: 3px 6px;
        border-radius: 3px;
        background: #ffffff;
        border: 1px solid #d6e4f0;
      }
      .step-row.active-row {
        background: #dceeff;
        border-color: #007bff;
        font-weight: bold;
      }
      .step-label {
        cursor: pointer;
        margin: 0;
        flex-grow: 1;
        font-size: 13px;
      }
    ")),
    tags$script(HTML("
      var hdImageUrl = '';

      Shiny.addCustomMessageHandler('set_hd_image', function(url) {
        hdImageUrl = url;
      });

      $(document).on('shiny:connected', function() {
        var wrapper = document.getElementById('plot-wrapper');
        var loupe = document.getElementById('circular-loupe');
        if (!wrapper || !loupe) return;

        $('#enable_loupe').on('change', function() {
          if ($(this).is(':checked')) {
            wrapper.classList.remove('no-loupe');
          } else {
            wrapper.classList.add('no-loupe');
            loupe.style.display = 'none';
          }
        });

        wrapper.addEventListener('mousemove', function(e) {
          var enabled = $('#enable_loupe').is(':checked');
          var imgEl = wrapper.querySelector('img');
          if (!enabled || !imgEl) {
            loupe.style.display = 'none';
            return;
          }

          var wrapRect = wrapper.getBoundingClientRect();
          var rect = imgEl.getBoundingClientRect();
          var x = e.clientX - rect.left;
          var y = e.clientY - rect.top;

          if (x < 0 || y < 0 || x > rect.width || y > rect.height) {
            loupe.style.display = 'none';
            return;
          }

          var zoom = parseFloat($('#loupe_zoom').val()) || 3;
          var loupeSize = 200;
          var half = loupeSize / 2;

          var bgSrc = (hdImageUrl && hdImageUrl.length > 0) ? hdImageUrl : imgEl.src;

          loupe.style.display = 'block';
          loupe.style.left = (e.clientX - wrapRect.left - half) + 'px';
          loupe.style.top = (e.clientY - wrapRect.top - half) + 'px';
          loupe.style.backgroundImage = 'url(\"' + bgSrc + '\")';
          loupe.style.backgroundSize = (rect.width * zoom) + 'px ' + (rect.height * zoom) + 'px';
          loupe.style.backgroundPosition = '-' + (x * zoom - half) + 'px -' + (y * zoom - half) + 'px';
        });

        wrapper.addEventListener('mouseleave', function() {
          loupe.style.display = 'none';
        });
      });
    "))
  ),
  
  titlePanel("MorPix: Precision Specimen Measurements from Images"),
  
  sidebarLayout(
    sidebarPanel(
      width = 4,
      style = "padding: 10px;",
      
      # Two-line Note Banner
      div(
        class = "compact-box",
        style = "background-color: #fff3cd; color: #856404; border-left: 4px solid #ffc07b; font-size: 12px;",
        strong("Note:"),
        tags$p(style = "margin: 3px 0 2px 0;", "1. Copy all images in the working directory for measuring."),
        tags$p(style = "margin: 0;", "2. Image size should be reduced for reduced response time.")
      ),
      
      # Compact Active Directory + Photo Selector
      div(
        class = "compact-box",
        style = "background-color: #f8f9fa; border: 1px solid #ddd;",
        div(style = "display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;",
            strong(style = "font-size: 12px;", "Working Directory:"),
            actionButton("open_dir_modal", "Change Folder", class = "btn-primary btn-xs")
        ),
        div(style = "word-break: break-all; font-size: 11px; color: #555; margin-bottom: 6px;",
            textOutput("active_dir_text")),
        selectInput("photo_selector", label = NULL, choices = NULL, width = "100%"),
        div(style = "font-size: 12px; margin-top: -10px;", strong(textOutput("photo_status_badge")))
      ),
      
      # Compact Lens Settings
      div(
        class = "compact-box",
        style = "background-color: #f8f9fa; border: 1px solid #ddd;",
        fluidRow(
          column(6, checkboxInput("enable_loupe", "HD Magnifier Lens", value = TRUE)),
          column(6, sliderInput("loupe_zoom", "Zoom:", min = 2, max = 8, value = 3.5, step = 0.5))
        )
      ),
      
      # Measurement Steps with individual right-side Clear buttons
      div(
        class = "compact-box",
        style = "background-color: #eef6ff; border-left: 4px solid #007bff;",
        div(style = "margin-bottom: 5px; font-size: 13px;", strong("Select Measurement Step:")),
        uiOutput("step_rows_ui"),
        
        # Compact Fine-Tune Nudge Controls
        div(
          style = "margin-top: 6px; padding-top: 6px; border-top: 1px solid #cfe2f3;",
          fluidRow(
            column(5, style = "font-size: 12px; padding-top: 4px;", strong("Nudge Point (1px):")),
            column(7,
                   div(style = "display: flex; gap: 3px;",
                       actionButton("nudge_left", "◀", class = "btn-xs btn-default"),
                       actionButton("nudge_up", "▲", class = "btn-xs btn-default"),
                       actionButton("nudge_down", "▼", class = "btn-xs btn-default"),
                       actionButton("nudge_right", "▶", class = "btn-xs btn-default")
                   )
            )
          )
        ),
        div(style = "font-size: 12px; color: #004085; margin-top: 4px;", strong(textOutput("click_status")))
      ),
      
      # Side-by-Side Scale Length (mm) and Sex Dropdown (with Berried)
      div(
        class = "compact-box",
        style = "background-color: #f8f9fa; border: 1px solid #ddd;",
        fluidRow(
          column(6, numericInput("scale_mm", "Scale Length (mm):", value = 10, min = 0.001, step = 1)),
          column(6, selectInput("sex", "Sex:",
                                choices = c("Not Determined", "Male", "Female", "Berried"),
                                selected = "Not Determined"))
        )
      ),
      
      # Save & Quit Buttons
      fluidRow(
        column(8, actionButton("save_next", "Save & Next Photo", class = "btn-success btn-block", style = "font-weight: bold;")),
        column(4, actionButton("quit_app", "Quit", class = "btn-default btn-block"))
      )
    ),
    
    mainPanel(
      width = 8,
      div(
        id = "plot-wrapper",
        div(
          id = "circular-loupe",
          div(id = "loupe-dot"),
          div(class = "loupe-arm loupe-top"),
          div(class = "loupe-arm loupe-bottom"),
          div(class = "loupe-arm loupe-left"),
          div(class = "loupe-arm loupe-right")
        ),
        plotOutput(
          "specimen_plot",
          width = "100%",
          height = "auto",
          click = "plot_click"
        )
      ),
      h5(strong("Current Photo Preview (mm)")),
      div(style = "overflow-x: auto;", tableOutput("preview_table")),
      h5(strong("Saved Excel Records")),
      div(style = "overflow-x: auto;", tableOutput("saved_table"))
    )
  )
)

# SERVER LOGIC ------------------------------------------------------------
server <- function(input, output, session) {
  
  hd_temp_dir <- file.path(tempdir(), "shiny_hd_loupe")
  if (!dir.exists(hd_temp_dir)) dir.create(hd_temp_dir, recursive = TRUE)
  shiny::addResourcePath("hd_loupe_dir", hd_temp_dir)
  empty_steps <- function() {
    lapply(1:5, function(i) data.frame(x = c(NA_real_, NA_real_), y = c(NA_real_, NA_real_)))
  }
  get_img_wh <- function(img) {
    dims <- dim(img)
    w <- head(dims, 1)
    h <- head(tail(dims, -1), 1)
    list(w = w, h = h)
  }
  draw_annotated_canvas <- function(img, step_pts, labels_vec, active_s = NULL, active_p = NULL, scale_mult = 1) {
    wh <- get_img_wh(img)
    w <- wh$w
    h <- wh$h
    par(mar = c(0, 0, 0, 0), oma = c(0, 0, 0, 0), plt = c(0, 1, 0, 1), xaxs = "i", yaxs = "i")
    plot.new()
    plot.window(xlim = c(0, w), ylim = c(0, h))
    rasterImage(as.raster(img), 0, 0, w, h, interpolate = FALSE)
    for (t_idx in 1:5) {
      pts <- step_pts[[t_idx]]
      col <- trait_colors[[t_idx]]
      for (pt_i in 1:2) {
        if (!is.na(pts$x[[pt_i]])) {
          points(pts$x[[pt_i]], pts$y[[pt_i]], pch = 3, cex = 1.6 * scale_mult, lwd = 2.2 * scale_mult, col = col)
          if (!is.null(active_s) && !is.null(active_p) && t_idx == active_s && pt_i == active_p) {
            points(pts$x[[pt_i]], pts$y[[pt_i]], pch = 1, cex = 2.6 * scale_mult, lwd = 2 * scale_mult, col = "white")
          }
        }
      }
      if (!any(is.na(pts$x))) {
        lines(pts$x, pts$y, lwd = 2.5 * scale_mult, col = col)
        text(mean(pts$x), mean(pts$y), labels = labels_vec[[t_idx]], col = col, pos = 3, font = 2, cex = 1.0 * scale_mult)
      }
    }
  }
  rv <- reactiveValues(
    folder_path = NULL,
    all_files = character(0),
    all_fnames = character(0),
    saved_df = data.frame(),
    saved_coords = data.frame(),
    step_pts = empty_steps(),
    active_step = 1,
    active_point = 1
  )
  output$step_rows_ui <- renderUI({
    step_names <- c(
      "1. Scale_Length (Cyan)",
      "2. Total_Length (Magenta)",
      "3. Carapace_Length (Yellow)",
      "4. Antero_Lateral_Margin (Green)",
      "5. Total_Width (Orange)"
    )
    lapply(1:5, function(s) {
      pts <- rv$step_pts[[s]]
      is_done <- !any(is.na(pts$x))
      row_class <- if (rv$active_step == s) "step-row active-row" else "step-row"
      status_icon <- if (is_done) "✔ " else ""
      div(
        class = row_class,
        actionLink(
          inputId = paste0("select_step_", s),
          label = paste0(status_icon, step_names[[s]]),
          class = "step-label"
        ),
        actionButton(
          inputId = paste0("clear_step_", s),
          label = "Clear",
          class = "btn-warning btn-xs"
        )
      )
    })
  })
  lapply(1:5, function(s) {
    observeEvent(input[[paste0("select_step_", s)]], {
      rv$active_step <- s
      pts <- rv$step_pts[[s]]
      rv$active_point <- if (is.na(head(pts$x, 1))) 1 else 2
    })
    observeEvent(input[[paste0("clear_step_", s)]], {
      rv$step_pts[[s]] <- data.frame(x = c(NA_real_, NA_real_), y = c(NA_real_, NA_real_))
      rv$active_step <- s
      rv$active_point <- 1
    })
  })
  show_dir_dialog <- function(current_path = getwd()) {
    showModal(modalDialog(
      title = "Set Working Directory",
      div(
        style = "background-color: #fff3cd; color: #856404; padding: 12px; border-left: 4px solid #ffc07b; border-radius: 4px; margin-bottom: 15px;",
        tags$p(style = "margin: 0 0 6px 0; font-weight: bold;",
               "1. Copy all images in the working directory for measuring."),
        tags$p(style = "margin: 0; font-weight: bold;",
               "2. Image size should be reduced for reduced response time.")
      ),
      textInput("modal_work_dir", "Working Directory Path:", value = current_path, width = "100%"),
      actionButton("modal_browse_btn", "Browse Folder...", icon = icon("folder-open")),
      footer = tagList(
        actionButton("modal_confirm_btn", "Set Working Directory & Load Photos", class = "btn-success")
      ),
      easyClose = FALSE
    ))
  }
  observeEvent(TRUE, {
    show_dir_dialog(getwd())
  }, once = TRUE)
  observeEvent(input$open_dir_modal, {
    cur <- if (!is.null(rv$folder_path)) rv$folder_path else getwd()
    show_dir_dialog(cur)
  })
  observeEvent(input$modal_browse_btn, {
    chosen_dir <- NULL
    if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
      chosen_dir <- rstudioapi::selectDirectory(caption = "Select Working Directory")
    } else if (.Platform$OS.type == "windows") {
      chosen_dir <- choose.dir(caption = "Select Working Directory")
    } else if (requireNamespace("tcltk", quietly = TRUE)) {
      chosen_dir <- tcltk::tk_choose.dir(caption = "Select Working Directory")
    }
    if (!is.null(chosen_dir) && !is.na(chosen_dir) && nzchar(chosen_dir)) {
      updateTextInput(session, "modal_work_dir", value = chosen_dir)
    }
  })
  observeEvent(input$modal_confirm_btn, {
    raw_dir <- trimws(input$modal_work_dir)
    if (!nzchar(raw_dir) || !dir.exists(raw_dir)) {
      showNotification("Directory does not exist. Please enter or browse a valid folder.", type = "error")
      return()
    }
    clean_dir <- normalizePath(raw_dir, winslash = "/", mustWork = TRUE)
    all_entries <- list.files(clean_dir, full.names = TRUE)
    all_entries <- all_entries[!dir.exists(all_entries)]
    valid_exts <- c("jpg", "jpeg", "png", "tif", "tiff")
    files <- all_entries[tolower(tools::file_ext(all_entries)) %in% valid_exts]
    if (length(files) == 0) {
      showNotification("No image files (jpg, png, tif) found in this folder!", type = "error")
      return()
    }
    setwd(clean_dir)
    rv$folder_path <- clean_dir
    rv$all_files <- files
    rv$all_fnames <- basename(files)
    out_path <- file.path(clean_dir, output_xlsx)
    rv$saved_df <- if (file.exists(out_path)) {
      as.data.frame(read_excel(out_path))
    } else {
      data.frame()
    }
    rv$saved_coords <- if (file.exists(out_path) && "Coordinates" %in% excel_sheets(out_path)) {
      as.data.frame(read_excel(out_path, sheet = "Coordinates"))
    } else {
      data.frame()
    }
    rv$step_pts <- empty_steps()
    refresh_dropdown()
    removeModal()
    showNotification(sprintf("Working directory set! Loaded %d images.", length(files)), type = "message")
  })
  output$active_dir_text <- renderText({
    if (is.null(rv$folder_path)) "Not set yet" else rv$folder_path
  })
  refresh_dropdown <- function(select_target = NULL) {
    req(length(rv$all_fnames) > 0)
    done <- rv$saved_df$File_Name
    labels <- ifelse(rv$all_fnames %in% done,
                     paste0("[SAVED - Edit] ", rv$all_fnames),
                     paste0("[NEW] ", rv$all_fnames))
    choices <- setNames(rv$all_fnames, labels)
    if (is.null(select_target)) {
      unmeasured <- rv$all_fnames[!rv$all_fnames %in% done]
      select_target <- if (length(unmeasured) > 0) head(unmeasured, 1) else head(rv$all_fnames, 1)
    }
    updateSelectInput(session, "photo_selector", choices = choices, selected = select_target)
  }
  current_img <- reactive({
    req(input$photo_selector, length(rv$all_files) > 0)
    idx <- match(input$photo_selector, rv$all_fnames)
    req(!is.na(idx))
    readImage(rv$all_files[[idx]])
  })
  observeEvent(input$photo_selector, {
    req(input$photo_selector, length(rv$all_fnames) > 0)
    fname <- input$photo_selector
    img <- current_img()
    temp_png_name <- paste0("hd_preview_", as.integer(Sys.time()), ".png")
    temp_png_path <- file.path(hd_temp_dir, temp_png_name)
    unlink(list.files(hd_temp_dir, full.names = TRUE))
    writeImage(img, temp_png_path, type = "png")
    session$sendCustomMessage("set_hd_image", paste0("hd_loupe_dir/", temp_png_name))
    if (nrow(rv$saved_coords) > 0 && fname %in% rv$saved_coords$File_Name) {
      row_c <- rv$saved_coords[rv$saved_coords$File_Name == fname, ][1, ]
      restored <- empty_steps()
      for (s in 1:5) {
        prefix <- traits[[s]]
        restored[[s]] <- data.frame(
          x = c(row_c[[paste0(prefix, "_x1")]], row_c[[paste0(prefix, "_x2")]]),
          y = c(row_c[[paste0(prefix, "_y1")]], row_c[[paste0(prefix, "_y2")]])
        )
      }
      rv$step_pts <- restored
      row_m <- rv$saved_df[rv$saved_df$File_Name == fname, ][1, ]
      updateNumericInput(session, "scale_mm", value = row_m$Scale_mm)
      updateSelectInput(session, "sex", selected = row_m$Sex)
    } else {
      rv$step_pts <- empty_steps()
    }
    rv$active_step <- 1
    rv$active_point <- 1
  })
  output$photo_status_badge <- renderText({
    if (length(rv$all_fnames) == 0) {
      return("Waiting for working directory...")
    }
    n_done <- sum(rv$all_fnames %in% rv$saved_df$File_Name)
    sprintf("Progress: %d of %d photos saved", n_done, length(rv$all_fnames))
  })
  output$click_status <- renderText({
    s <- rv$active_step
    p <- rv$active_point
    pts <- rv$step_pts[[s]]
    if (!any(is.na(pts$x))) {
      sprintf("Done: %s recorded (click Clear to redo).", traits[[s]])
    } else {
      sprintf("Click %s for %s.", ifelse(p == 1, "Start Point (1)", "End Point (2)"), traits[[s]])
    }
  })
  observeEvent(input$plot_click, {
    req(length(rv$all_files) > 0)
    s <- rv$active_step
    pts <- rv$step_pts[[s]]
    p <- if (is.na(head(pts$x, 1))) 1 else 2
    pts$x[[p]] <- input$plot_click$x
    pts$y[[p]] <- input$plot_click$y
    rv$step_pts[[s]] <- pts
    rv$active_point <- p
    if (p == 1) {
      rv$active_point <- 2
    } else if (p == 2 && s < 5) {
      next_pts <- rv$step_pts[[s + 1]]
      if (is.na(head(next_pts$x, 1))) {
        rv$active_step <- s + 1
        rv$active_point <- 1
      }
    }
  })
  nudge_point <- function(dx, dy) {
    s <- rv$active_step
    p <- rv$active_point
    pts <- rv$step_pts[[s]]
    if (!is.na(pts$x[[p]])) {
      pts$x[[p]] <- pts$x[[p]] + dx
      pts$y[[p]] <- pts$y[[p]] + dy
      rv$step_pts[[s]] <- pts
    }
  }
  observeEvent(input$nudge_up,    { nudge_point(0,  1) })
  observeEvent(input$nudge_down,  { nudge_point(0, -1) })
  observeEvent(input$nudge_left,  { nudge_point(-1, 0) })
  observeEvent(input$nudge_right, { nudge_point( 1, 0) })
  output$specimen_plot <- renderPlot({
    draw_annotated_canvas(
      img = current_img(),
      step_pts = rv$step_pts,
      labels_vec = traits,
      active_s = rv$active_step,
      active_p = rv$active_point,
      scale_mult = 1
    )
  },
  width = function() {
    wh <- get_img_wh(current_img())
    avail_w <- session$clientData$output_specimen_plot_width
    if (is.null(avail_w) || avail_w <= 0) avail_w <- 800
    scale_ratio <- min(avail_w / wh$w, 650 / wh$h)
    round(wh$w * scale_ratio)
  },
  height = function() {
    wh <- get_img_wh(current_img())
    avail_w <- session$clientData$output_specimen_plot_width
    if (is.null(avail_w) || avail_w <= 0) avail_w <- 800
    scale_ratio <- min(avail_w / wh$w, 650 / wh$h)
    round(wh$h * scale_ratio)
  })
  calc_current_row <- reactive({
    req(input$photo_selector, length(rv$all_files) > 0)
    get_dist <- function(s_idx) {
      pts <- rv$step_pts[[s_idx]]
      if (any(is.na(pts$x))) return(NA_real_)
      sqrt(diff(pts$x)^2 + diff(pts$y)^2)
    }
    scale_px <- get_dist(1)
    mm_per_px <- ifelse(!is.na(scale_px) && scale_px > 0, input$scale_mm / scale_px, NA_real_)
    data.frame(
      File_Name                = input$photo_selector,
      Sex                      = input$sex,
      Scale_mm                 = input$scale_mm,
      Scale_Length_px          = round(scale_px, 2),
      Total_Length_mm          = round(get_dist(2) * mm_per_px, 3),
      Carapace_Length_mm       = round(get_dist(3) * mm_per_px, 3),
      Antero_Lateral_Margin_mm = round(get_dist(4) * mm_per_px, 3),
      Total_Width_mm           = round(get_dist(5) * mm_per_px, 3),
      stringsAsFactors         = FALSE
    )
  })
  output$preview_table <- renderTable({
    calc_current_row()
  })
  save_measured_image_copy <- function(img, fname, row_data) {
    out_dir <- file.path(rv$folder_path, "Measured_Images")
    if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
    base_name <- tools::file_path_sans_ext(fname)
    out_img_path <- file.path(out_dir, paste0(base_name, "_R_Measured.png"))
    wh <- get_img_wh(img)
    avail_w <- session$clientData$output_specimen_plot_width
    if (is.null(avail_w) || avail_w <= 0) avail_w <- 800
    scale_ratio <- min(avail_w / wh$w, 650 / wh$h)
    screen_w <- round(wh$w * scale_ratio)
    screen_h <- round(wh$h * scale_ratio)
    out_w <- screen_w * 2
    out_h <- screen_h * 2
    mm_labels <- c(
      sprintf("Scale: %g mm", row_data$Scale_mm),
      sprintf("Total_Length: %.2f mm", row_data$Total_Length_mm),
      sprintf("Carapace_Length: %.2f mm", row_data$Carapace_Length_mm),
      sprintf("Antero_Lateral_Margin: %.2f mm", row_data$Antero_Lateral_Margin_mm),
      sprintf("Total_Width: %.2f mm", row_data$Total_Width_mm)
    )
    png(filename = out_img_path, width = out_w, height = out_h, units = "px", res = 192)
    draw_annotated_canvas(
      img = img,
      step_pts = rv$step_pts,
      labels_vec = mm_labels,
      active_s = NULL,
      active_p = NULL,
      scale_mult = 1
    )
    dev.off()
    return(basename(out_img_path))
  }
  observeEvent(input$save_next, {
    req(rv$folder_path, length(rv$all_files) > 0)
    all_complete <- all(sapply(rv$step_pts, function(p) !any(is.na(p$x))))
    if (!all_complete) {
      showNotification("Please complete all 5 steps (Start & End points) before saving.", type = "error")
      return()
    }
    if (is.na(input$scale_mm) || input$scale_mm <= 0) {
      showNotification("Please enter a valid positive Scale Length in mm.", type = "error")
      return()
    }
    fname <- input$photo_selector
    new_row <- calc_current_row()
    coord_list <- list(File_Name = fname)
    for (s in 1:5) {
      prefix <- traits[[s]]
      coord_list[[paste0(prefix, "_x1")]] <- head(rv$step_pts[[s]]$x, 1)
      coord_list[[paste0(prefix, "_y1")]] <- head(rv$step_pts[[s]]$y, 1)
      coord_list[[paste0(prefix, "_x2")]] <- tail(rv$step_pts[[s]]$x, 1)
      coord_list[[paste0(prefix, "_y2")]] <- tail(rv$step_pts[[s]]$y, 1)
    }
    new_coord_row <- as.data.frame(coord_list, stringsAsFactors = FALSE)
    if (nrow(rv$saved_df) > 0 && fname %in% rv$saved_df$File_Name) {
      rv$saved_df[rv$saved_df$File_Name == fname, ] <- new_row
    } else {
      rv$saved_df <- rbind(rv$saved_df, new_row)
    }
    if (nrow(rv$saved_coords) > 0 && fname %in% rv$saved_coords$File_Name) {
      rv$saved_coords[rv$saved_coords$File_Name == fname, ] <- new_coord_row
    } else {
      rv$saved_coords <- rbind(rv$saved_coords, new_coord_row)
    }
    out_path <- file.path(rv$folder_path, output_xlsx)
    write_xlsx(
      list(Measurements = rv$saved_df, Coordinates = rv$saved_coords),
      path = out_path
    )
    saved_img_name <- save_measured_image_copy(current_img(), fname, new_row)
    showNotification(paste("Saved Excel & image:", saved_img_name), type = "message")
    unmeasured <- rv$all_fnames[!rv$all_fnames %in% rv$saved_df$File_Name]
    next_target <- if (length(unmeasured) > 0) head(unmeasured, 1) else fname
    refresh_dropdown(select_target = next_target)
  })
  output$saved_table <- renderTable({
    req(nrow(rv$saved_df) > 0)
    rv$saved_df[nrow(rv$saved_df):1, , drop = FALSE]
  })
  observeEvent(input$quit_app, {
    stopApp()
  })
}
shinyApp(ui = ui, server = server)