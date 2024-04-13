# Packages
library(shiny)
library(bslib)
library(shinythemes)
library(tidyverse)
library(magrittr)
library(openxlsx)
library(shinyjs)

# User interface  
ui <- page_sidebar(
  
  theme = shinytheme("cosmo"),
  
  title = h1("CEGEP CI Calculator"),
  
  tags$head(
    tags$style(
      ".progress-bar {
          background-image: linear-gradient(to right, lightgreen , darkgreen) !important;
          background-size: auto !important;
      }
      
      .btn:hover{
          border-color: darkgrey !important;
          background-color: darkgrey;
          box-shadow: 0px;
          border:0px;
          -webkit-box-shadow: 0px;
      }
      
      .btn:focus{
          border-color: darkgrey !important;
          background-color: darkgrey !important;
          box-shadow: 0px;
          border: 0px;
          -webkit-box-shadow: 0px;
      }")
  ),
  
  tags$script(src="fileInput_text.js"),
  
  sidebar = sidebar(
    br(),
    h4(tags$b("How to use this application")),
    h5("1. Download the", tags$a(href="https://github.com/chloedebyser/CEGEP-CI-CalculatoR/raw/main/Course%20allocation%20template.xlsx", "course allocation template form")),
    h5("2. Enter your course allocations as per instructions in the README tab"),
    h5("3. Upload your form to the app"),
    h5("4. Click `Caculate CI`"),
    h5("5. Your CI results will download automatically"),
    br(),
    br(),
    br(),
    br(),
    br(),
    br(),
    br(),
    br(),
    br(),
    br(),
    h6("Source code",
       tags$a(href = "https://github.com/chloedebyser/CEGEP-CI-CalculatoR",
              icon("github", "fa-2x"))),
    tags$style(".fa-github {color:#13294B}"),
    open="always"
    ),
  
  layout_columns(
    fileInput("file",
              label = h3(tags$b("Upload your course allocation form")),
              placeholder = "or drop files here",
              multiple = F,
              accept = c(".xlsx"),
              width = 700)
  ),
  tags$div(uiOutput("runButton"), align = "center"),
  br(),
  conditionalPanel(condition = "input.runScript == 1",
                   tags$div(downloadButton("downloadData", "Download"), align = "center"))
)

# Server function
server <- function(input, output){
  
  source("R/CI-CalculatoR.R")
  
  file_df <- reactive({
    req(input$file)
    df <- input$file
  })
  
  output$runButton <- renderUI({
    if(is.null(file_df())) return()
    actionButton("runScript", "Caculate CI", style = "border-color: darkgrey !important")
  })
  
  r <- reactiveValues(convertRes = NULL)
  
  observeEvent(input$runScript, {
    
    r$convertRes <- CICalculatoR(input = file_df()$datapath)
    
    output$downloadData <- downloadHandler(
      filename = "CI Results.xlsx",
      content = function(file) {saveWorkbook(r$convertRes, file, overwrite = T)}
    )
  }, ignoreNULL = TRUE, ignoreInit = TRUE)
}

# Shiny app
shinyApp(ui, server)
