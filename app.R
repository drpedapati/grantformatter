library(shiny)
library(tidyverse)
library(scales)
library(flextable)


# Define UI for data upload app ----
ui <- fluidPage(
  # App title ----
  titlePanel("Biosketch Grant Formatter"),
  # Sidebar layout with input and output definitions ----
  sidebarLayout(
    
    # Sidebar panel for inputs ----
    sidebarPanel(
      
      # Button
      downloadButton("downloadData", "Download Excel Fill-in Template"),
      
      # Horizontal line ----
      tags$hr(),
      
      # Input: Select a file ----
      fileInput("file1", "Upload Completed Excel Template",
                multiple = TRUE,
                accept = c(".xlsx",
                           ".xls")),
      
      # Horizontal line ----
      tags$hr(),
      
      downloadButton("convert", "Format and Download to Word")

      
    ),
    
    # Main panel for displaying outputs ----
    mainPanel(
      
      # Output: Data file ----
      tableOutput("contents")
      
    )
    
  )
)

# Define server logic to read selected file ----
server <- function(input, output, session) {
  
  output$contents <- renderTable({
    
    # input$file1 will be NULL initially. After the user selects
    # and uploads a file, head of that data file by default,
    # or all rows if selected, will be shown.
    
    req(input$file1)
    
    df <- readxl::read_excel(path = input$file1$datapath, sheet = 1)
    
    return(df)  
    
  })
  

    output$convert <- downloadHandler(
      filename = function() {
        paste("Formatted_Grant_List.docx")
      },
      content = function(file) {
        df <- readxl::read_excel(path = input$file1$datapath, sheet = 1)
        
        df  %>% 
          mutate(
            main = sprintf("Title: %s\r\nStatus: %s %s\nAward Type: %s\tTotal Funding: %s\nAgency: %s\t\nDates: %s-%s\nPI: %s\tRole: %s", 
                           Title, Status, ifelse(!is.na(Notes), paste0('(',Notes,')'),''), Type, dollar(Amount),  Sponsor, Start, End, PI, Role)
          ) %>% select(main) %>% bind_rows() %>% flextable(cwidth = 6.5) %>% flextable::save_as_docx(path = file)   
      }
    )
  
  # Downloadable csv of selected dataset ----
  
  output$downloadData <- downloadHandler(
    filename = function() {
      paste("Grant_Blank_Template.xlsx")
    },
  content = function(file) {
    file.copy(from = "Grant_Blank_Template.xlsx", to = file)
  }
  )


}
# Run the app ----
shinyApp(ui, server)