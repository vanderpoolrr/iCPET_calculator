### iCPET Calculator Shiny app for calculating alpha distensibility
##  Author: Rebecca R Vanderpool
##  Last revised: 2023-03-10
## Parts of code adapted from: https://github.com/royfrancis/shinyapp_calendar_plot/blob/master/app.R
## load libraries
library(shiny)
library(tidyverse)
library(ggpubr)
library(cowplot)
library(viridis)
library(openxlsx)

library(colourpicker)
library(optimx)
#library(extrafont)
#loadfonts()

## load colours
# cols <- toupper(c(
#   "#bebada","#fb8072","#80b1d3","#fdb462","#b3de69","#fccde5","#FDBF6F","#A6CEE3",
#   "#56B4E9","#B2DF8A","#FB9A99","#CAB2D6","#A9C4E2","#79C360","#FDB762","#9471B4",
#   "#A4A4A4","#fbb4ae","#b3cde3","#ccebc5","#decbe4","#fed9a6","#ffffcc","#e5d8bd",
#   "#fddaec","#f2f2f2","#8dd3c7","#d9d9d9"))

cols <- toupper(magma(22))

alpha_fit <- function(p, Pw, Q, Ro, Ppa) {
  
  Ppa_fit = 1/p*(((1+p*Pw)^5 + 5*p*Ro*Q)^0.2-1);
  
  ## correct if alpha goes negative
  if (p <= 0) {
    Ppa_fit = 1/p*(((1+p*Pw)^5 + 5*p*Ro*Q)^0.2-1)+200;
  }
  
  Error_Vector=(Ppa-Ppa_fit) ;
  sse=sum(Error_Vector^2, na.rm = TRUE);
  
  
  return(sse)
}

alpha_fit_ro <- function(p, Pw, Q, a_norm, Ppa) {
  
  Ppa_fit = 1/a_norm*(((1+a_norm*Pw)^5 + 5*a_norm*p*Q)^0.2-1);
  
  ## correct if alpha goes negative
  if (p <= 0) {
    Ppa_fit = 1/a_norm*(((1+a_norm*Pw)^5 + 5*a_norm*p*Q)^0.2-1)+200;
  }
  
  Error_Vector=(Ppa-Ppa_fit) ;
  sse=sum(Error_Vector^2, na.rm = TRUE);
  
  
  return(sse)
}

# UI ---------------------------------------------------------------------------

ui <- fluidPage(
  
  pageWithSidebar(
    
    headerPanel(title="iCPET Calculator",windowTitle="iCPET Calculator"),
    
    sidebarPanel(
      
      helpText("Change settings in a top-down manner."),
      
      tags$script('
          $(document).ready(function(){
          var d = new Date();
          var target = $("#clientTime");
          target.val(d.toLocaleString());
          target.trigger("change");
          });
          '),
      div(class="row",
          #div(class="col-xs-3",style=list("padding-top: 10px;"),
          #    actionButton("date_update", label = "Update Time")
          #),
          
          
      ),
      
      
      h3("Test Dates"),
      div(class="row",
          div(class="col-md-6",
              textInput("clientTime", "Current Time", value = "")
          ),
          div(class="col-md-6",
              textInput("sub_id","Participant ID",value=paste("Sub", format(as.Date(Sys.time(),"%Y-%m-%d",tz="Europe/Stockholm"),"%Y-%m-%d"), sep="_"))
          ), 
          div(class="col-md-6",
              dateInput("in_duration_date_start","Date of the RHC (format: year-month-day)",value=format(as.Date(Sys.time(),"%Y-%m-%d",tz="Europe/Stockholm"),"%Y-%m-%d"))
          ),
          div(class="col-md-6",
              dateInput("in_duration_date_end","Date analyzed",value=format(as.Date(Sys.time(),"%Y-%m-%d",tz="Europe/Stockholm"),"%Y-%m-%d"))
          ), 
          div(class="col-md-6",
              textInput("analyst","Analyst",value="")
          )

      ),
      
      
      h3("RHC/iCPET Stages"),
      helpText("If number of exercise stages are changed, all exercise variables are reset."),
      sliderInput("in_num_tracks","Number of exercise stages including rest",value=3,min=3,max=20,step=1),
      uiOutput("tracks"),
      #### Tests for outputs
      #uiOutput("out_display"),
      #uiOutput("out_display2"),
      

      
      
      
      tags$hr(),
      helpText("2021 | Vanderpool Lab, ** Program developed for research and modeling purposes not meant for clincial diagnostics** "), 
      width = 4.5
    ),
    
    
    mainPanel(
      h3("Exercise Data used in the Analysis"),
      tableOutput("out_dataTable"),
      tags$br(),
      sliderInput("in_scale","Image preview scale",min=0,max=3,step=0.10,value=0.7),
      helpText("Scale only controls preview here and does not affect download."),
      
      tags$br(),

      imageOutput("out_plot", inline = TRUE), 

      tags$br(),
      
      
      h3("Settings"),
      
      helpText("Select Colors for the Pressure"), 
      
      ## Ability to pick colors for conditions
      div(class="row",
              div(class="col-xs-2",style=list("padding-right: 5px;"),
                  colourpicker::colourInput("in_track_colour_mPAP",label="mPAP",
                                            palette="limited",allowedCols=cols,value=cols[1])),
              div(class="col-xs-2",style=list("padding-right: 5px;"),
                  colourpicker::colourInput("in_track_colour_linehan",label="Linehan",
                                            palette="limited",allowedCols=cols,value=cols[18])),
              div(class="col-xs-2",style=list("padding-right: 5px;"),
                  colourpicker::colourInput("in_track_colour_PAWP",label="PAWP",
                                            palette="limited",allowedCols=cols,value=cols[7])),
              div(class="col-xs-2",style=list("padding-right: 5px;"),
                  colourpicker::colourInput("in_track_colour_RAP",label="RAP",
                                            palette="limited",allowedCols=cols,value=cols[12]))
      ),
      
      
      selectInput("in_legend_position",label="Legend position",
                  choices=c("top","right","left","bottom"),selected="right",multiple=F),
      div(class="row",
          div(class="col-xs-6",style=list("padding-right: 5px;"),
              selectInput("in_legend_justification",label="Legend justification",
                          choices=c("left","right"),selected="right",multiple=F)
          ),
          div(class="col-xs-6",style=list("padding-left: 5px;"),
              selectInput("in_legend_direction",label="Legend direction",
                          choices=c("vertical","horizontal"),selected="vertical",multiple=F)
          )
      ),
      div(class="row",
          # div(class="col-xs-6",style=list("padding-right: 5px;"),
          #     numericInput("in_themefontsize",label="Theme font size",value=8,step=0.5)
          # ),
          # div(class="col-xs-6",style=list("padding-left: 5px;"),
          #     numericInput("in_datefontsize",label="Date font size",value=2.5,step=0.1)
          # )
      ),
      div(class="row",
          div(class="col-xs-6",style=list("padding-right: 5px;"),
              numericInput("in_equationfontsize",label="Equation font size",value=3,step=0.5)
          ),
          div(class="col-xs-6",style=list("padding-left: 5px;"),
              numericInput("in_legendfontsize",label="Legend font size",value=8,step=0.5)
          )
      ),
      tags$br(),
      h3("Figure Download Options"),
      helpText("File type is only applicable to download and does not change preview."),
      div(class="row",
          div(class="col-xs-6",style=list("padding-right: 5px;"),
              numericInput("in_height","Height (cm)",step=0.5,value=8)
          ),
          div(class="col-xs-6",style=list("padding-left: 5px;"),
              numericInput("in_width","Width (cm)",step=0.5,value=10)
          )
      ),
      div(class="row",
          div(class="col-xs-6",style=list("padding-right: 5px;"),
              selectInput("in_res","Res/DPI",choices=c("200","300","400","500"),selected="200")
          ),
          div(class="col-xs-6",style=list("padding-left: 5px;"),
              selectInput("in_format","File type",choices=c("png","tiff","jpeg"),selected="png",multiple=FALSE,selectize=TRUE)
          )
      ),
      
      tags$hr(), 
      
      
      h3("Download/Save Options"),
      helpText("NOTE: Download Excel Option includes data table and figure in one file. "),
      
      
      downloadButton("btn_downloadplot","Download Plot"),

      downloadButton("btn_downloadData","Download Table"),
      
      downloadButton("btn_downloadData_xls","Download Excel"),
      
      
      
      
      tags$br()
      
      #verbatimTextOutput("out_display"),
      #verbatimTextOutput("out_display2")
    )
  )
)

# SERVER -----------------------------------------------------------------------

server <- function(input, output, session) {
  
  store <- reactiveValues(week=NULL)
  
  ## UI: Exercise Stages ----------------------------------------------------------------
  ## conditional ui for exercise stages (Tracks)
  
  output$tracks <- renderUI({
    num_tracks <- as.integer(input$in_num_tracks)
    
    # create date intervals
    #dseq <- seq(as.Date(input$in_duration_date_start),as.Date(input$in_duration_date_end),by=1)
    #r1 <- unique(as.character(cut(dseq,breaks=num_tracks+1)))
    
    ex_string = c("Rest", rep("Ex", num_tracks-1))
    
    
    lapply(1:num_tracks, function(i) {
      

     div(class="row",
          #div(class="col-xs-1",style=list("padding-right: 2px; padding-left: 2px;"),
          #    colourpicker::colourInput(paste0("in_track_colour_",i),label="Colour",
          #                              palette="limited",allowedCols=cols,value=cols[i])
          #),

          div(class="col-xs-2",style=list("padding-right: 2px; padding-left: 2px;"),
              numericInput(paste0("in_track_W_",i),label="W",value="", min = 0, step = 1)
          ),
          div(class="col-xs-2",style=list("padding-right: 2px; padding-left: 2px;"),
              numericInput(paste0("in_track_RAP_",i),label="RAP",value="")
          ),
          div(class="col-xs-2",style=list("padding-right: 2px; padding-left: 2px;"),
              numericInput(paste0("in_track_mPAP_",i),label="mPAP",value="")
          ),
          div(class="col-xs-2",style=list("padding-right: 2px; padding-left: 2px;"),
              numericInput(paste0("in_track_PAWP_",i),label="PAWP",value="")
          ),
          div(class="col-xs-2",style=list("padding-right: 2px; padding-left: 2px;"),
              numericInput(paste0("in_track_CO_",i),label="CO",value="", min = 0)
          ),
          div(class="col-xs-1",style=list("padding-right: 2px; padding-left: 2px; padding-top: 20px;"),
               checkboxInput(paste0("in_track_Ex_",i),label=paste0(ex_string[i],i-1),value=FALSE)
          )

      )
    })
    
  })
  
  
  
  ## OUT: out_display -------------------------------------------------------
  ## prints general variables for debugging
  
  output$out_display <- renderPrint({
    
    cat(paste0(
      "Duration Start: ",input$in_duration_date_start,"\n",
      "Duration End: ",input$in_duration_date_end,"\n",
      "Num Tracks: ",input$in_num_tracks,"\n",
      "Available colour: ",input$in_track_colour_available,"\n",
      "Weekend colour: ",input$in_track_colour_weekend,"\n",
      "Legend position: ",input$in_legend_position,"\n",
      "Legend justification: ",input$in_legend_justification,"\n",
      "Legend direction: ",input$in_legend_direction,"\n",
      "Theme font size: ",input$in_themefontsize,"\n",
      "Date font size: ",input$in_datefontsize,"\n",
      "Month font size: ",input$in_monthfontsize,"\n",
      "Legend font size: ",input$in_legendfontsize))
    
  })
  
  ## OUT: out_display2 --------------------------------------------------------
  ## prints track variables
  
  output$out_display2 <- renderPrint({
    
    num_tracks <- input$in_num_tracks
    data_table = data.frame(W = double(),
                            RAP = double(),
                            mPAP = double(), 
                            PAWP = double(), 
                            CO = double()
                              )
    for(i in 1:num_tracks)
    {
      cat(paste0("Workload: ",eval(parse(text=paste0("input$in_track_W_",i))),"\n",
                 "RAP: ",eval(parse(text=paste0("input$in_track_RAP_",i))),"\n",
                 "mPAP: ",eval(parse(text=paste0("input$in_track_mPAP_",i))),"\n",
                 "PAWP: ",eval(parse(text=paste0("input$in_track_PAWP_",i))),"\n",
                 "CO: ",eval(parse(text=paste0("input$in_track_CO_",i))),"\n",
                 "colour: ",eval(parse(text=paste0("input$in_track_colour_",i))),"\n\n"
      ))
      data_table[i,1] = eval(parse(text=paste0("input$in_track_W_",i)))
      data_table[i,2] = eval(parse(text=paste0("input$in_track_RAP_",i)))
      data_table[i,3] = eval(parse(text=paste0("input$in_track_mPAP_",i)))
      data_table[i,4] = eval(parse(text=paste0("input$in_track_PAWP_",i)))
      data_table[i,5] = eval(parse(text=paste0("input$in_track_CO_",i)))
      
    }
    #cat(paste0(data_table))
    
  })
  
  
  ## RFN: fn_plot -----------------------------------------------------------
  ## core plotting function, returns a ggplot object -- 2021 -- need to update for the exercise data... 
  ####################### Distensibility Calculations
  
  fn_plot <- reactive({
    
    shiny::req(input$in_num_tracks)
    shiny::req(input$in_duration_date_start)
    shiny::req(input$in_duration_date_end)
    shiny::req(input$in_legend_position)
    shiny::req(input$in_legend_justification)
    shiny::req(input$in_legend_direction)
    #shiny::req(input$in_themefontsize)
    #shiny::req(input$in_datefontsize)
    #shiny::req(input$in_monthfontsize)
    #shiny::req(input$in_legendfontsize)
    
    #if(is.na(input$in_themefontsize)) {themefontsize <- 10}else{themefontsize <- input$in_themefontsize}
    
    
    ### Create data_table based on the entered data
    num_tracks <- input$in_num_tracks
    data_table = data.frame(ex_cond = logical(),
                            W = double(),
                            RAP = double(),
                            mPAP = double(), 
                            PAWP = double(), 
                            CO = double()
    )
    
    
    for(i in 1:num_tracks){
      if (!is.null(eval(parse(text=paste0("input$in_track_W_",i))))) {
        data_table[i,1] = eval(parse(text=paste0("input$in_track_Ex_",i)))
        data_table[i,2] = eval(parse(text=paste0("input$in_track_W_",i)))
        data_table[i,3] = eval(parse(text=paste0("input$in_track_RAP_",i)))
        data_table[i,4] = eval(parse(text=paste0("input$in_track_mPAP_",i)))
        data_table[i,5] = eval(parse(text=paste0("input$in_track_PAWP_",i)))
        data_table[i,6] = eval(parse(text=paste0("input$in_track_CO_",i)))
      }
    }
    
    
    data_table
    
    validate(
      need(sum(data_table$ex_cond) >= 3,"Need at least 3 conditions (including rest) to perform analysis." )
    )
    
    validate(
      need(sum(data_table$ex_cond) >= 3,"Need at least 3 conditions (including rest) to perform analysis." )
    )
    
    #if(sum(data_table$ex_cond) < 3) stop("Need at least 3 conditions (including rest) to perform analysis.")
     
    # data_table = data.frame(ex_cond = c(TRUE, TRUE, TRUE, TRUE, TRUE),
    #                         W = c(0, 40, 60, 80, 100),
    #                         RAP = c(2,3,2, 3, 5),
    #                         mPAP = c(61.5, 79.4, 91.9, 96.22, 102.5),
    #                         PAWP = c(7, 13.6, 8.4, 14.4, 10.2),
    #                         CO = c(5.7, 8.4, 9.2, 10.3, 10.85))

    data_table
    
    
    ### Optimize Linehan Equation
    ex_analyze = which(data_table$ex_cond)
    data_table = data_table %>% 
      mutate(Ro = mPAP[ex_analyze[1]]/CO[ex_analyze[1]])
    
    pw_linear = lm(PAWP ~ CO, data = data_table)
    
    ### estimating pw when not avaiable based on linear fit of available pw
    pw_index_fit = which(is.na(data_table$PAWP))
    
    data_table_fit = data_table
    data_table_fit$PAWP[pw_index_fit] = pw_linear$coefficients[2]*data_table_fit$CO[pw_index_fit] + pw_linear$coefficients[1]
    
    
    ### Seems the Nelder-Mead method is similar to the Matlab Results
    result = optimx(p = c(0.05), fn=alpha_fit, Pw = data_table_fit$PAWP[ex_analyze] , 
                    Q = data_table_fit$CO[ex_analyze], 
                    Ro = data_table_fit$Ro[ex_analyze], 
                    Ppa = data_table_fit$mPAP[ex_analyze])
    a_fit = result$p1[1]
    a_normal = 0.02
    a_normal1 = 0.01
    
    Qplot = seq(min(data_table$CO[ex_analyze], na.rm = TRUE), 
                max(data_table$CO[ex_analyze], na.rm = TRUE), 0.001 )
    
    
    pw_linear = lm(PAWP ~ CO, data = data_table)
    Pwplot = pw_linear$coefficients[2]*Qplot + pw_linear$coefficients[1]
    
    model_data = data.frame(Qplot = Qplot, 
                            PWplot = Pwplot)
    ## initialize Ro for the plots
    Ro_plot = data_table$Ro[1]
    
    
    result_ro = optimx(p = c(Ro_plot), fn=alpha_fit_ro, Pw = data_table_fit$PAWP[ex_analyze[1]] , 
                    Q = data_table_fit$CO[ex_analyze[1]], 
                    a_norm = a_normal, 
                    Ppa = data_table_fit$mPAP[ex_analyze[1]])
    ro_fit = result_ro$p1[1]
    
    result_ro = optimx(p = c(Ro_plot), fn=alpha_fit_ro, Pw = data_table_fit$PAWP[ex_analyze[1]] , 
                       Q = data_table_fit$CO[ex_analyze[1]], 
                       a_norm = a_normal1, 
                       Ppa = data_table_fit$mPAP[ex_analyze[1]])
    ro_fit1 = result_ro$p1[1]
    
    model_data = model_data %>% 
      mutate(Pplot = 1/a_fit*(((1+a_fit*Pwplot)^5 + 5*a_fit*Ro_plot*Qplot)^0.2-1), 
             Pnormal = 1/a_normal*(((1+a_normal*Pwplot)^5 + 5*a_normal*ro_fit*Qplot)^0.2-1),
             Pnormal1 = 1/a_normal1*(((1+a_normal1*Pwplot)^5 + 5*a_normal1*ro_fit1*Qplot)^0.2-1)) 
    
    # prepare colours
    all_cols <- c(input$in_track_colour_mPAP,
                  input$in_track_colour_PAWP, 
                  input$in_track_colour_RAP, 
                  input$in_track_colour_linehan)
    
    
   p2 =  data_table %>% 
      filter(ex_cond == TRUE) %>% 
      pivot_longer(cols = RAP:PAWP, names_to = "pressure_type", values_to = "pressure") %>%
      ggscatter(x = "CO", y = "pressure", 
                add = "reg.line", 
                color = "pressure_type", 
                palette = all_cols, 
                shape = "pressure_type", 
                xlim = c(0, max(data_table$CO, na.rm = TRUE)*1.3), 
                ylim = c(0, round(max(data_table$mPAP)*1.2)), 
                xlab = "Cardiac Output (L/min)", 
                ylab = "Pressure (mmHg)", 
                #legend = "right", 
                #legend.title = "Pressure Type"
                )  +
      stat_regline_equation(aes(color = pressure_type, label =  paste(after_stat(eq.label), after_stat(adj.rr.label), sep = "~`,`~")),
                            label.x = 0, 
                            label.y = c(max(data_table$mPAP)*1.2, max(data_table$mPAP)*1.1, max(data_table$mPAP)*1), 
                            size = input$in_equationfontsize)+
     theme(legend.title=element_blank(),
           panel.grid=element_blank(),
           panel.border=element_blank(),
           axis.ticks=element_blank(),
           axis.title=element_text(colour="Black"),
           strip.background=element_blank(),
           #strip.text=element_text(size=input$in_monthfontsize),
           legend.position=input$in_legend_position,
           legend.justification=input$in_legend_justification,
           legend.direction=input$in_legend_direction,
           legend.text=element_text(size=input$in_legendfontsize),
           legend.key.size=unit(0.3,"cm"),
           legend.spacing.x=unit(0.2,"cm"))
   
   p3 = p2 + geom_line(aes(x = Qplot, y = Pplot), 
                  #linetype = "dashed", 
                  linewidth = 0.8, 
                  color = input$in_track_colour_linehan, data = model_data)+ 
     annotate(geom = "text", x = max(model_data$Qplot), y = max(model_data$Pplot), 
              label = paste(round(a_fit*100, 3), "%", sep =""), 
              color = "black", 
              size=input$in_equationfontsize, 
              hjust = -0.2, vjust = -0.2) + 
     geom_line(aes(x = Qplot, y = Pnormal), 
               linetype = "dashed", 
               linewidth = 1.2, 
               color = input$in_track_colour_linehan, data = model_data)+ 
     annotate(geom = "text", x = max(model_data$Qplot), y = max(model_data$Pnormal), 
              label = paste(round(a_normal*100, 2), "%", sep =""), 
              color = "black", 
              size=input$in_equationfontsize,
              hjust = -0.2, vjust = -0.2)+ 
     geom_line(aes(x = Qplot, y = Pnormal1), 
               linetype = "dashed", 
               linewidth = 1.2, 
               color = input$in_track_colour_linehan, data = model_data)+ 
     annotate(geom = "text", x = max(model_data$Qplot), y = max(model_data$Pnormal1), 
              label = paste(round(a_normal1*100, 2), "%", sep =""), 
              color = "black", 
              size=input$in_equationfontsize,
              hjust = -0.2, vjust = -0.2)+
     labs(caption = paste("iCPET Calculator, Analysis Vanderpool Lab, Created: ", 
                          as.Date(as.Date(input$in_duration_date_end)), sep = "")) + 
     theme(plot.caption = element_text(size=input$in_equationfontsize+2))
    
    return(p3)
  })
  
  
  ## OUT: out_dataTable  --------------------------------------------------------
  ## prints out the data table of inputed exercise conditions
  
  output$out_dataTable <- renderTable({
    
    num_tracks <- input$in_num_tracks
    data_table = data.frame(ex_cond = logical(),
                            W = double(),
                            RAP = double(),
                            mPAP = double(), 
                            PAWP = double(), 
                            CO = double()
    )
    
    
      for(i in 1:num_tracks){
        if (!is.null(eval(parse(text=paste0("input$in_track_W_",i))))) {
          data_table[i,1] = eval(parse(text=paste0("input$in_track_Ex_",i)))
          data_table[i,2] = eval(parse(text=paste0("input$in_track_W_",i)))
          data_table[i,3] = eval(parse(text=paste0("input$in_track_RAP_",i)))
          data_table[i,4] = eval(parse(text=paste0("input$in_track_mPAP_",i)))
          data_table[i,5] = eval(parse(text=paste0("input$in_track_PAWP_",i)))
          data_table[i,6] = eval(parse(text=paste0("input$in_track_CO_",i)))
        }
      }
    
    data_table
    
  })
  
  ## OUT: out_dataTable  --------------------------------------------------------
  ## prints out the data table of inputed exercise conditions
  
  dn_table <- reactive({
    
    num_tracks <- input$in_num_tracks
    data_table = data.frame(ex_cond = logical(),
                            W = double(),
                            RAP = double(),
                            mPAP = double(), 
                            PAWP = double(), 
                            CO = double()
    )
    
    
    for(i in 1:num_tracks){
      if (!is.null(eval(parse(text=paste0("input$in_track_W_",i))))) {
        data_table[i,1] = eval(parse(text=paste0("input$in_track_Ex_",i)))
        data_table[i,2] = eval(parse(text=paste0("input$in_track_W_",i)))
        data_table[i,3] = eval(parse(text=paste0("input$in_track_RAP_",i)))
        data_table[i,4] = eval(parse(text=paste0("input$in_track_mPAP_",i)))
        data_table[i,5] = eval(parse(text=paste0("input$in_track_PAWP_",i)))
        data_table[i,6] = eval(parse(text=paste0("input$in_track_CO_",i)))
      }
    }
    
    data_table
    
  })
  

  
  
  ## OUT: out_plot ------------------------------------------------------------
  ## plots figure after it has been created before in function  fn_plot
  
  output$out_plot <- renderImage({
    
    shiny::req(fn_plot())
    shiny::req(input$in_height)
    #shiny::req(input$in_width)
    shiny::req(input$in_res)
    shiny::req(input$in_scale)
    
    height <- as.numeric(input$in_height)
    width <- as.numeric(input$in_width)
    res <- as.numeric(input$in_res)
    
    #if(is.na(width)) width <- (store$week*1)+1
    

    p <- fn_plot()
    #suppressMessages(print(p))
    #png_num = dev.copy(png, "calendar_plot2.png", height=height,width=width, units="cm",res=res)
    #dev_num = dev.off()
    ggsave("iCPET_plot.png",p,height=height,width=width,units="cm",dpi=res,type="cairo")
    
    return(list(src="iCPET_plot.png",
                contentType="image/png",
                width=round(((width*res)/2.54)*input$in_scale,0),
                height=round(((height*res)/2.54)*input$in_scale,0),
                alt="iCPET_plot"))
  }, deleteFile = TRUE)
  
  # FN: fn_downloadplotname ----------------------------------------------------
  # creates filename for download plot
  
  fn_downloadplotname <- function()
  {
    return(paste0(input$sub_id,"_iCPET_plot_",  as.Date(input$in_duration_date_end), 
                  format(strptime(input$clientTime, format='%m/%d/%Y, %I:%M:%S %p'), " %H_%M"),  ".",input$in_format))
  }
  
  # FN: fn_downloadcsvname ----------------------------------------------------
  # creates filename for download plot
  
  fn_downloadcsvname <- function()
  {
    return(paste0(input$sub_id,"_iCPET_plot_",  as.Date(input$in_duration_date_end), 
                  format(strptime(input$clientTime, format='%m/%d/%Y, %I:%M:%S %p'), " %H_%M"),".csv"))
  }
  
  
  # FN: fn_downloadxlsname ----------------------------------------------------
  # creates filename for download plot
  
  fn_downloadxlsname <- function()
  {
    return(paste0(input$sub_id,"_iCPET_plot_",  as.Date(input$in_duration_date_end), 
                  format(strptime(input$clientTime, format='%m/%d/%Y, %I:%M:%S %p'), " %H_%M"),".xlsx"))
  }
  
  ## FN: fn_downloadplot -------------------------------------------------
  ## function to download plot  
  
  fn_downloadplot <- function(){
    shiny::req(fn_plot())
    shiny::req(input$in_height)
    shiny::req(input$in_res)
    shiny::req(input$in_scale)
    
    height <- as.numeric(input$in_height)
    width <- as.numeric(input$in_width)
    res <- as.numeric(input$in_res)
    format <- input$in_format
    
    
    #if(is.na(width)) width <- (store$week*1)+1
    
    p <- fn_plot()
    if(format=="pdf" | format=="svg"){
      ggsave(fn_downloadplotname(),p,height=height,width=width,units="cm",dpi=res)
      #embed_fonts(fn_downloadplotname())
    }else{
      ggsave(fn_downloadplotname(),p,height=height,width=width,units="cm",dpi=res,type="cairo")
      
    }
  }
  
  
  ## FN: fn_downloadcsv -------------------------------------------------
  ## function to download csv  
  
  fn_downloadcsv <- function(){
    #shiny::req(fn_plot())

    
    dn_table_csv <- dn_table()
    write.csv(dn_table_csv, fn_downloadcsvname() )
    

  }
  
  
  ## FN: fn_downloadcsv -------------------------------------------------
  ## function to download csv  
  
  fn_downloadxls <- function(){
    #shiny::req(fn_plot())
    
    shiny::req(fn_plot())
    shiny::req(input$in_height)
    shiny::req(input$in_res)
    shiny::req(input$in_scale)
    
    height <- as.numeric(input$in_height)
    width <- as.numeric(input$in_width)
    res <- as.numeric(input$in_res)
    format <- input$in_format
    
   
    
    #if(is.na(width)) width <- (store$week*1)+1
    
    p <- fn_plot()
    if(format=="pdf" | format=="svg"){
      ggsave(fn_downloadplotname(),p,height=height,width=width,units="cm",dpi=res)
      #embed_fonts(fn_downloadplotname())
    }else{
      ggsave(fn_downloadplotname(),p,height=height,width=width,units="cm",dpi=res,type="cairo")

    }

    
    dn_table_csv <- dn_table()
    
    ## Optimize Linehan Equation
    ex_analyze = which(dn_table_csv$ex_cond)
    data_table = dn_table_csv %>%
      mutate(Ro = case_when(ex_cond == TRUE ~ mPAP[ex_analyze[1]]/CO[ex_analyze[1]]),
             PVR =case_when(ex_cond == TRUE ~ (mPAP-PAWP)/CO ))

    
    pw_linear = lm(PAWP ~ CO, data = data_table)
    pw_index_fit = which(is.na(data_table$PAWP))
    
    data_table_fit = data_table
    data_table_fit$PAWP[pw_index_fit] = pw_linear$coefficients[2]*data_table_fit$CO[pw_index_fit] + pw_linear$coefficients[1]
    
    
    ### Seems the Nelder-Mead method is similar to the Matlab Results
    result = optimx(p = c(0.05), fn=alpha_fit, Pw = data_table_fit$PAWP[ex_analyze] ,
                    Q = data_table_fit$CO[ex_analyze],
                    Ro = data_table_fit$Ro[ex_analyze],
                    Ppa = data_table_fit$mPAP[ex_analyze])
    a_fit = result$p1[1]

    pw_linear = lm(PAWP ~ CO, data = data_table)
    pw_slope = pw_linear$coefficients[2]
    pw_intercept = pw_linear$coefficients[1]
    ### adding R2 20230309
    pw_rsq = summary(pw_linear)$r.squared
    pw_rsq_adj = summary(pw_linear)$adj.r.squared

    
    ra_slope = NULL
    ra_intercept = NULL
    ra_rsq = NULL
    ra_rsq_adj = NULL
    ra_index_fit = which(!is.na(data_table$RAP))
    if (length(ra_index_fit) >= 2) {
      ra_linear = lm(RAP ~ CO, data = data_table)
      ra_slope = ra_linear$coefficients[2]
      ra_intercept = ra_linear$coefficients[1]
      ### adding R2 20230309
      ra_rsq = summary(ra_linear)$r.squared
      ra_rsq_adj = summary(ra_linear)$adj.r.squared
    } 
    
    

    pq_linear = lm(mPAP ~ CO, data = data_table)
    pq_slope = pq_linear$coefficients[2]
    pq_intercept = pq_linear$coefficients[1]
    ### adding R2 20230309
    pq_rsq = summary(pq_linear)$r.squared
    pq_rsq_adj = summary(pq_linear)$adj.r.squared
    
    my_workbook <- createWorkbook()
    
    addWorksheet(
      wb = my_workbook,
      sheetName = input$sub_id
    )
    
    setColWidths(
      my_workbook,
      1,
      cols = 1:4,
      widths = c(6, 6, 10, 10)
    )
    
    writeData(
      my_workbook,
      sheet = 1,
      c(
        paste("Subject: ", input$sub_id, sep = ""),
        paste("Analyzed by: ", input$analyst, sep = ""), 
        paste("Date of RHC: ", input$in_duration_date_start, sep = ""), 
        paste("Date Analzyed: ", input$in_duration_date_end, sep = "")
      ),
      startRow = 1,
      startCol = 1
    )
    
    writeData(
      my_workbook,
      sheet = 1,
      c("Alpha: ",
        "Linear: ", 
        "mPAP: ",
        "Pw: ",
        "RAP: "
      ),
      startRow = 7,
      startCol = 1
    )


    writeData(
      my_workbook,
      sheet = 1,
      c("", 
        as.numeric(a_fit),
        "Slope", 
        as.numeric(pq_slope),
        as.numeric(pw_slope),
        as.numeric(ra_slope)
      ),
      startRow = 6,
      startCol = 2
    )

    writeData(
      my_workbook,
      sheet = 1,
      c("Percent",
        as.numeric(a_fit)*100,
        "Intercept", 
        as.numeric(pq_intercept),
        as.numeric(pw_intercept),
        as.numeric(ra_intercept)
      ),
      startRow = 6,
      startCol = 3
    )
    
    ### adding R2 values 20230309
    writeData(
      my_workbook,
      sheet = 1,
      c("",
        "",
        "R square", 
        as.numeric(pq_rsq),
        as.numeric(pw_rsq),
        as.numeric(ra_rsq)
      ),
      startRow = 6,
      startCol = 4
    )
    
    ### adding R2adj values 20230309
    writeData(
      my_workbook,
      sheet = 1,
      c("",
        "",
        "adj R square", 
        as.numeric(pq_rsq_adj),
        as.numeric(pw_rsq_adj),
        as.numeric(ra_rsq_adj)
      ),
      startRow = 6,
      startCol = 5
    )

    #input$clientTime,
    
    addStyle(
      my_workbook,
      sheet = 1,
      style = createStyle(
        #fontSize = 24,
        textDecoration = "bold"
      ),
      rows = 1:4,
      cols = 1
    )
    
    writeData(
      my_workbook,
      sheet = 1,
      data_table,
      startRow = 13,
      startCol = 1
    )
    
    addStyle(
      my_workbook,
      sheet = 1,
      style = createStyle(
        #fgFill = "#1a5bc4",
        halign = "center",
        #fontColour = "#ffffff"
      ),
      rows = 13:(13+dim(dn_table_csv)[1]),
      cols = 2:6,
      gridExpand = TRUE
    )
    
    # addStyle(
    #   my_workbook,
    #   sheet = 1,
    #   style = createStyle(
    #     fgFill = "#7dafff"
    #   ),
    #   rows = 6:10,
    #   cols = 1:4,
    #   gridExpand = TRUE
    # )
    
    
    insertImage(
      my_workbook,
      sheet = 1, 
      fn_downloadplotname(), 
      width = width, 
      height = height, 
      startRow = 6, 
      startCol = 10,
      units = "cm", 
      dpi = res
    
    )
    
    
    saveWorkbook(my_workbook, fn_downloadxlsname())
    
    
    
    
    # write.csv(dn_table_csv, fn_downloadcsvname() )
    
    
  }
  
  
  
  ## DHL: btn_downloadplot ----------------------------------------------------
  ## download handler for downloading plot
  
  output$btn_downloadplot <- downloadHandler(
    filename=fn_downloadplotname,
    content=function(file) {
      fn_downloadplot()
      file.copy(fn_downloadplotname(),file,overwrite=T)
    }
  )
  
  ## Download the csv file --------------------------
  
  
  output$btn_downloadData <- downloadHandler(
    filename = fn_downloadcsvname,
    content=function(file) {
      fn_downloadcsv()
      file.copy(fn_downloadcsvname(),file,overwrite=T)
    }
  )
  
  ## Download the xls file --------------------------
  
  
  output$btn_downloadData_xls <- downloadHandler(
    filename = fn_downloadxlsname,
    content=function(file) {
      fn_downloadxls()
      file.copy(fn_downloadxlsname(),file,overwrite=T)
    }
  )
  
  
  
}

shinyApp(ui=ui, server=server)