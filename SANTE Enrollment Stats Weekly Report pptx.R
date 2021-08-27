'~~~~~~~~~~~~~~~~~~~~~~~
  Project Name: SANTE
  Program Name: SANTE Enrollment Stats Weekly Report PPTX.R
  Programmer: Jason Bailey
  Date: 22JUL2021
  Purpose: Build slide deck for Weekly report
    
  Edited by:
  Edited date:
  Current Version: 0.1
  Dependencies: 
  SANTE Enrollment Stats.R
  SANTE Location Analysis.R
  
  Other necessary files: none


  Output files: Weekly_Enrollment_Report_MMDDYYYY.pptx"
  ~~~~~~~~~~~~~~~~~~~~~~~'

#General Weekly Reports
source("C:/Users/jason/Documents/R/Project Files/SANTE/R Scripts/SANTE Enrollment Stats.R")
#Location analysis
#source("C:/Users/jason/Documents/R/Project Files/SANTE/R Scripts/SANTE_location_analysis.R")
#Fadima's report
#source("C:/Users/jason/Documents/R/Project Files/SANTE/R Scripts/SANTE Fadima Enrollment Report.R")
#Data Queries
#source("C:/Users/jason/Documents/R/Project Files/SANTE/R Scripts/SANTE Data Queries.R")

library(officer)
my_pres<-read_pptx("C:/Users/jason/Documents/R/Project Files/SANTE/Reports/blank.pptx")
knitr::kable(layout_summary(my_pres))

my_pres<-
  # Load template
  read_pptx("C:/Users/jason/Documents/R/Project Files/SANTE/Reports/blank.pptx") %>%
  # Add a slide
  add_slide(layout="Title Slide", master="Office Theme") %>%
  ph_with(value = c("Weekly SANTE Enrollment/Study Status Report"),
          location = ph_location_type(type = "ctrTitle")) %>%
  ph_with(value = c(paste("Summary of weekly recruitment and outcomes for ", format(Sys.Date(),"%b-%Y"),sep="")),  
          location = ph_location_type(type = "subTitle")) %>%
  # SLIDE 2  
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Enrollment by Day: Mother-Infant Cohort", location = ph_location_type(type = "title")) %>%
  ph_with(value = rr, location = ph_location_type(type = "body")) %>%
  # SLIDE 3
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Enrollment by Day: Infant-Only Cohort", location = ph_location_type(type = "title")) %>%
  ph_with(value = ss, location = ph_location_type(type = "body")) %>%
  # SLIDE 4
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Enrollment by Month: Mother Infant Cohort", location = ph_location_type(type = "title")) %>%
  ph_with(value = pp, location = ph_location_type(type = "body")) %>%
  # SLIDE 5
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Enrollment by Month: Infant-Only Cohort", location = ph_location_type(type = "title")) %>%
  ph_with(value = qq, location = ph_location_type(type = "body")) %>%
  # SLIDE 6
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Total Enrollment & Follow Up", location = ph_location_type(type = "title")) %>%
  ph_with(value = uu, location = ph_location_type(type = "body")) %>%
  # SLIDE 7
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Monthly and Cumulative Enrollment Ratios", location = ph_location_type(type = "title")) %>%
  ph_with(value = tt, location = ph_location_type(type = "body")) %>%
  # SLIDE 8
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Projected Enrollment", location = ph_location_type(type = "title")) %>%
  ph_with(value = MIcohort, location = ph_location_type(type = "body")) %>%
  # SLIDE 9
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Projected Outcomes", location = ph_location_type(type = "title")) %>%
  ph_with(value = Infoutcomes, location = ph_location_type(type = "body")) %>%
  # SLIDE 10
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Overall Enrollment by District, MI Cohort", location = ph_location_type(type = "title")) %>%
  ph_with(value = a, location = ph_location_type(type = "body")) %>%
  # SLIDE 11
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Overall Enrollment by District, IO Cohort", location = ph_location_type(type = "title")) %>%
  ph_with(value = aa, location = ph_location_type(type = "body")) %>%
  # SLIDE 12
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Total Completed Visits", location = ph_location_type(type = "title")) %>%
  ph_with(value = ppq, location = ph_location_type(type = "body")) %>%
  # SLIDE 13
  add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = "Total Completed Visits", location = ph_location_type(type = "title")) %>%
  ph_with(value = ppr, location = ph_location_type(type = "body"))
  

  print(my_pres, target = paste("C:/Users/jason/Documents/R/Project Files/SANTE/Reports/Weekly_Enrollment_Report_",Sys.Date(),".pptx", sep=""))

  
  
  
  # 
  # 
  #  # SLIDE 11
  # add_slide(layout = "Title and Content", master = "Office Theme") %>%
  # ph_with(value = "Overall Enrollment by Area, Koutiala District, MI Cohort", location = ph_location_type(type = "title")) %>%
  # ph_with(value = b, location = ph_location_type(type = "body")) %>%
  # # SLIDE 12
  # add_slide(layout = "Title and Content", master = "Office Theme") %>%
  # ph_with(value = "Overall Enrollment by Area, Kignan District, MI Cohort", location = ph_location_type(type = "title")) %>%
  # ph_with(value = c, location = ph_location_type(type = "body")) %>%
  # # SLIDE 13
  # add_slide(layout = "Title and Content", master = "Office Theme") %>%
  # ph_with(value = "Overall Enrollment by Area, Niena District, MI Cohort", location = ph_location_type(type = "title")) %>%
  # ph_with(value = d, location = ph_location_type(type = "body")) %>%
  # # SLIDE 13
  # add_slide(layout = "Title and Content", master = "Office Theme") %>%
  # ph_with(value = "Weekly Enrollment by Area, MI Cohort", location = ph_location_type(type = "title")) %>%
  # ph_with(value = f, location = ph_location_type(type = "body")) %>%
  # # SLIDE 14
  # add_slide(layout = "Title and Content", master = "Office Theme") %>%
  # ph_with(value = "Weekly Enrollment by Area, IO Cohort", location = ph_location_type(type = "title")) %>%
  # ph_with(value = g, location = ph_location_type(type = "body"))  %>%
  # 
  # 

# #Migration Slides
#   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#   ph_with(value = "Enrollment vs Delivery Location", location = ph_location_type(type = "title")) %>%
#   ph_with(value = i, location = ph_location_type(type = "body")) %>%
#   # SLIDE 16
#   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#   ph_with(value = "Enrollment vs PENTA1 Visit Location", location = ph_location_type(type = "title")) %>%
#   ph_with(value = j, location = ph_location_type(type = "body")) %>%
#   # SLIDE 17
#   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#   ph_with(value = "Enrollment vs PENTA3 Visit Location", location = ph_location_type(type = "title")) %>%
#   ph_with(value = k, location = ph_location_type(type = "body"))  %>%
#   # SLIDE 17
#   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#   ph_with(value = "Enrollment vs Infant Outcome Location", location = ph_location_type(type = "title")) %>%
#   ph_with(value = l, location = ph_location_type(type = "body")) %>%  
#   
#   # SLIDE 18
#   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#   ph_with(value = "Infant Outcome Location by District", location = ph_location_type(type = "title")) %>%
#   ph_with(value = m, location = ph_location_type(type = "body")) %>%    
#   # SLIDE 19
#   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#   ph_with(value = "Infant Outcome Location by Area (Aire)", location = ph_location_type(type = "title")) %>%
#   ph_with(value = n, location = ph_location_type(type = "body")) 

