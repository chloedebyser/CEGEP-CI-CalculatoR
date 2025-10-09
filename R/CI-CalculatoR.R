CICalculatoR <- function(input){
  
  # Get data
        # Workbook
  wb <- loadWorkbook(input)
  
        # Courses
  courses <- read.xlsx(wb, sheet = "Courses") %>%
    mutate(Code = trimws(Code),
           `Lecture/Lab/Stage` = trimws(`Lecture/Lab/Stage`)) %>%
    select_if(~sum(!is.na(.)) > 0)
  
        # Teachers
  teachers <- read.xlsx(wb, sheet = "Teachers")
  teacherList <- teachers %>%
    pull(Last.Name) %>%
    trimws() %>%
    unique() %>%
    sort()
  
        # Scenarios
  scenarioList <- colnames(courses) %>%
    {.[grepl("Scenario.", .)]} %>%
    {gsub("Scenario.", "", .)} %>%
    {as.integer(.)} %>%
    unique() %>%
    sort()
  
  # Styles
        # Header
  headerStyle <- createStyle(
    textDecoration = "bold",
    fgFill = "#F2F2F2"
  )
  
        # Totals
  totalStyle <- createStyle(
    textDecoration = "bold"
  )
  
  # Processing
  for(scenario in scenarioList){
    
    courses %<>%
      rename(Teacher.Last.Name = paste0("Scenario.", scenario)) %>%
      mutate(Teacher.Last.Name = trimws(Teacher.Last.Name))
  
    for(semester in unique(trimws(courses$Semester))){
      
      teachers %<>%
        mutate(SemesterCI = NA)
      
      for(teacher in teacherList){
          
        # Get courses taught by that teacher
        courseAllocation <- courses %>%
          filter(Teacher.Last.Name == teacher,
                 Semester == semester)
        
        if(nrow(courseAllocation) == 0){
          
          finalCI <- 0
          
        }else{
        
          # Calculate prep factor
          prepFactor <- courseAllocation %>%
            filter(!`Lecture/Lab/Stage` == "Stage") %>%
            pull(Code) %>%
            unique() %>%
            length() %>%
            {ifelse(. < 3, 0.9, ifelse(. == 3, 1.1, 1.75))}
          
          # Calculate HP
          courseAllocation %<>%
            mutate(CombinedCode = paste(Code, `Lecture/Lab/Stage`, sep="-")) %>%
            mutate(PrepFactorApplicable = replace(1, row_number() != 1, 0), .by=CombinedCode) %>%
            mutate(HP = PrepFactorApplicable*Hours*prepFactor,
                   HC = Hours*1.2,
                   PES = Hours*`Students/section`,
                   CI = HP+HC,
                   StageCI = ifelse(`Lecture/Lab/Stage` == "Stage", (`Students/section`/Nejk)*0.89*40, 0))
          
          # Calculate Total PES
          totalPES <- sum(courseAllocation$PES, na.rm=T)
          
          # Calculate PES CI
          PESCI <- min(415, totalPES)*0.04 + max(0, (totalPES-415))*0.07
          
          # Calculate NES
          labLectureCodes <- courseAllocation %>%
            filter(`Lecture/Lab/Stage` %in% c("Lecture", "Lab")) %>%
            pull(Code) %>%
            unique() %>%
            sort()
          
          NES <- 0
          
          for(code in labLectureCodes){
            
            # Get course information
            course <- courseAllocation %>%
              filter(Code == code)
            
            # Compute lecture NES
            lectureStudents <- course %>%
              filter(`Lecture/Lab/Stage` == "Lecture") %>%
              pull(`Students/section`) %>%
              sum()
            
            lectureHours <- course %>%
              filter(`Lecture/Lab/Stage` == "Lecture") %>%
              pull(Hours) %>%
              sum()
            
            lectureNES <- ifelse(lectureHours < 2,
                                 0,
                                 ifelse(lectureHours < 3,
                                        lectureStudents * 0.8,
                                        lectureStudents))
            
            # Compute lab NES
            labStudents <- course %>%
              filter(`Lecture/Lab/Stage` == "Lab") %>%
              pull(`Students/section`) %>%
              sum()
            
            labHours <- course %>%
              filter(`Lecture/Lab/Stage` == "Lab") %>%
              pull(Hours) %>%
              sum()
            
            extraStudents <- labStudents - lectureStudents
            extraStudents <- ifelse(extraStudents < 0, 0, extraStudents)
            
            labNES <- ifelse(lectureNES == 0,
                             ifelse(labHours < 2,
                                    0,
                                    ifelse(labHours < 3,
                                           labStudents * 0.8,
                                           labStudents)),
                             extraStudents)
            
            # Total course NES
            courseNES <- lectureNES + labNES
            
            # Total NES
            NES <- NES + courseNES
          }
          
          # Calculate NES 75
          NES75 <- ifelse(NES < 75, 0, NES)
          
          # Calculate NES 160
          NES160 <- ifelse((NES-160)<1, 0, NES-160)
          
          # Calculate NES 75 CI
          NES75CI <- NES75*0.01
          
          # Calculate NES 160 CI
          NES160CI <- (NES160^2)*0.1
          
          # Calculate Sub-Total CI
          subTotalCI <- sum(courseAllocation$CI, na.rm=T) + PESCI + NES75CI + NES160CI
          
          # Calculate Stage CI
          stage <- sum(courseAllocation$StageCI, na.rm=T)
          
          # Calculate CI
          finalCI <- subTotalCI + stage
        }
        
        # Add Release CI
        release <- teachers %>%
          filter(Last.Name == teacher) %>%
          pull(paste0(semester, ".Release")) %>%
          {ifelse(is.na(.), 0, .)} %>%
          {.*40}
        
        finalCI <- finalCI + release
        
        # Add CI to teachers table
        teachers %<>%
          mutate(SemesterCI = case_when(Last.Name == teacher ~ finalCI,
                                        .default = SemesterCI))
      }
      
      teachers %<>% rename("{semester}.CI_{scenario}" := SemesterCI)
    }
    
    # Calculate annual CI
    teachers %<>% mutate("Annual.CI_{scenario}" := rowSums(select(., ends_with(paste0(".CI_", scenario)))))
    
    # Remove scenario
    courses %<>% select(-Teacher.Last.Name)
  }
  
  # Format output
  teachers %<>%
    select(-ends_with(".Release")) %>%
    bind_rows(summarise(.,
                        across(where(is.numeric), sum),
                        across(where(is.character), ~"Total")))
  
  # Add tab to input spreadsheet
  removeWorksheet(wb, sheet = "README")
  addWorksheet(wb, "CI Results")
  writeData(wb, sheet = "CI Results", x = teachers)
  
  # Format sheets
  for(sheet in sheets(wb)){
     ncols <- ncol(read.xlsx(wb, sheet))
     addStyle(wb, sheet = sheet, headerStyle, rows = 1, cols = 1:ncols, gridExpand = TRUE)
     setColWidths(wb, sheet, cols = 1:ncols, widths = 14)
     freezePane(wb, sheet, firstRow = TRUE, firstCol = FALSE)
     
     if(sheet == "CI Results"){
       addStyle(wb, sheet = sheet, totalStyle, rows = nrow(teachers)+1, cols = 1:ncols, gridExpand = TRUE)
     }
  }
  
  # Return spreadsheet
  return(wb)
}
