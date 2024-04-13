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
  
  # Header style
  headerStyle <- createStyle(
    textDecoration = "bold",
    fgFill = "#F2F2F2"
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
          
          courseAllocation %<>%
            mutate(ID = 1:nrow(.))
        
          # Calculate prep factor
          prepFactor <- courseAllocation %>%
            pull(Code) %>%
            unique() %>%
            length() %>%
            {ifelse(. < 3, 0.9, ifelse(. == 3, 1.1, 1.9))}
          
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
          totalPES <- sum(courseAllocation$PES)
          
          # Calculate PES CI
          PESCI <- min(415, totalPES)*0.04 + max(0, (totalPES-415))*0.07
          
          # Calculate NES
          labLectureCodes <- courseAllocation %>%
            filter(`Lecture/Lab/Stage` %in% c("Lecture", "Lab")) %>%
            pull(Code) %>%
            unique() %>%
            sort()
          
          courseAllocation %<>% mutate(NESFactor = NA)
          
          for(code in labLectureCodes){
            
            course <- courseAllocation %>%
              filter(Code == code)
            
            lecture <- course %>%
              filter(`Lecture/Lab/Stage` == "Lecture") %>%
              mutate(NESFactor = ifelse(Hours > 2, 1, 0))
            
            lab <- course %>%
              filter(`Lecture/Lab/Stage` == "Lab") %>%
              arrange(desc(`Students/section`)) %>%
              mutate(cumStudents = cumsum(`Students/section`),
                     extraLab = (cumStudents > sum(lecture$`Students/section`) | (sum(lecture$NESFactor) == 0)),
                     NESFactor = ifelse(extraLab & (Hours > 2), 1, 0))
            
            course <- bind_rows(lecture, lab)
            
            courseAllocation %<>%
              full_join(., course[,c("ID", "NESFactor")], by=c("ID", "NESFactor")) %>%
              group_by(ID) %>%
              mutate(NESFactor = first(na.omit(NESFactor))) %>%
              filter(!is.na(Code))
          }
          
          NES <- courseAllocation %>%
            filter(NESFactor == 1) %>%
            pull(`Students/section`) %>%
            sum()
          
          # Calculate NES 75
          NES75 <- ifelse(NES < 75, 0, NES)
          
          # Calculate NES 160
          NES160 <- ifelse((NES-160)<1, 0, NES-160)
          
          # Calculate NES 75 CI
          NES75CI <- NES75*0.01
          
          # Calculate NES 160 CI
          NES160CI <- (NES160^2)*0.1
          
          # Calculate Sub-Total CI
          subTotalCI <- sum(courseAllocation$CI) + PESCI + NES75CI + NES160CI
          
          # Calculate Stage CI
          stage <- sum(courseAllocation$StageCI)
          
          # Calculate Release CI
          release <- teachers %>%
            filter(Last.Name == teacher) %>%
            pull(paste0(semester, ".Release")) %>%
            {ifelse(is.na(.), 0, .)} %>%
            {.*40}
          
          # Calculate CI
          finalCI <- subTotalCI + stage + release
        }
        
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
  teachers %<>% select(-ends_with(".Release"))
  
  # Add tab to input spreadsheet
  addWorksheet(wb, "CI Results")
  writeData(wb, sheet = "CI Results", x = teachers)
  
  # Format sheets
  for(sheet in sheets(wb)){
    
    if(!sheet == "README"){
      ncols <- ncol(read.xlsx(wb, sheet))
      addStyle(wb, sheet = sheet, headerStyle, rows = 1, cols = 1:ncols, gridExpand = TRUE)
      setColWidths(wb, sheet, cols = 1:ncols, widths = 14)
      freezePane(wb, sheet, firstRow = TRUE, firstCol = FALSE)
    }
  }
  
  # Return spreadsheet
  return(wb)
}
