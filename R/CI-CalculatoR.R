CICalculatoR <- function(input){
  
  # Get data
        # Workbook
  wb <- loadWorkbook(input)
  
        # Courses
  courses <- read.xlsx(wb, sheet = "Courses") %>%
    mutate(Teacher.Last.Name = trimws(Teacher.Last.Name),
           Code = trimws(Code),
           `Lecture/Lab/Stage` = trimws(`Lecture/Lab/Stage`))
  
        # Teachers
  teachers <- read.xlsx(wb, sheet = "Teachers")
  teacherList <- teachers %>%
    pull(Last.Name) %>%
    trimws() %>%
    unique() %>%
    sort()
  
  # Processing
  for(semester in unique(trimws(courses$Semester))){
    
    teachers %<>%
      mutate(SemesterCI = NA)
    
    for(teacher in teacherList){
        
      # Get courses taught by that teacher
      courseAllocation <- courses %>%
        filter(Teacher.Last.Name == teacher,
               Semester == semester)
      
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
      NES <- courseAllocation %>%
        filter(Hours > 2) %>%
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
      
      # Add CI to teachers table
      teachers %<>%
        mutate(SemesterCI = case_when(Last.Name == teacher ~ finalCI,
                                      .default = SemesterCI))
    }
    
    teachers %<>% rename("{semester}.CI" := SemesterCI)
  }
  
  # Calculate annual CI
  teachers %<>% mutate(AnnualCI = rowSums(select(., ends_with(".CI"))))
  return(teachers)
}