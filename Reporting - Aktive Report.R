require(DBI)
require(odbc)
require(openxlsx)
require(tidyr)
require(dplyr)


# Function to establish database connection
GetWSFLAzureConnection <- function() {
  dbConnect(
    odbc(),
    Driver = "SQL Server",
    Server = "heimatau.database.windows.net",
    Database = "WSFL",
    Port = 1433,
    Uid = Sys.getenv("WSNZDBUSER"),
    Pwd = Sys.getenv("WSNZDBPASS")
  )
}
con <- GetWSFLAzureConnection()

ProviderID = 2
query = paste("select s.NSN, s.FirstName, s.LastName, s.PreferredName, e.Description as Ethnicity, s.DateOfBirth,c.CalendarYear,
sd.SchoolName as School,p.Description as Provider,c.Term,co.Description as Competency,  yg.Description as YearGroup, sc.CompetencyStatusID, sc.Date, sy.YearLevelID
FROM StudentCompetency sc
JOIN Student s on (sc.NSN = s.NSN)
JOIN StudentClass cs on (cs.NSN = s.NSN)
JOIN Class c on (c.ClassID = cs.ClassID)
JOIN StudentYearLevel sy on sy.NSN = s.NSN and sy.CalendarYear = c.CalendarYear and sy.Term = c.Term
JOIN SchoolProvider sp on (sp.MOENumber = c.MOENumber and sp.CalendarYear = c.CalendarYear and sp.Term = c.Term)
JOIN Ethnicity e on e.EthnicityID = s.EthnicityID
JOIN MOE_SchoolDirectory sd on sd.MOENumber = c.MOENumber
JOIN Provider p on p.ProviderID = sp.ProviderID
JOIN Competency co on co.CompetencyID = sc.CompetencyID and co.YearGroupID = sc.YearGroupID
JOIN YearGroup yg on co.YearGroupID = yg.YearGroupID
Where sp.providerID =",ProviderID)

df = dbGetQuery(con, query)
# MASTER = df
system("taskkill /IM Excel.exe")
setwd("~/WSFL (all folders code)/WSFL - Jan 2025")
wb <- createWorkbook()

for(u in sort(unique(df$CalendarYear))){
  cat("Processing",u,"\n")
  Competency = dbGetQuery(con, paste("GetRelevantCompetenciesCalendarYear",u))
  
  if(as.numeric(format(Sys.Date(), "%Y"))==u){
    YTD = T
    cat("YTD results\n")
  }else{
    YTD = F
  }
  
  
  
  df_year = df[df$CalendarYear == u, ]
  
  df_year$CompetencyStatus <- ifelse(df_year$CompetencyStatusID == 1, "Completed", "Not Completed")
  
  # Aggregate counts
  agg <- aggregate(
    x = list(Count = rep(1, nrow(df_year))),
    by = list(
      Competency = df_year$Competency,
      YearGroup = df_year$YearGroup,
      YearLevelID = df_year$YearLevelID,
      CompetencyStatus = df_year$CompetencyStatus
    ),
    FUN = sum
  )
  
  # Create a wide-format table using reshape
  wide <- reshape(
    agg,
    idvar = c("Competency", "YearGroup", "CompetencyStatus"),
    timevar = "YearLevelID",
    direction = "wide"
  )
  wide[is.na(wide)] <- 0
  
  names(wide) <- sub("Count\\.", "Year Level ", colnames(wide))
  names(wide)[2:3]=c("Year Group", "Competency Status")
  wide = wide[wide$Competency %in% Competency$CompetencyDesc & wide$`Year Group` %in% Competency$YearGroupDesc,]
  note_tags <- c("*", "**", "***", "†", "‡")  # You can expand this as needed
  note_df <- data.frame(Tag = character(), Note = character(), stringsAsFactors = FALSE)
  
  # Start tagging competencies with a note
  if (any(!is.na(Competency$Note) & Competency$Note != "") & u != 2023) {
    notes <- unique(na.omit(Competency$Note[Competency$Note != ""]))
    print
    for (i in seq_along(notes)) {
      tag <- note_tags[i]
      message <- notes[i]
      
      # Get rows in Competency that match this note
      matches <- Competency[Competency$Note == message, ]
      
      for (j in seq_len(nrow(matches))) {
        comp_desc <- matches$CompetencyDesc[j]
        yg_desc <- matches$YearGroupDesc[j]
        
        # Update the matching rows in wide to include the tag
        rows_to_update <- which(wide$Competency == comp_desc & wide$`Year Group` == yg_desc)
        wide$Competency[rows_to_update] <- paste0(wide$Competency[rows_to_update], tag)
      }
      
      # Append to note list for printing below table
      note_df <- rbind(note_df, data.frame(Tag = tag, Note = message, stringsAsFactors = FALSE))
    }
  }
  sheetname = ifelse(YTD,paste0("Provider Results ", u," (YTD)"), paste0("Provider Results ", u))
  addWorksheet(wb, sheetname)
  
  title_text <- sheetname
  
  writeData(wb, sheetname, title_text, startRow = 1, startCol = 1)
  addStyle(wb, sheet = sheetname, style = createStyle(valign = "center"),
           cols = 1:ncol(wide), rows = 3:(nrow(wide) + 2), gridExpand = TRUE)

  titleStyle <- createStyle(fontSize = 14, textDecoration = "bold")
  addStyle(wb, sheetname, style = titleStyle, rows = 1, cols = 1, gridExpand = TRUE)
  
  
  writeData(wb, sheetname, wide[order(wide$`Year Group`, wide$Competency),], startRow = 3, startCol = 1)
  setColWidths(wb, sheet = sheetname, cols = 1, widths = 40) 
  setColWidths(wb, sheet = sheetname, cols = 3, widths = 17)  
  
  addStyle(wb, sheet = sheetname, style = createStyle(wrapText = TRUE, valign = "top"), cols = 1, 
           rows = 3:(nrow(wide) + 2), gridExpand = TRUE)
  
  addStyle(wb, sheet = sheetname, style = createStyle(textDecoration = "bold"), cols = 1:ncol(wide), 
           rows = 3, gridExpand = TRUE)
  
  if (nrow(note_df)) {
    note_start_row <- 5 + nrow(wide)
    
    # Write "Notes" header
    writeData(wb, sheetname, "Notes", startCol = 1, startRow = note_start_row)
    addStyle(wb, sheetname, createStyle(textDecoration = "bold"), rows = note_start_row, cols = 1, gridExpand = TRUE)
    
    # Write each note line-by-line
    for (i in seq_len(nrow(note_df))) {
      note_line <- paste0(note_df$Tag[i], " ", note_df$Note[i])
      print(note_line)
      writeData(wb, sheetname, note_line, startCol = 1, startRow = note_start_row + i)
    }
  }
  sheetname = ifelse(YTD,paste0(u," Schools (YTD)"), paste0(u, " Schools"))
  
  addWorksheet(wb, sheetname)
  
  title_text = ifelse(YTD,paste0("Schools with data for ", u," (YTD)"), paste0("Schools with data for ", u))
  writeData(wb, sheetname, title_text, startRow = 1, startCol = 1)
  
  titleStyle <- createStyle(fontSize = 14, textDecoration = "bold")
  addStyle(wb, sheetname, style = titleStyle, rows = 1, cols = 1, gridExpand = TRUE)
  
  unique_school_data <- unique(df$School[df$CalendarYear == u])
  writeData(wb, sheetname, unique_school_data[order(unique_school_data)], startRow = 3, startCol = 1)
  
  
  
}
filename = paste0(unique(df$Provider), " Yearly Results (", gsub("-", ".", as.character(Sys.Date())), ").xlsx")
saveWorkbook(wb, filename, overwrite=T)
shell.exec(filename)




