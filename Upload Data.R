require(readxl)
require(DBI)
require(odbc)
require(utils)
require(stringdist)
require(dplyr)
require(tidyr)
require(lubridate)


# Function to establish database connection
GetWSNZAzureConnection <- function() {
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

escape_single_quotes <- function(x) {
  gsub("'", "''", x)
}

# Function to determine folder path
GetFolderPath <- function(CALENDARYEAR, TERM) {
  if (CALENDARYEAR == 2023 & TERM == 3) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2023 Providers School Class Lists etc/Term 3 Schools 2023"
  } else if (CALENDARYEAR == 2023 & TERM == 4) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2023 Providers School Class Lists etc/Term 4 Schools 2023"
  } else if (CALENDARYEAR == 2024 & TERM == 1) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2023 Providers School Class Lists etc/Term 1 Schools 2024"
  } else if (CALENDARYEAR == 2024 & TERM == 2) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2023 Providers School Class Lists etc/Term 2 Schools 2024"
  } else if (CALENDARYEAR == 2024 & TERM == 3) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2024_25 Providers School Class Lists etc/Term 3 Schools 2024"
  } else if (CALENDARYEAR == 2024 & TERM == 4) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2024_25 Providers School Class Lists etc/Term 4 Schools 2024"
  } else if (CALENDARYEAR == 2025 & TERM == 1) {
    "C:/Users/stella/Water Safety New Zealand/WSNZ Team - Documents/WSNZ/Interventions/Funding/WSFL Databases June 23/WSFL 2024_25 Providers School Class Lists etc/Term 1 Schools 2025"
  } else {
    stop(
      paste0(
        "No path specified for selected Term (",
        TERM,
        ") and Year (",
        CALENDARYEAR,
        ")"
      )
    )
  }
}

GetRelevantCompetencies <- function(CurrentDate) {
  # use sql to select all rows in competency where date is between startdate and enddate
  # some will have NULL but those are still active
  # so if both are not NULL check between the two dates
  # if enddate is NULL check if CurrentDate is greater than the start date
  # if any rows have start date as NULL throw an error
  # connection is stored as con
  if (!inherits(CurrentDate, "Date")) {
    stop("CurrentDate must be a Date object")
  }
  query <- "SELECT *
            FROM Competency
            WHERE
              StartDate IS NOT NULL AND
              (
                (EndDate IS NOT NULL AND ? BETWEEN StartDate AND EndDate) OR
                (EndDate IS NULL AND StartDate <= ?)
              )"
  
  # Execute the query and fetch the data
  c <- dbGetQuery(con, query, params = list(CurrentDate, CurrentDate))
  # Check for rows with NULL StartDate (error if none are returned from the main query)
  if (nrow(c) == 0) {
    stop("Error: Some rows have NULL StartDate. Please verify the data.")
  }
  
  return(c)
}

GetDefaultDate <- function(CALENDARYEAR, TERM) {
  if (CALENDARYEAR == 2023 & TERM == 3) {
    start_date <- as.Date("2023-07-17") # Example start date for Term 3, 2023
    end_date <- as.Date("2023-09-22")   # Example end date for Term 3, 2023
  } else if (CALENDARYEAR == 2023 & TERM == 4) {
    start_date <- as.Date("2023-10-09") # Example start date for Term 4, 2023
    end_date <- as.Date("2023-12-15")   # Example end date for Term 4, 2023
  } else if (CALENDARYEAR == 2024 & TERM == 1) {
    start_date <- as.Date("2024-02-12") # Example start date for Term 1, 2024
    end_date <- as.Date("2024-04-12")   # Example end date for Term 1, 2024
  } else if (CALENDARYEAR == 2024 & TERM == 2) {
    start_date <- as.Date("2024-04-29") # Example start date for Term 2, 2024
    end_date <- as.Date("2024-07-05")   # Example end date for Term 2, 2024
  } else if (CALENDARYEAR == 2024 & TERM == 3) {
    start_date <- as.Date("2024-07-22") # Example start date for Term 3, 2024
    end_date <- as.Date("2024-09-27")   # Example end date for Term 3, 2024
  } else if (CALENDARYEAR == 2024 & TERM == 4) {
    start_date <- as.Date("2024-10-06") # Example start date for Term 3, 2024
    end_date <- as.Date("2024-12-19")   # Example end date for Term 3, 2024
  }  else if (CALENDARYEAR == 2025 & TERM == 1) {
    start_date <- as.Date("2025-01-27") # Example start date for Term 3, 2024
    end_date <- as.Date("2025-04-11")   # Example end date for Term 3, 2024
  } else {
    stop(
      paste0(
        "No dates specified for selected Term (",
        TERM,
        ") and Year (",
        CALENDARYEAR,
        ")"
      )
    )
  }
  
  # Calculate the midpoint
  midpoint <- as.Date((as.numeric(start_date) + as.numeric(end_date)) / 2, origin = "1970-01-01")
  return(midpoint)
}

# Function to insert a new provider
InsertProvider <- function(con, ProviderName) {
  query <- paste0("INSERT INTO PROVIDER(ProviderID, Description) VALUES (",
                  ", '",
                  ProviderName,
                  "')")
  dbExecute(con, query)
  cat("Inserted new provider:", ProviderName, "\n")
}

# Function to extract MOE Number from an Excel file
ExtractMOENumber <- function(filename) {
  tryCatch({
    suppressMessages({
      RawSheet <- read_excel(filename, col_names = FALSE)
    })
    RowMOE <- which(RawSheet[[1]] == "MOE Number:")
    if (length(RowMOE) == 0)
      return(NA)
    suppressMessages({
      read_excel(filename,
                 range = paste0("B", RowMOE, ":B", RowMOE),
                 col_names = F)[[1]]
    })
  }, error = function(e) {
    FullFileName <- file.path(getwd(), filename)
    cat("Error processing file:", filename, "\n")
    cat(FullFileName)
    ErrorMessage <- "'MOE Number not found"
    # WriteError(con,
    #            "CLA",
    #            as.integer(99999),
    #            as.integer(99999),
    #            FullFileName,
    #            ErrorMessage)
    
    return(NA)
  })
}

# Function to process each school folder
CheckMOENumber <- function(s, ProviderID) {
  con <- GetWSNZAzureConnection()
  
  FinalMOENumber <- NA
  print(s)
  setwd(s)
  cat("\n\nWorking Directory set as:\n", s, "\n")
  
  ExcelFiles <- list.files(pattern = "\\.(xlsm|xlsx)$")
  if (length(ExcelFiles) == 0) {
    cat("No Excel files found in directory:", s, "\n")
    return(NULL)
  }
  
  UniqueMOENumbers <- unique(na.omit(sapply(ExcelFiles, ExtractMOENumber)))
  print(UniqueMOENumbers)
  if (length(UniqueMOENumbers) == 0) {
    stop("No MOE Number found for", sub(".*\\/", "", s))
  } else if (length(UniqueMOENumbers) > 1) {
    matched_schools <- MOE_SchoolDirectory[MOE_SchoolDirectory$MOENumber %in% UniqueMOENumbers, ]
    if (nrow(matched_schools) == 0) {
      stop(paste("Invalid MOE Numbers", UniqueMOENumbers))
    }
    
    directory_name <- sub(".*\\/", "", s)
    similarity_scores <- sapply(matched_schools$SchoolName, function(name) {
      adist(tolower(name), tolower(directory_name))
    })
    best_MatchIndex <- which.min(similarity_scores)
    if (length(best_MatchIndex) == 1) {
      chosen_moe <- matched_schools$MOENumber[best_MatchIndex]
      chosen_school <- matched_schools$SchoolName[best_MatchIndex]
      cat("Automatically selected MOE Number",
          chosen_moe,
          "for ",
          chosen_school,
          "\n")
      FinalMOENumber <- chosen_moe
    } else {
      stop(paste(
        "Could not determine the correct MOE Number for",
        sub(".*\\/", "", s),
        "\n",
        paste(apply(matched_schools, 1, function(row)
          paste0(row['MOENumber'], " (", row['SchoolName'], ")")), collapse = ", ")
      ))
    }
  } else {
    name <- MOE_SchoolDirectory$SchoolName[MOE_SchoolDirectory$MOENumber == UniqueMOENumbers]
    if (length(name) < 1) {
      stop(paste("Invalid MOE Number", UniqueMOENumbers))
    }
    cat("MOE Number found:", UniqueMOENumbers, "(", name, ")\n")
    FinalMOENumber <- UniqueMOENumbers
  }
  
  
  if (nrow(ProviderEducationRegion[ProviderEducationRegion$ProviderID == ProviderID &
                                   ProviderEducationRegion$EducationRegionID == MOE_SchoolDirectory$EducationRegionID[MOE_SchoolDirectory$MOENumber == FinalMOENumber], ]) == 0) {
    if (!is.na(FinalMOENumer) && FinalMOENumber != 498)
      stop("Invalid MOE Number")
  }
  
  query <- paste0(
    "
   MERGE INTO SchoolProvider AS target
      USING (VALUES (",
    FinalMOENumber,
    ", ",
    TERM,
    ", ",
    CALENDARYEAR,
    ", ",
    ProviderID,
    "))
   AS source (MOENumber, Term, CalendarYear, ProviderID)
  ON target.MOENumber = source.MOENumber
     AND target.Term = source.Term
     AND target.CalendarYear = source.CalendarYear
  WHEN NOT MATCHED THEN
     INSERT (MOENumber, Term, CalendarYear, ProviderID)
     VALUES (source.MOENumber, source.Term, source.CalendarYear, source.ProviderID);
  "
  )
  dbExecute(con, query)
  cat(
    "Checked and conditionally inserted SchoolProvider entry for MOENumber:",
    FinalMOENumber,
    "\n"
  )
  
  
  
  ReadClassData(s, FinalMOENumber)
  return(NULL)
}

# Function to process each provider folder
ProcessProvider <- function(p, Provider, con, SchoolIndexRange) {
  ProviderName <- sub(".*\\/", "", p)
  
  if (grepl("missing|incomplete", ProviderName))
    return(NULL)
  if (CALENDARYEAR == 2024 & TERM == 1) {
    ProviderName  = sub(" Term 1 Student Achievements", "", ProviderName)
    if (ProviderName == "Christchurch Council") {
      ProviderName = "Christchurch City Council"
    }
    
    
  }
  if (ProviderName == "Aquatic skills") {
    ProviderName = "Aquatic Skills (Tauranga)"
  }
  if (ProviderName %in% c("Christchurch Council", "Christchurch")) {
    ProviderName = "Christchurch City Council"
  }
  
  matching_provider <- Provider[tolower(Provider$Description) == tolower(ProviderName), ]
  
  if (nrow(matching_provider) == 0) {
    cat("No matching provider found for", ProviderName, "\n")
    InsertProvider(con, ProviderName)
    return(NULL) # Add return here to avoid errors if ProviderID is not set
  } else {
    ProviderID <- matching_provider$ProviderID[1]
    cat("Processing existing provider:",
        ProviderName,
        "(ProviderID:",
        ProviderID,
        ")\n")
  }
  
  FolderSchool <- list.dirs(p, recursive = FALSE)
  invisible(sapply(FolderSchool[SchoolIndexRange], function(s) {
    CheckMOENumber(s, ProviderID) # Explicitly pass ProviderID
  }))
  return(NULL)  # Ensure no unintended return values
}

ReadClassData <- function(Directory, MOENumber) {
  # Get all class files in the directory
  ClassFiles <- list.files(Directory, recursive = FALSE, pattern = "\\.(xlsm|xlsx)$")
  cat("\n", length(ClassFiles), "Classes Found\n\n")
  
  for (index in seq_along(ClassFiles)) {
    File <- ClassFiles[index]
    FullFileName <- file.path(Directory, File)  # Full path of the class file
    cat("Processing File:", FullFileName, "\n")
    
    # Initialize variables
    ClassID <- NULL
    
    # Load Class table from database
    Class <- dbGetQuery(con, "SELECT * FROM Class")
    
    # Read the raw Excel sheet
    suppressMessages({
      RawSheet <- read_excel(FullFileName, col_names = FALSE)
      
    })
    
    
    rawwww <<- RawSheet
    # Identify key rows for School Name, Teacher Name, and Class Name
    RowSchoolName <- which(RawSheet$...1 == "School Name:")
    RowTeacherName <- which(RawSheet$...4 == "Teacher Name:")
    RowClassName <- which(RawSheet$...4 == "Class Name:")
    
    
    
    
    # Check for 'School Name:' row
    if (length(RowSchoolName) == 1) {
      SchoolName <- RawSheet$...2[RowSchoolName]  # Assuming School Name is in the second column
    } else {
      SchoolName <- NULL
      ErrorMessage <- "'School Name:' not found"
      WriteError(con, "CLA", NULL, index, FullFileName, ErrorMessage)
      
      next
    }
    
    
    # Check for 'Teacher Name:' row
    if (length(RowTeacherName) == 1) {
      TeacherName <- RawSheet$...5[RowTeacherName]  # Assuming Teacher Name is in the fifth column
    } else {
      TeacherName <- NULL
      ErrorMessage <- "'Teacher Name:' not found"
      WriteError(con, "CLA", NULL, index, FullFileName, ErrorMessage)
      
      next
    }
    
    # Check for 'Class Name:' row
    if (length(RowClassName) == 1) {
      ClassName <- RawSheet$...5[RowClassName]  # Assuming Class Name is in the fifth column
    } else {
      ClassName <- NULL
      ErrorMessage <- "'Class Name:' not found"
      WriteError(con, "CLA", NULL, index, FullFileName, ErrorMessage)
      
      next
    }
    # Log extracted details
    cat("Extracted Data:\n")
    cat("  School Name:", SchoolName, "\n")
    cat("  Teacher Name:", TeacherName, "\n")
    cat("  Class Name:", ClassName, "\n")
    
    # Check if ClassID already exists
    query <- paste0(
      "SELECT ClassID FROM Class WHERE ",
      "FilePath = '",
      escape_single_quotes(FullFileName),
      "' OR ",
      "(Term = ",
      TERM,
      " AND CalendarYear = ",
      CALENDARYEAR,
      " AND MOENumber = ",
      MOENumber,
      " AND (TeacherName = '",
      escape_single_quotes(TeacherName),
      "')",
      " AND (ClassName = '",
      escape_single_quotes(ClassName),
      "'))"
    )
    
    # Execute the query
    ExistingClass <- dbGetQuery(con, query)
    
    if (nrow(ExistingClass) == 1) {
      ClassID <- ExistingClass$ClassID
      cat("Class found in database. ClassID:", ClassID, "\n")
      
      query <- paste0(
        "SELECT 1 FROM ClassFilePath WHERE ClassID = ",
        ClassID,
        " AND FilePath = '",
        escape_single_quotes(FullFileName),
        "'"
      )
      ExistsInClassFilePath <- dbGetQuery(con, query)
      
      if (nrow(ExistsInClassFilePath) == 0) {
        # Insert into ClassFilePath if not exists
        query <- paste0(
          "INSERT INTO ClassFilePath (ClassID, FilePath) VALUES (",
          ClassID,
          ", '",
          escape_single_quotes(FullFileName),
          "')"
        )
        dbExecute(con, query)
        cat("ClassFilePath entry added.\n")
      } else {
        cat("ClassFilePath entry already exists.\n")
      }
      
    } else if (nrow(ExistingClass) > 1) {
      print(ExistingClass)
      stop("More than 1 existing class")
    } else {
      # Insert a new class entry if no match found
      c = dbGetQuery(con, "select max(ClassID) from Class")
      query <- paste0(
        "INSERT INTO Class (ClassID, ClassName, TeacherName, MOENumber, CalendarYear, Term, FilePath) ",
        "VALUES (",
        c + 1,
        ", ",
        ifelse(
          is.na(ClassName),
          "NULL",
          paste0("'", escape_single_quotes(ClassName), "'")
        ),
        ", ",
        ifelse(
          is.na(TeacherName),
          "NULL",
          paste0("'", escape_single_quotes(TeacherName), "'")
        ),
        ", ",
        MOENumber,
        ", ",
        CALENDARYEAR,
        ", ",
        TERM,
        ", '",
        escape_single_quotes(FullFileName),
        "')"
      )
      
      dbExecute(con, query)
      
      # Retrieve the new ClassID
      ClassID <- dbGetQuery(
        con,
        paste0(
          "SELECT ClassID FROM Class WHERE FilePath = '",
          escape_single_quotes(FullFileName),
          "'"
        )
      )$ClassID
      #cat("New class inserted into database. ClassID:", ClassID, "\n")
      
      query <- paste0(
        "INSERT INTO ClassFilePath (ClassID, FilePath) VALUES (",
        ClassID,
        ", '",
        escape_single_quotes(FullFileName),
        "')"
      )
      dbExecute(con, query)
      cat("ClassFilePath entry added.\n")
    }
    
    # Validate ClassID
    if (is.null(ClassID)) {
      ErrorMessage <- "ClassID could not be determined or inserted"
      WriteError(con, "CLA", NULL, index, FullFileName, ErrorMessage)
      
    }
    
    # Read student data
    StudentDataStart <- which(RawSheet$...1 == "NSN")
    if (length(StudentDataStart) == 0) {
      ErrorMessage <- "NSN column not found"
      WriteError(con, "CLA", ClassID, index, FullFileName, ErrorMessage)
      
      next
    }
    suppressMessages({
      Students <- read_excel(FullFileName, skip = StudentDataStart - 1, col_names = TRUE)
    })
    
    
    
    if (ncol(RawSheet) == 42) {
      NewColNames <- as.character(RawSheet[1, 8:38]) # Use only columns 8 to 38
      #print(NewColNames)
      if (TERM %in% c(3, 4) &&
          CALENDARYEAR == 2024  | TERM %in% c(1) && CALENDARYEAR == 2025) {
        # Update column names if specific conditions are met
        NewColNames[NewColNames == "Perform safety of self and others sequence in deep water"] <-
          "Perform personal buoyancy sequence - Signal for help"
        NewColNames[NewColNames == "Perform personal buoyancy sequence - lifejacket in deep water"] <-
          "Perform safety of self and others sequence in deep water"
      }
      
      colnames(Students)[8:38] <- NewColNames # Assign updated names
    } else if (ncol(RawSheet) > 42) {
      stop("Unwanted competencies ")
    } else {
      NewColNames <- as.character(RawSheet[1, 8:ncol(RawSheet)]) # Use all remaining columns
      
      if (TERM %in% c(3, 4) &&
          CALENDARYEAR == 2024 | TERM %in% c(1) && CALENDARYEAR == 2025) {
        # Update column names if specific conditions are met
        NewColNames[NewColNames == "Perform safety of self and others sequence in deep water"] <-
          "Perform personal buoyancy sequence - Signal for help"
        NewColNames[NewColNames == "Perform personal buoyancy sequence - lifejacket in deep water"] <-
          "Perform safety of self and others sequence in deep water"
      }
      
      colnames(Students)[8:ncol(Students)] <- NewColNames # Assign updated names
    }
    
    # Save updated column names to global variable
    ncn <<- NewColNames
    
    
    sssss <<- Students
    
    # in students if colname is 'Perform personal buoyancy sequence - lifejacket in deep water'
    # make it 'Perform safety of self and others sequence in deep water'
    # in students if colname is 'Perform safety of self and others sequence in deep water'
    # make it 'Perform personal buoyancy sequence - Signal for help'
    
    
    rel_competncies <- GetRelevantCompetencies(as.Date(GetDefaultDate(CALENDARYEAR, TERM)))
    
    rc <<- rel_competncies
    # print("Fetched Relevant Competencies from Database:")
    # print(rc)
    
    # Normalize text function for consistent comparisons
    normalize_text <- function(x) {
      tolower(gsub("\\s+", " ", trimws(x)))
    }
    
    # Normalize both NewColNames and rel_competncies descriptions
    normalized_colnames <- sapply(NewColNames, normalize_text)
    normalized_descriptions <- sapply(rel_competncies$Description, normalize_text)
    
    # Filter out NA values
    filtered_colnames <- normalized_colnames[!is.na(normalized_colnames)]
    filtered_descriptions <- normalized_descriptions[!is.na(normalized_descriptions)]
    filtered_colnames <- na.omit(normalized_colnames)
    filtered_descriptions <- na.omit(normalized_descriptions)
    
    # Identify mismatched competencies
    mismatched_descriptions <- setdiff(filtered_colnames, filtered_descriptions)
    mismatched_descriptions = na.omit(mismatched_descriptions)
    mismatched_descriptions <- mismatched_descriptions[mismatched_descriptions != "na"] # Remove string "na"
    
    # Log mismatched descriptions if any
    if (!all(is.na(mismatched_descriptions)) &&
        length(mismatched_descriptions) > 0) {
      ErrorMessage <- paste(
        "Competencies in the template do not match the database:",
        paste(mismatched_descriptions, collapse = ", ")
      )
      print(ErrorMessage)
      WriteError(con, "CLA", NULL, index, FullFileName, ErrorMessage)
      
      print("Mismatched Descriptions:")
      print(mismatched_descriptions)
      
      # Proceed to the next file in case of mismatches
      next
    } else {
      print("Relevant Competencies Found")
    }
    
    if (any(
      RawSheet == "Demonstrate use of multiple skills to respond to two different scenarios",
      na.rm = TRUE
    )) {
      # Identify the last 4 columns of Students
      last4_col_indices <- (ncol(Students) - 3):ncol(Students)
      
      # Rename the last 4 columns with the specified names
      colnames(Students)[last4_col_indices] <- c(
        "Scenario One - Chosen Option",
        "Scenario One - Completed",
        "Scenario Two - Chosen Option",
        "Scenario Two - Completed"
      )
      
      cat("Renamed last 4 columns to match scenario descriptions.\n")
      
      
      
      # if there is a cell wiwth "Demonstrate use of multiple skills
      # to respond to two different scenarios" rename the last 4 columns with
      #Scenario One - Chosen Option	Scenario One - Completed	Scenario Two -  Chosen Option	Scenario Two -  Completed
      
    }
    
    ss <<- Students
    
    
    
    
    
    CheckStudentData(Students,
                     ClassID,
                     FullFileName,
                     MOENumber,
                     rel_competncies)
  }
}

WriteError <- function(con,
                       errorType,
                       errorID,
                       errorIndex,
                       errorFilePath,
                       errorMessage) {
  query <- "
    MERGE INTO Error AS target
USING (VALUES (?, ?, ?, ?, ?, ?)) AS source (Type, ID, [Index], FilePath, Error, DateTimeStamp)
ON target.Type = source.Type
   AND (target.ID = source.ID OR (target.ID IS NULL AND source.ID IS NULL))
   AND target.[Index] = source.[Index]
   AND target.FilePath = source.FilePath
   AND target.Error = source.Error
WHEN MATCHED THEN
    UPDATE SET DateTimeStamp = source.DateTimeStamp
WHEN NOT MATCHED THEN
    INSERT (Type, ID, [Index], FilePath, Error, DateTimeStamp)
    VALUES (source.Type, source.ID, source.[Index], source.FilePath, source.Error, source.DateTimeStamp);
"
  
  # Execute the MERGE query
  dbExecute(
    con,
    query,
    params = list(
      errorType,
      errorID,
      errorIndex,
      errorFilePath,
      errorMessage,
      format(Sys.time(), tz = "UTC", usetz = FALSE)
    )
  )
}

validate_nsn <- function(nsn) {
  #cat(" Validating NSN ",nsn)
  if (is.na(nsn)) {
    return("NSN is missing or not numeric.")
  }
  if (nchar(as.character(nsn)) < 8) {
    return("NSN length is invalid.")
  }
  cat("\nValidating NSN", nsn, "\n")
  return(NULL)  # No errors
}

library(lubridate)

ValidateBirthDate <- function(birth_date) {
  # Convert to character and remove time component
  birth_date <- as.character(birth_date)
  birth_date <- sub(" .*", "", birth_date)  # Strip time if present
  
  # Return NA if input is empty or NA
  if (is.na(birth_date) || trimws(birth_date) == "") {
    return(NA)
  }
  
  # Try multiple common date formats
  parsed <- suppressWarnings(parse_date_time(
    birth_date,
    orders = c("ymd", "dmy", "mdy", "d-b-Y", "Y/m/d", "d.m.Y", "B d, Y", "d B Y")
  ))
  
  # Return parsed date or NA
  if (is.na(parsed)) {
    return(NA)
  }
  
  return(as.Date(parsed))
}


normalize_name <- function(name) {
  gsub("[^a-z ]", "", tolower(name))  # Includes a space in the regex pattern
}
normalize_text <- function(x) {
  tolower(gsub("[[:space:][:punct:]]+", " ", x))
}
# Updated CheckNSNExists Function
CheckNSNExists <- function(row) {
  ss <- dbGetQuery(con, paste("select * from Student where NSN = ", row$NSN))
  
  # Normalize names in both datasets
  row$`First Name` <- normalize_name(row$`First Name`)
  row$`Preferred Name` <- normalize_name(row$`Preferred Name`)
  row$`Family Name` <- normalize_name(row$`Family Name`)
  #row$`Birth Date` <- as.Date(row$`Birth Date`)
  ss$FirstName <- normalize_name(ss$FirstName)
  ss$PreferredName <- normalize_name(ss$PreferredName)
  ss$LastName <- normalize_name(ss$LastName)
  ss$DateOfBirth <- as.Date(ss$DateOfBirth)
  rrr <<- row
  sss <<- ss
  
  # Check if NSN exists in the Student dataset
  if (as.integer(row$NSN) %in% ss$NSN) {
    # Get the index of the matching NSN
    MatchIndex <- which(ss$NSN == as.integer(row$NSN))
    
    # Check First Name, Last Name, or Preferred Name with Damerau-Levenshtein Distance
    FirstNameMatch <- !is.na(stringdist(row$`First Name`, ss$FirstName[MatchIndex], method = "dl")) &&
      stringdist(row$`First Name`, ss$FirstName[MatchIndex], method = "dl") <= 2
    
    PreferredNameMatch <- !is.na(stringdist(row$`Preferred Name`, ss$PreferredName[MatchIndex], method = "dl")) &&
      stringdist(row$`Preferred Name`, ss$PreferredName[MatchIndex], method = "dl") <= 2
    
    FamilyNameMatch <- !is.na(stringdist(row$`Family Name`, ss$LastName[MatchIndex], method = "dl")) &&
      stringdist(row$`Family Name`, ss$LastName[MatchIndex], method = "dl") <= 2
    
    # Check if DOB is the same
    DOBMatch <- !is.na(row$`Birth Date`) &&
      !is.na(ss$DateOfBirth[MatchIndex]) &&
      row$`Birth Date` == ss$DateOfBirth[MatchIndex]
    
    # Logic for matching
    if (isTRUE(row$`First Name` == ss$FirstName[MatchIndex])) {
      # First Name matches, check Last Name or DOB
      if (isTRUE(FamilyNameMatch) || isTRUE(DOBMatch)) {
        return(TRUE)
      } else if (isTRUE(DOBMatch)) {
        return(TRUE)
      }
    } else if (isTRUE(row$`Preferred Name` == ss$PreferredName[MatchIndex])) {
      # Preferred Name matches, check Last Name or DOB
      if (isTRUE(FamilyNameMatch) || isTRUE(DOBMatch)) {
        return(TRUE)
      }
    } else if (isTRUE(row$`Family Name` == ss$LastName[MatchIndex])) {
      # Family Name matches, check First Name or Preferred Name or DOB
      if (isTRUE(FirstNameMatch) ||
          isTRUE(PreferredNameMatch) || isTRUE(DOBMatch)) {
        return(TRUE)
      }
    } else if (isTRUE(row$`First Name` == ss$`Preferred Name`[MatchIndex])) {
      if (isTRUE(DOBMatch)) {
        return(TRUE)
      }
    }
  }
  
  return(FALSE)
}

# Main CheckStudentData Function
CheckStudentData <- function(dataframe,
                             ClassID,
                             FileName,
                             MOENumber,
                             RelevantCompetencies) {
  cat("Checking student data for ClassID:", ClassID, "\n")
  cat(nrow(dataframe), " Students Found \n\n")
  con <- GetWSNZAzureConnection()
  
  for (i in 1:nrow(dataframe)) {
    row <- dataframe[i, ]
    ro <<- row
    # Validate NSN
    row$NSN <- gsub("[^0-9]", "", row$NSN)
    row$NSN = as.numeric(row$NSN)
    NSNError <- validate_nsn(row$NSN)
    
    if (!is.null(NSNError)) {
      ErrorMessage <- paste("Unable to insert:", NSNError)
      
      WriteError(con,
                 "STU",
                 as.integer(row$NSN),
                 i,
                 FileName,
                 ErrorMessage)
      
      next
    }
    
    if (is.na(row$`Family Name`)) {
      query <- "INSERT INTO Error(Type, ID, Index, Error, FilePath, DateTimeStamp) VALUES(?, ?, ?, ?, ?,?)"
      ErrorMessage <- paste("Family Name is NA")
      
      WriteError(con,
                 "STU",
                 as.integer(row$NSN),
                 i,
                 FileName,
                 ErrorMessage)
      
      print(ErrorMessage)
      next
    }
    
    if (is.na(row$`First Name`) || trimws(row$`First Name`) == "") {
      # Log error for missing First Name
      ErrorMessage <- "First Name is missing or blank"
      WriteError(con,
                 "STU",
                 as.integer(row$NSN),
                 i,
                 FileName,
                 ErrorMessage)
      
      cat("Error: First Name is missing or blank for NSN:",
          as.integer(row$NSN),
          "\n")
      next  # Skip to the next row
    }
    
    
    if (is.na(row$`Year Group`)) {
      yl =  dbGetQuery(
        con,
        paste(
          "Select TOP(1) * from StudentYearLevel where NSN = ",
          row$NSN,
          "ORDER BY CalendarYear, Term, YearlevelID DESC"
        )
      )
      if (nrow(yl) == 0) {
        ErrorMessage <- paste("Unable to insert as Year Group is NA")
        WriteError(con,
                   "STU",
                   as.integer(row$NSN),
                   i,
                   FileName,
                   ErrorMessage)
        print(ErrorMessage)
        next
      } else{
        diff = year(Sys.Date()) - yl$CalendarYear
        newYL = yl$YearLevelID + diff
        row$`Year Group` = newYL
        print("Inferred year level")
        exists <- dbGetQuery(
          con,
          "
          SELECT 1 FROM InferredStudentYearLevel
          WHERE NSN = ? AND CalendarYear = ? AND Term = ? AND FilePath = ?
          ",
          params = list(as.integer(row$NSN), CALENDARYEAR, TERM, FileName)
        )
        
        if (nrow(exists) == 0) {
          # Insert guessed year level only if not already present
          dbExecute(
            con,
            "
            INSERT INTO InferredStudentYearLevel (NSN, CalendarYear, Term, YearLevelID, FilePath)
            VALUES (?, ?, ?, ?, ?)",
            params = list(
              as.integer(row$NSN),
              CALENDARYEAR,
              TERM,
              newYL,
              FileName
            )
          )
        }
      }
      
      
    }
    
    # Validate Year Group
    # if  row$`Year Group` contains Y or y remove it
    row$`Year Group` <- gsub("y", "", row$`Year Group`, ignore.case = TRUE)
    
    if (nrow(SchoolTypeYearLevel[SchoolTypeYearLevel$YearLevelID == row$`Year Group` &
                                 SchoolTypeYearLevel$SchoolTypeID == (MOE_SchoolDirectory$SchoolTypeID[MOE_SchoolDirectory$MOENumber == MOENumber]), ]) == 0) {
      print(SchoolTypeYearLevel[SchoolTypeYearLevel$YearLevelID == row$`Year Group` &
                                  SchoolTypeYearLevel$SchoolTypeID == (MOE_SchoolDirectory$SchoolTypeID[MOE_SchoolDirectory$MOENumber == MOENumber]), ])
      ErrorMessage <- paste(
        "Unable to insert",
        as.integer(row$NSN),
        "YearLevelID (",
        row$`Year Group`,
        ") to MOENumber",
        MOENumber
      )
      WriteError(con,
                 "STU",
                 as.integer(row$NSN),
                 i,
                 FileName,
                 ErrorMessage)
      
      print(ErrorMessage)
      
      next
    }
    
    # Validate Ethnicity
    TidyEthnicity <- TidyEthnicityValue(row$Ethnicity)
    if (!(TidyEthnicity %in% Ethnicity$Ethnicity)) {
      ErrorMessage <- paste("Unable to insert:", TidyEthnicity)
      WriteError(con,
                 "STU",
                 as.integer(row$NSN),
                 i,
                 FileName,
                 ErrorMessage)
      
      print(ErrorMessage)
      next
    } else {
      row$Ethnicity <- TidyEthnicity
    }
    row = merge(row, Ethnicity, by = "Ethnicity")
    
    # Validate Birth Date
    row$`Birth Date` <- ValidateBirthDate(row$`Birth Date`)
    if (is.na(row$`Birth Date`)) {
      cat("\nDOB was incorrect format. Has been changed to NA\n")
    }
    
    # Handle missing Preferred Name
    if (is.na(row$`Preferred Name`)) {
      row$`Preferred Name` <- row$`First Name`
      print("Preferred Name was NA and has been changed to the given first name")
    }
    
    cat("\nSuccessfully processed record")
    Exists = F
    if (as.integer(row$NSN) %in% Student$NSN) {
      Exists = CheckNSNExists(row)
      if (!Exists) {
        ErrorMessage <- paste("NSN",
                              as.integer(row$NSN),
                              "does not match current record in Student")
        WriteError(con,
                   "STU",
                   as.integer(row$NSN),
                   i,
                   FileName,
                   ErrorMessage)
        print(ErrorMessage)
        next
      }
      
      if (as.integer(row$NSN) %in% Student$NSN) {
        ss <- Student[Student$NSN == as.integer(row$NSN), ]
        
        # If existing DOB is NA and Excel has valid DOB, update Student table
        if (is.na(ss$DateOfBirth) && !is.na(row$`Birth Date`)) {
          update_query <- "
      UPDATE Student
      SET DateOfBirth = ?
      WHERE NSN = ?"
          
          dbExecute(con, update_query, params = list(
            format(row$`Birth Date`, "%Y-%m-%d"),
            as.integer(row$NSN)
          ))
          cat("Updated missing DOB for NSN:", row$NSN, "\n")
          
          # Update the local Student dataframe to avoid updating again for same NSN
          Student$DateOfBirth[Student$NSN == row$NSN] <- row$`Birth Date`
        }
        
        
        
        
      } else {
        query <- "INSERT INTO Student(NSN, FirstName, LastName, PreferredName, DateOfBirth, EthnicityID) VALUES(?, ?, ?, ?, ?,?)"
        dbExecute(con,
                  query,
                  params = list(
                    ifelse(is.null(as.integer(row$NSN)), NA, as.integer(row$NSN)),
                    ifelse(is.null(row$`First Name`), NA, row$`First Name`),
                    ifelse(is.null(row$`Family Name`), NA, row$`Family Name`),
                    ifelse(
                      is.null(row$`Preferred Name`),
                      NA,
                      row$`Preferred Name`
                    ),
                    ifelse(
                      is.null(row$`Birth Date`),
                      NA,
                      format(as.Date(row$`Birth Date`), "%Y-%m-%d")
                    ),
                    ifelse(is.null(row$EthnicityID), NA, row$EthnicityID)
                  ))
        error_exists <- subset(Error,
                               Type == "STU" &
                                 ID == as.integer(row$NSN) &
                                 Index == i &
                                 FilePath == FileName)
        
        
        # If the record exists in the Error table, delete it
        if (nrow(error_exists) > 0) {
          print("removing from error")
          delete_error_query <- "
        DELETE FROM Error
        WHERE Type = ? AND ID = ? AND [Index] = ? AND FilePath = ?
        "
          dbExecute(con,
                    delete_error_query,
                    params = list("STU", as.integer(row$NSN), i, FileName))
          cat(
            "Deleted record from Error table for NSN:",
            as.integer(row$NSN),
            "Index:",
            i,
            "FileName:",
            FileName,
            "\n"
          )
        }
        
      }
      
      
      merge_query <- "
    MERGE INTO StudentClass AS target
    USING (VALUES (?, ?)) AS source (NSN, ClassID)
    ON target.NSN = source.NSN AND target.ClassID = source.ClassID
    WHEN NOT MATCHED THEN
        INSERT (NSN, ClassID)
        VALUES (source.NSN, source.ClassID);
"
      
      dbExecute(con, merge_query, params = list(as.integer(row$NSN), as.integer(ClassID)))
      row <- row[, !names(row) %in% "EthnicityID"]
      
      LastIndex <- if (ncol(row) == 42)
        38
      else
        ncol(row)
      rrr <<- row
      row[, c(8:LastIndex)] <- as.data.frame(lapply(row[, c(8:LastIndex)], function(x) {
        ifelse(x %in% c(1, "Y", "A", "y", "a"), 1, 0)
      }))
      
      
      transformed_data <- row %>%
        pivot_longer(
          cols = all_of(8:LastIndex),
          names_to = "Description",
          values_to = "CompetencyStatusID"
        )
      
      
      
      
      # get the relevant competencies for the given time
      # match to those
      #transformed_data = transformed_data[,c("NSN", "Description", "CompetencyStatusID")]
      transformed_data <- transformed_data[!grepl("NA", transformed_data$Description), ]
      td <<- transformed_data
      
      # Merge the transformed_data and relevant competencies
      all_data = merge(transformed_data, RelevantCompetencies, by = "Description")
      if (nrow(transformed_data) != nrow(all_data)) {
        # if competencies are a mismatch
        ErrorMessage <- paste("Unable to insert:",
                              as.integer(row$NSN),
                              "as a competency is invalid")
        stop(ErrorMessage)
      }
      all_data = all_data[, c("NSN",
                              "CompetencyID",
                              "YearGroupID",
                              "CompetencyStatusID")]
      
      ## need to find all columns (CompetencyID, YearGroupID) that are in Competency and not in all_data
      all_data = merge(
        Competency,
        all_data,
        by = c("CompetencyID", "YearGroupID"),
        all.x = T
      )
      all_data = all_data[, c("NSN",
                              "CompetencyID",
                              "YearGroupID",
                              "CompetencyStatusID")]
      
      all_data$NSN = row$NSN
      all_data$CompetencyStatusID[is.na(all_data$CompetencyStatusID)] = 0
      
      all_data$Date = ifelse(all_data$CompetencyStatusID == 1,
                             GetDefaultDate(CALENDARYEAR, TERM),
                             NA)
      all_data = all_data[!(all_data$CompetencyID %in% c(32, 33)), ]
      competency_query <- "
    MERGE INTO StudentCompetency AS target
    USING (VALUES (?, ?, ?, ?, ?)) AS source (NSN, CompetencyID, YearGroupID, CompetencyStatusID, Date)
    ON target.NSN = source.NSN AND target.CompetencyID = source.CompetencyID AND target.YearGroupID = source.YearGroupID
    WHEN MATCHED AND target.Date IS NULL AND source.CompetencyStatusID = 1 THEN
        UPDATE SET
            CompetencyStatusID = source.CompetencyStatusID,
            Date = source.Date
    WHEN NOT MATCHED THEN
        INSERT (NSN, CompetencyID, YearGroupID, CompetencyStatusID, Date)
        VALUES (source.NSN, source.CompetencyID, source.YearGroupID, source.CompetencyStatusID, source.Date);
    "
      
      
      all_data$NSN[is.na(all_data$NSN)] <- as.integer(names(sort(table(all_data$NSN), decreasing = TRUE))[1])
      
      # Ensure NSN is stored as an integer
      all_data$NSN <- as.integer(all_data$NSN)
      all_data$CompetencyStatusID[is.na(all_data$CompetencyStatusID)] = 0
      all_data$NSN <- as.integer(all_data$NSN)
      
      # Ensure CompetencyID is an integer
      all_data$CompetencyID <- as.integer(all_data$CompetencyID)
      
      # Ensure CompetencyStatusID is an integer
      all_data$CompetencyStatusID <- as.integer(all_data$CompetencyStatusID)
      all_data$YearGroupID <- as.integer(all_data$YearGroupID)
      
      # Ensure Date is in the correct format (e.g., as Date or character)
      all_data$Date <- as.Date(all_data$Date)
      all_data$NSN <- as.integer(as.numeric(all_data$NSN))
      all_data$CompetencyID <- as.integer(as.numeric(all_data$CompetencyID))
      all_data$YearGroupID <- as.integer(as.numeric(all_data$YearGroupID))
      
      
      ad <<- all_data
      #print("***")
      
      for (j in 1:nrow(all_data)) {
        # print(paste("NSN:", all_data$NSN[j],
        #             "CompetencyID:", all_data$CompetencyID[j],
        #             "YearGroupID:", all_data$YearGroupID[j],
        #             "CompetencyStatusID:", all_data$CompetencyStatusID[j],
        #             "Date:", all_data$Date[j]))
        
        #Execute query
        dbExecute(
          con,
          competency_query,
          params = list(
            all_data$NSN[j],
            all_data$CompetencyID[j],
            all_data$YearGroupID[j],
            all_data$CompetencyStatusID[j],
            all_data$Date[j]
          )
        )
      }
      ProcessScenario(row, con)
      
      query <- "
    INSERT INTO StudentYearLevel(NSN, YearLevelID, CalendarYear, Term)
    SELECT ?, ?, ?, ?
    WHERE NOT EXISTS (
        SELECT 1
        FROM StudentYearLevel
        WHERE NSN = ? AND YearLevelID = ? AND CalendarYear = ? AND Term = ?
    )
    "
      dbExecute(
        con,
        query,
        params = list(
          as.integer(row$NSN),
          as.integer(row$`Year Group`),
          as.integer(CALENDARYEAR),
          as.integer(TERM),
          as.integer(row$NSN),
          as.integer(row$`Year Group`),
          as.integer(CALENDARYEAR),
          as.integer(TERM)
        )
      )
    }
    
    
    cat("Finished checking student data for ClassID:", ClassID, "\n")
  }
}
ProcessScenario <- function(row, con) {
  if (ncol(row) == 42) {
    print("Processing a row with 42 columns.")
    
    # Extract ScenarioID
    if (all(c("Scenario One - Chosen Option", "Scenario Two - Chosen Option") %in% names(row))) {
      ScenarioID <- c(row[["Scenario One - Chosen Option"]], row[["Scenario Two - Chosen Option"]])
      print("ScenarioID extracted using explicit column names.")
    } else if (all(c("7-8...39", "7-8...41") %in% names(row))) {
      ScenarioID <- c(row[["7-8...39"]], row[["7-8...41"]])
      print("ScenarioID extracted using fallback column names.")
    } else if (all(c("45511...39", "45511...41") %in% names(row))) {
      ScenarioID <- c(row[["45511...39"]], row[["45511...41"]])
      print("ScenarioID extracted using fallback column names.")
    } else{
      ScenarioID <- c(0, 0)
    }
    
    # Handle multiple or missing ScenarioID values
    if (length(ScenarioID) > 2) {
      warning("Multiple ScenarioID values found; selecting the first two.")
      ScenarioID <- ScenarioID[1:2]
    } else if (length(ScenarioID) < 2) {
      warning("Missing ScenarioID values; padding with 0.")
      ScenarioID <- c(ScenarioID, rep(0, 2 - length(ScenarioID)))
    }
    
    # Extract Status
    if (all(c("Scenario One - Completed", "Scenario Two - Completed") %in% names(row))) {
      Status <- c(row[["Scenario One - Completed"]], row[["Scenario Two - Completed"]])
      print("Status extracted using explicit column names.")
    } else if (all(c("7-8...40", "7-8...42") %in% names(row))) {
      Status <- c(row[["7-8...40"]], row[["7-8...42"]])
      print("Status extracted using fallback column names.")
    } else if (all(c("45511...40", "45511...42") %in% names(row))) {
      Status <- c(row[["45511...40"]], row[["45511...42"]])
      print("Status extracted using fallback column names.")
    } else{
      Status <- c(0, 0)
    }
    
    # Process Status and ScenarioID
    Status[is.na(Status)] <- 0
    Status <- ifelse(Status %in% c(1, "Y", "A", "y", "a"), 1, 0)
    print(paste("Processed Status:", paste(Status, collapse = ", ")))
    
    ScenarioID[is.na(ScenarioID)] <- 0
    ScenarioID <- ifelse(ScenarioID %in% 0:4, ScenarioID, 0)
    print(paste("Processed ScenarioID:", paste(ScenarioID, collapse = ", ")))
    
    
    CountScenario <- as.numeric(dbGetQuery(
      con,
      paste("SELECT count(*) FROM StudentScenario WHERE NSN =", row$NSN)
    ))
    print(CountScenario)
    if (CountScenario == 0) {
      print(paste(
        "NSN",
        row$NSN,
        "not found in StudentScenario. Adding both scenarios."
      ))
      
      # Add both scenarios to StudentScenario
      dbWriteTable(
        con,
        "StudentScenario",
        data.frame(
          NSN = row$NSN,
          ScenarioIndex = c(1, 2),
          ScenarioID = ScenarioID
        ),
        append = TRUE,
        row.names = FALSE
      )
    }
    CountComp <- as.numeric(dbGetQuery(
      con,
      paste(
        "SELECT count(*) FROM StudentCompetency WHERE CompetencyID = 33 and NSN =",
        row$NSN
      )
    ))
    print(CountComp)
    if (CountComp == 0) {
      # Add both competencies to StudentCompetency
      dbWriteTable(
        con,
        "StudentCompetency",
        data.frame(
          NSN = row$NSN,
          CompetencyID = c(32, 33),
          YearGroupID = 4,
          CompetencyStatusID = Status,
          Date = NA
        ),
        append = TRUE,
        row.names = FALSE
      )
      
      print(paste("Added default scenarios and competencies for NSN", row$NSN))
      
      
    }
    # Check completed competencies
    CountComplete <- as.numeric(dbGetQuery(
      con,
      paste(
        "SELECT count(*) FROM StudentCompetency WHERE NSN =",
        row$NSN,
        "AND CompetencyStatusID=1 AND CompetencyID IN (32,33)"
      )
    ))
    
    print(paste("Count of completed competencies:", CountComplete))
    
    if (CountComplete == 2) {
      print("Both scenarios are complete. Nothing needs changing.")
    } else if (CountComplete == 0) {
      print("No scenarios are complete. Updating both.")
      CompetencyID <- c(32, 33)
      defaultDate <- GetDefaultDate(CALENDARYEAR, TERM)
      
      for (i in 1:2) {
        CompetencyUpdate <- paste(
          "UPDATE StudentCompetency SET CompetencyStatusID =",
          Status[i],
          ", Date = '",
          defaultDate,
          "'",
          "WHERE NSN =",
          row$NSN,
          "AND CompetencyID =",
          CompetencyID[i],
          "AND YearGroupID = 4"
        )
        dbExecute(con, CompetencyUpdate)
        print(paste(
          "Updated CompetencyID",
          CompetencyID[i],
          "for NSN",
          row$NSN
        ))
        
        ScenarioUpdate <- paste(
          "UPDATE StudentScenario SET ScenarioID =",
          ScenarioID[i],
          "WHERE NSN =",
          row$NSN,
          "AND ScenarioIndex =",
          i
        )
        dbExecute(con, ScenarioUpdate)
        print(paste("Updated ScenarioIndex", i, "for NSN", row$NSN))
      }
    } else if (CountComplete == 1) {
      print("One scenario is complete. Determining updates.")
      CompleteID <- dbGetQuery(
        con,
        paste(
          "SELECT CompetencyID FROM StudentCompetency
         WHERE NSN =",
          row$NSN,
          "AND CompetencyID IN (32, 33) AND CompetencyStatusID = 1"
        )
      )$CompetencyID
      
      if (!is.null(CompleteID) && CompleteID %in% ScenarioID) {
        print("The completed ScenarioID is already present.")
        ScenarioID <- setdiff(ScenarioID, CompleteID)
      }
      
      # Update logic for incomplete scenario
      if (length(Status[Status == 1]) == 1) {
        dataframe <- data.frame(ID = ScenarioID, St = Status)
        dataframe <- dataframe[dataframe$St == 1, ]
        CompetencyUpdate <- paste(
          "UPDATE StudentCompetency SET CompetencyStatusID =",
          dataframe$St,
          ", Date = '",
          GetDefaultDate(CALENDARYEAR, TERM),
          "'",
          "WHERE NSN =",
          row$NSN,
          "AND CompetencyID = 33 AND YearGroupID = 4"
        )
        dbExecute(con, CompetencyUpdate)
        ScenarioUpdate <- paste(
          "UPDATE StudentScenario SET ScenarioID =",
          dataframe$ID,
          "WHERE NSN =",
          row$NSN,
          "AND ScenarioIndex = 2"
        )
        dbExecute(con, ScenarioUpdate)
      }
    } else {
      stop("NSN",
           row$NSN,
           "has more than 2 scenarios marked as complete.")
    }
  } else{
    CountScenario <- as.numeric(dbGetQuery(
      con,
      paste("SELECT count(*) FROM StudentScenario WHERE NSN =", row$NSN)
    ))
    print(CountScenario)
    if (CountScenario == 0) {
      print(paste(
        "NSN",
        row$NSN,
        "not found in StudentScenario. Adding both scenarios."
      ))
      
      # Add both scenarios to StudentScenario
      dbWriteTable(
        con,
        "StudentScenario",
        data.frame(
          NSN = row$NSN,
          ScenarioIndex = c(1, 2),
          ScenarioID = c(0, 0)
        ),
        append = TRUE,
        row.names = FALSE
      )
    }
    CountComp <- as.numeric(dbGetQuery(
      con,
      paste(
        "SELECT count(*) FROM StudentCompetency WHERE CompetencyID = 33 and NSN =",
        row$NSN
      )
    ))
    print(CountComp)
    if (CountComp == 0) {
      # Add both competencies to StudentCompetency
      dbWriteTable(
        con,
        "StudentCompetency",
        data.frame(
          NSN = row$NSN,
          CompetencyID = c(32, 33),
          YearGroupID = 4,
          CompetencyStatusID = 0,
          Date = NA
        ),
        append = TRUE,
        row.names = FALSE
      )
      
      print(paste("Added default scenarios and competencies for NSN", row$NSN))
    }
  }
}

TidyEthnicityValue <- function(ethnicity) {
  # Ensure the input is trimmed
  ethnicity <- trimws(ethnicity)
  
  # Maori
  if (grepl("Māori|Ma|maori|NZ Maori|MNZ", ethnicity, ignore.case = TRUE)) {
    return("Maori")
  }
  
  # Pacific Peoples
  if (grepl(
    "Pacific People|pacific People|Pacific Island|Pacfic People|Pacifc people|Tongan|Samoan|Tokelauan|Cook Island Maori|Cook Isl Maori|Fijian|pacific|cook island maori",
    ethnicity,
    ignore.case = TRUE
  )) {
    return("Pacific Peoples")
  }
  
  # Unknown
  if (is.na(ethnicity) ||
      grepl("Not Stated", ethnicity, ignore.case = TRUE)) {
    return("Unknown")
  }
  
  # NZ European
  if (grepl(
    "NZ Europena|Z Europea|NZ Eroupean|Pākehā|New Zealand European/Pākehā",
    ethnicity,
    ignore.case = TRUE
  )) {
    return("NZ European")
  }
  
  # Asian
  if (grepl(
    "asian|asisan|indian|vietnamese|chinese|Sri Lankan|Pakistani|Filipino|Cambodian|Indonesian|Korean",
    ethnicity,
    ignore.case = TRUE
  )) {
    return("Asian")
  }
  
  # Other
  if (grepl(
    "Other European|Middle East|Other Groups|Others|Othrt|Ethiopian|Irish|Dutch|Latin American|other|african|Australian",
    ethnicity,
    ignore.case = TRUE
  )) {
    return("Other")
  }
  
  # If none of the above, return the input value unchanged
  return(ethnicity)
}

# Main script
CALENDARYEAR = 2024
TERM = 4
ProviderIndex = 2
SchoolIndexRange = c(6)
con <- GetWSNZAzureConnection()
Provider <- dbGetQuery(con, "select * from Provider")
Ethnicity <- dbGetQuery(con,
                        "select EthnicityID, Description as Ethnicity from Ethnicity")
MOE_SchoolDirectory <- dbGetQuery(
  con,
  "SELECT MOENumber, SchoolName, EducationRegionID, SchoolTypeID FROM MOE_SchoolDirectory"
)
Class <- dbGetQuery(con, "select * from Class")
ProviderEducationRegion <- dbGetQuery(con, "select * from ProviderEducationRegion")
SchoolTypeYearLevel <- dbGetQuery(con, "select * from SchoolTypeYearLevel")
Student <- dbGetQuery(con, "select * from Student")
Competency <- dbGetQuery(con, "select * from Competency")
Error <- dbGetQuery(con, "SELECT * FROM error")



path <- GetFolderPath(CALENDARYEAR, TERM)
cat("Path now set to:\n", path, "\n")

FoldersProviders <- list.dirs(path, recursive = FALSE)
cat("Found", length(FoldersProviders), "Providers\n")


# Subset the folders to process
FoldersBatch <- FoldersProviders[ProviderIndex]


# Process the subset of providers
invisible(sapply(FoldersBatch, function(p)
  ProcessProvider(p, Provider, con, SchoolIndexRange)))



