library(DBI)
library(odbc)
library(openxlsx)
library(readxl)
library(stringdist)
names(data)
# Step 1: DB connection
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

con <- GetWSNZAzureConnection()

all_moe_query <- "
select sd.MOENumber, sd.Schoolname, ta.Description as TLA,
       sd.towncity as Town, ei.EquityIndex, p.Description as Provider,
       bt.Description as PoolType
from MOE_SchoolDirectory sd 
left join MOE_EquityIndex ei on ei.MOENumber = sd.MOENumber and ei.Year = 2024 
left join SchoolProvider sp on sp.MOENumber = sd.MOENumber 
     and ((sp.CalendarYear = 2024 and sp.Term IN (3,4)) or (sp.CalendarYear = 2025 and sp.Term IN (1)))
left join provider p on sp.ProviderID = p.ProviderID
left join MOE_TerritorialAuthority ta on sd.TerritorialAuthorityID = ta.TerritorialAuthorityID
left join MOE_SchoolPool po on po.MOENumber = sd.MOENumber
left join MOE_BuildingType bt on bt.BuildingTypeCode = po.BuildingTypeCode
"
sd = dbGetQuery(con, "Select * from MOE_SchoolDirectory")
all_data <- dbGetQuery(con, all_moe_query)
final_df = dbGetQuery(con, "Select * from HealthyActiveLearningSchool")
final_df = merge(final_df, all_data, by ='MOENumber')
final_df=unique(final_df)
wb <- createWorkbook()
hal_list <- unique(final_df$RegionalSportTrust)

for (hal in hal_list[order(hal_list)]) {
  df_subset <- final_df[final_df$RegionalSportTrust == hal, ]
  if (all(is.na(df_subset$Label))) {
    df_subset$Label <- NULL
  }
  sheet_name <- substr(hal, 1, 31)
  addWorksheet(wb, sheet_name)
  writeDataTable(wb, sheet_name, df_subset)
}

addWorksheet(wb, "All HALs")
writeDataTable(wb, "All HALs", final_df)

addWorksheet(wb, "School Data")
writeDataTable(wb, "School Data", sd[,c(1:16)])


saveWorkbook(wb, paste0("HALS Data (",gsub("-", ".", as.character(Sys.Date())),").xlsx"), overwrite = TRUE)
