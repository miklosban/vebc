#' @title 
#' Generate metadata Excel templates for OBM upload
#'
#' @description
#' Connects to the `VEBC (sex_ratio_evolution)` OpenBioMaps project using the obm.R package, 
#' retrieves data from a specified table, and generates two Excel 
#' templates: `table_metadata.xlsx` and `variable_metadata.xlsx`.
#' This files needed to document table- and variable-level metadata.
#'
#' @param table_name Character. The name of the table to query from OpenBioMaps.
#' @param schema_name Character. Target schema. This is optional, default is sex_ratio_evolution
#' @param url Character. OBM server URL (default: "https://openbiomaps.org").
#' @param project Character. OBM project name (default: "sex_ratio_evolution").
#' @param output_dir Character. Output directory for the generated Excel files.
#' If `NULL` (default), an interactive directory selection dialog is opened.
#'
#' @return Creates two Excel files in the working directory:
#' * `<table_name>_table_metadata.xlsx`
#' * `<table_name>_variable_metadata.xlsx`
#'
#' Invisibly returns a list containing the generated metadata tables and
#' the full paths of the created files.
#'
#' @examples
#' \dontrun{
#' generate_metadata_files("my_table", "optional: schema_name", "optional: url", "optional: project", "optional: output_dir")
#' }
#'
#' @import obm
#' @import openxlsx
#' @export
generate_metadata_files <- function(
    table_name,
    schema_name = "sex_ratio_evolution",
    url = "https://openbiomaps.org",
    project = "sex_ratio_evolution",
    output_dir = getwd()
) {
  # Required packages
  #if (!requireNamespace("obm", quietly = TRUE)) stop("Package 'obm' must be installed.")

  if (is.na(output_dir)) {
    stop("No output directory selected.")
  }

  # Ask for user information
  username <- readline("Enter your name: ")
  useremail <- readline("Enter your login email address: ")
  # Current date and time
  date_uploading <- format(Sys.time(), "%Y-%m-%dT%H:%M:%S%z") 


  message("Output directory: ", output_dir)
  message("User: ", username, " <", useremail, ">")
  message("Date of uploading: ", date_uploading)


  # Create output file paths
  table_metadata_file <- file.path(
    output_dir,
    paste0(table_name, "_table_metadata.xlsx")
  )

  variable_metadata_file <- file.path(
    output_dir,
    paste0(table_name, "_variable_metadata.xlsx")
  )
  


  # Initialize OBM connection
  obm::obm_init(project = project, url = url, api_version = 2.3)
  obm::obm_auth(username = useremail)

  header_style <- createStyle(
    textDecoration = c("bold"), 
    fgFill = "#DCE6F1"
  )

  excluded_fields <- c(
    "obm_id",
    "obm_files_id",
    "obm_uploading_id",
    "obm_modifier_id",
    "obm_validation",
    "obm_comments",
    "obm_geometry"
  )
  
  # Retrieve data
  message(paste("Querying table:", table_name))

  full_table_name <- paste0(schema_name, ".", table_name)
  tbl <- obm::obm_get("get_data", "*", table = full_table_name)
  data <- tbl[, !(colnames(tbl) %in% excluded_fields), drop = FALSE]
  
  # --- table_metadata.xlsx ---
  table_headers <- c(
    "table_owner",
    "date_uploading",
    "table_name",
    "focus_group",
    "data_type",
    "data_type_var",
    "species_var",
    "population_var",
    "date_end_datacollection",
    "comment"
  )
  
  table_metadata <- data.frame(matrix(ncol = length(table_headers), nrow = 1))
  colnames(table_metadata) <- table_headers
  table_metadata[1, "table_name"] <- table_name
  table_metadata[1, "table_owner"] <- paste(username, useremail) 
  table_metadata[1, "date_uploading"] <- date_uploading

  wb <- createWorkbook()
  addWorksheet(wb, "metadata")
  writeData(wb, "metadata", table_metadata, colNames = TRUE)
  
  # Focus group dropdown
  focus_options <- c("mammals","birds","reptiles","amphibians","fish",
                   "mixed_tetrapoda","mixed_amniote","mixed_vertebrate",
                   "mixed_other")
  dataValidation(wb, "metadata", cols = which(colnames(table_metadata) == "focus_group"),
               rows = 2, type = "list",
               value = paste(focus_options, collapse = ","))

  # Data type dropdown
  data_type_options <- c("species", "population", "mixed")
  dataValidation(wb, "metadata", cols = which(colnames(table_metadata) == "data_type"),
               rows = 2, type = "list",
               value = paste(data_type_options, collapse = ","))

  # Species_var dropdown: table fieldnames
  colnames_str <- paste(colnames(data), collapse = ",")
  dataValidation(wb, "metadata", cols = which(colnames(table_metadata) == "species_var"),
               rows = 2, type = "list", value = colnames_str)

  # Population_var dropdown: table field names
  dataValidation(wb, "metadata", cols = which(colnames(table_metadata) == "population_var"),
               rows = 2, type = "list", value = colnames_str)

  addStyle(wb, sheet = "metadata", style = header_style,
         rows = 1, cols = 1:ncol(table_metadata), gridExpand = TRUE)

  setColWidths(wb, "metadata", cols = 1:ncol(table_metadata), widths = "auto")


  # --- variable description tab ---
  addWorksheet(wb, "variable_descriptions")
  description_data <- data.frame(
    variable = c("table_owner","date_uploading","table_name","focus_group","data_type","data_type_var","species_var","population_var","date_end_datacollection","comment"),
    description = c(
      "Name/contact of the data owner. Automatically filled.",
      "Data of uploading the data to the server. Automatically filled.",
      "Name of the data file o the server. Automatically filled.",
      "The smallest possible taxonomic group which contains all the species from you data.",
      "Is the data about species or is it on the population-level? If it contains data on both levels, then please choose 'mixed'.",
      "If you data_type is mixed, then please indicate the variable from your dataset which cathegorises your observations into species or population-level data. If you do not have such a variable, choose 'None'.",
      "Indicate the name of the variable in your dataset which contains the species names (scientific names).",
      "Indicate the name of the variable in your dataset which identifies your populations.",
      "Indicate in which year did the data collection end.",
      "Any information you think might be useful for other users."
      ),
        stringsAsFactors = FALSE
  )
  
  writeData(wb, "variable_descriptions", description_data, colNames = TRUE)
  addStyle(wb, sheet = "variable_descriptions", style = header_style, rows = 1, cols = 1:ncol(description_data), gridExpand = TRUE)
  setColWidths(wb, "variable_descriptions", cols = 1:ncol(description_data), widths = "auto")

  # --- categories explanation tab ---
  addWorksheet(wb, "categories")
  
  category_data <- data.frame(
    variable = c(rep("focus_group",9),rep("data_type",3)),
    categories = c(focus_options, data_type_options),
    description = c(
      "The file only contains data on mammals.",
      "The file only contains data on birds.",
      "The file only contains data on reptiles.",
      "The file only contains data on amphibians.",
      "The file only contains data on fish.",
      "The file contains data on tetrapods.",
      "The file contains data on amniotes.",
      "The file contains data on vertebrates.",
      "The file contains data on several different taxonomic groups which couldn't be fit into the categories provided.",
      "The file only contains data on species.",
      "The file only contains data on population.",
      "The file only contains data on mixed."),
        stringsAsFactors = FALSE
  )
  
  writeData(wb, "categories", category_data, colNames = TRUE)
  addStyle(wb, sheet = "categories", style = header_style, rows = 1, cols = 1:ncol(category_data), gridExpand = TRUE)
  setColWidths(wb, "categories", cols = 1:ncol(category_data), widths = "auto")

  saveWorkbook(
        wb,
        table_metadata_file,
        overwrite = TRUE
  )
  message(
    "Created: ",
    table_metadata_file
  )

  # --- variable_metadata.xlsx ---
  var_headers <- c("variable_name", "var_category", "var_unit", "var_type", "var_description")
  variable_metadata <- data.frame(matrix(ncol = length(var_headers), 
                                         nrow = length(colnames(data))))
  colnames(variable_metadata) <- var_headers
  variable_metadata$variable_name <- colnames(data)

  wb <- createWorkbook()
  addWorksheet(wb, "variable_metadata")

  writeData(wb, "variable_metadata", variable_metadata, colNames = TRUE)
  
  var_category_options <- c("Behaviour","Climate","Geography","Demography",
                          "Ecology","LifeHistory","Morphology","Taxonomy",
                          "Metadata","Method")

  var_type_options <- c("character","numeric","factor","date","logical")

  n_vars <- nrow(variable_metadata)

  dataValidation(wb,"variable_metadata",cols = which(colnames(variable_metadata) == "var_category"),
    rows = 2:(n_vars + 1),type = "list",
    value = paste(var_category_options, collapse = ",")
  )
  dataValidation(wb,"variable_metadata",cols = which(colnames(variable_metadata) == "var_type"),
    rows = 2:(n_vars + 1),type = "list",
    value = paste(var_type_options, collapse = ",")
  )

  addStyle(wb, sheet = "variable_metadata", style = header_style,
         rows = 1, cols = 1:ncol(variable_metadata), gridExpand = TRUE)

  setColWidths(wb, "variable_metadata", cols = 1:ncol(variable_metadata), widths = "auto")

  # --- variable description tab ---
  addWorksheet(wb, "variable_descriptions")
  description_data <- data.frame(
    variable = c("variable_name","var_category","var_unit","var_type","var_description"),
    description = c(
      "Name of the given variable.",
      "Category of the given variable. See the description of the different categories on the 'categories' sheet.",
      "Unit of the given variable, if applicable. E.g. for morphological measurements or time related variables (months or years).",
      "Type of the variable in terms of R programming language. See the different types on the 'categories' sheet.",
      "Description of the given variable. Provide a short but clear description of the variable in question."
      ),
        stringsAsFactors = FALSE
  )
  
  writeData(wb, "variable_descriptions", description_data, colNames = TRUE)
  addStyle(wb, sheet = "variable_descriptions", style = header_style, rows = 1, cols = 1:ncol(description_data), gridExpand = TRUE)
  setColWidths(wb, "variable_descriptions", cols = 1:ncol(description_data), widths = "auto")

  # --- categories explanation tab ---
  addWorksheet(wb, "categories")
  
  category_data <- data.frame(
    variable = c(
      rep("var_category", length(var_category_options)),
      rep("var_type", length(var_type_options))
    ),
    categories = c(var_category_options, var_type_options),
    description = c(
      rep("", length(var_category_options) - 3),
      "Any variable that contains information on taxonomy, e.g. species names, scientific names, order, family",
      "Any variable that helps identify the observations in your data, e.g. IDs, references or any additional information regarding your observations.",
      "Variables that are related to the methods with which an observation was collected, e.g. method of data collection, data quality, sample size, year of data collection, duration of data collection.",
      rep("", length(var_type_options))
    ),
    stringsAsFactors = FALSE
  )
  
  writeData(wb, "categories", category_data, colNames = TRUE)
  addStyle(wb, sheet = "categories", style = header_style, rows = 1, cols = 1:ncol(category_data), gridExpand = TRUE)
  setColWidths(wb, "categories", cols = 1:ncol(category_data), widths = "auto")

  saveWorkbook(
    wb,
    variable_metadata_file,
    overwrite = TRUE
  )

  message(
    "Created: ",
    variable_metadata_file
  )
  
  invisible(list(
    table_metadata = table_metadata,
    variable_metadata = variable_metadata,
    files = list(
      table_metadata = table_metadata_file,
      variable_metadata = variable_metadata_file
    )
  ))
}
