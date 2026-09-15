#### Process Structure results based on Clumpak output #########################

## Summary of code
# 1) Creates Excel file with assignment probabilities for each k-value
# 2) Assign individuals to clusters for each K based on defined assignment probability thresholds
# 3) Creates combined dataframe and csv file containing all Clumpak results with assignment probabilities and cluster IDs across different K-values usable for subsequent analyses or plotting
# 4) Create maps with pie charts showing cluster assignment probabilities for each k-value
# 5) Create Structure-like stacked barplots showing cluster assignment probabilities for each k-value

## Set environment
rm(list = ls()) #clear environment
setwd("C:/Users/danie/Desktop/PhD research/Hemileuca maia research/Analyses/Structure")


## Define input
input_assignment_file <- "Clumpak/Hmg_pop_45_Clumpak_assignment.txt" #define input file containing individual ID, species/subspecies, latitude, and longitude in same order as Structure/CLUMPAK input
clumpak_folder <- "Clumpak_results/Hmg_pop_45" #define folder containing K=1, K=2, ... Clumpak result folders
structure_vcf_file <- "../Stacks/Hmg_pop_45.vcf" #define original vcf file used to create Structure input
outgroup <- c("CA_OroGrandeWash_Hmg086", "CA_OroGrandeWash_Hmg087") #define outgroup individuals excluded from Structure analyses

ID <- "ID_sequence" #define ID column name in input_assignment_file
Species <- "Species_subspecies" #define Species column name in input_assignment_file
kmin <- 2 #define minimum k-value used for Structure analyses
kmax <- 15 #define maximal k-value used for Structure analyses
threshold_high <- 0.75 #define threshold: individuals with assignment probabilities above this threshold will be assigned to respective single cluster (e.g., cluster "1")
max_shared_clusters <- 3 #define maximum number of clusters used for shared assignments (e.g., cluster "1 + 2 + 3")
admixed_individuals <- "High admixture" #set name for individuals requiring more than three clusters to reach threshold_high
mapping_regions <- c("Canada", "USA") #set mapping regions (should contain entire range of samples)
states <- "yes" #specify "yes" or "no" to include or exclude US state boundaries (only if USA is included in map)


## Define output parameters
results_dir <- "Clumpak_results_processed/Hmg_pop_45" #define output folder where processed Excel files, combined CSV file, and plot SVG files will be saved
Excel_individual_file_name <- "Clumpak_results" #define Excel output file name for individual datasets for each k value (without ".xlsx")
CSV_combined_file_name <- "Clumpak_results_combined" #define CSV output file name for combined data (without ".csv")
Excel_output_decimal <- "." #specify decimals for Excel/CSV output (e.g., "." or ",")


## Check and install required packages
install_if_missing <- function(pkg) { #function to check and install required packages
  if (!requireNamespace(pkg, quietly = TRUE)) utils::install.packages(pkg, dependencies = T)
}
install_if_missing("viridis") #check and install required packages
install_if_missing("openxlsx") #check and install required packages
install_if_missing("maps") #check and install required packages


## Set mapping parameters
jitter_setting <- 0.2 #set degree of jitter (to avoid overlapping pie charts)
jitter_seed <- 1 #set seed for reproducible jitter so individual locations are identical across k-values
pie_size <- 1.3 #set size of pie charts
pie_border_color <- "black" #modify pie chart border color
pie_border_thickness <- 0.8 #modify pie chart border thickness
viridis_palette <- viridis::magma #set colorblind-friendly viridis color palette (alternatively use: "viridis::viridis" or "viridis::mako" or "viridis::inferno")
buffer_percentage <- 0.7 #modify buffer around range of coordinates for map
state_border_color <- "gray30" #modify state border color for map
state_border_thickness <- 0.3 #modify state border thickness for map
country_border_color <- "gray30" #modify country border color for map
country_border_thickness <- 0.75 #modify country border thickness for map
width_plot <- 16/2.54 #modify width of plot for map
height_plot <- 20/2.54 #modify height of plot for map
map_filling_color <- "lightgrey" #modify map filling color
legend_position <- "topright" #modify position of cluster legend


## Set Structure-like plot parameters
structure_plot_width <- 35/2.54 #modify width of Structure-like plot
structure_plot_height <- 14/2.54 #modify height of Structure-like plot
structure_sort_by_cluster <- 1 #define first dominant cluster used for sorting individuals (NULL = keep original Structure order)
structure_show_individual_labels <- TRUE #specify whether individual IDs are shown below Structure-like plot
structure_individual_label_size <- 0.35 #modify size of individual ID labels
structure_bottom_margin <- 15 #modify bottom margin to provide space for individual ID labels
structure_left_margin <- 5 #modify left margin
structure_top_margin <- 2 #modify top margin
structure_right_margin <- 2 #modify right margin
structure_y_axis_title <- "Cluster assignment coefficient" #define y-axis title

structure_individual_lines <- "yes" #specify "yes" or "no" to draw vertical separation lines between individuals
structure_individual_line_color <- "grey" #modify color of vertical separation lines between individuals
structure_individual_line_thickness <- 0.25 #modify thickness of vertical separation lines between individuals

structure_bar_lines <- "yes" #specify "yes" or "no" to draw borders around individual colored bar segments
structure_bar_line_color <- "black" #modify color of borders around colored bar segments
structure_bar_line_thickness <- 0.2 #modify thickness of borders around colored bar segments

structure_label_file <- "../Data/2025_05_07_Hmg_data.csv" #optional file containing replacement individual labels (NULL = use original ID_sequence labels)
structure_label_match_column <- "ID_sequence" #column containing original individual IDs used for matching
structure_label_replacement_column <- "ID" #column containing replacement individual labels
structure_label_file_sep <- ";" #separator used in structure_label_file
structure_label_file_dec <- "." #decimal separator used in structure_label_file


## No further input required below

## Process Structure data using Clumpak output and plot cluster assignments
if (!file.exists(input_assignment_file)) stop("Input assignment file does not exist: ", input_assignment_file) #stop if assignment file does not exist
if (!dir.exists(clumpak_folder)) stop("Clumpak results folder does not exist: ", clumpak_folder) #stop if Clumpak folder does not exist
if (!file.exists(structure_vcf_file)) stop("Structure vcf file does not exist: ", structure_vcf_file) #stop if original vcf file does not exist
if (threshold_high <= 0.5 || threshold_high > 1) stop("threshold_high must be > 0.5 and <= 1") #check threshold_high
if (max_shared_clusters < 2) stop("max_shared_clusters must be at least 2") #check maximum number of shared clusters
if (!is.null(structure_sort_by_cluster) && (!is.numeric(structure_sort_by_cluster) || length(structure_sort_by_cluster) != 1 || is.na(structure_sort_by_cluster) || structure_sort_by_cluster < 1)) stop("structure_sort_by_cluster must be NULL or a positive integer") #check Structure-like sorting parameter
if (!(tolower(structure_individual_lines) %in% c("yes", "no"))) stop("structure_individual_lines must be 'yes' or 'no'") #check individual line setting
if (!(tolower(structure_bar_lines) %in% c("yes", "no"))) stop("structure_bar_lines must be 'yes' or 'no'") #check bar line setting
if (!dir.exists(results_dir)) dir.create(results_dir, recursive = TRUE) #create output directory if needed
dataset <- utils::read.table(file = input_assignment_file, sep = "\t", header = TRUE, stringsAsFactors = FALSE) #import dataset containing Individual ID, Species ID, Latitude and Longitude in the same order as Structure input
head(dataset) #check dataset
nrow(dataset)
required_columns <- c(ID, Species, "Latitude", "Longitude") #define required columns
missing_columns <- setdiff(required_columns, colnames(dataset)) #check for missing columns
if (length(missing_columns) > 0) stop("Missing required columns in input_assignment_file: ", paste(missing_columns, collapse = ", "))
dataset_ids <- as.character(dataset[[ID]]) #extract individual IDs from assignment file
if (any(is.na(dataset_ids) | dataset_ids == "")) stop("Missing individual IDs in input_assignment_file") #stop if individual IDs are missing
if (anyDuplicated(dataset_ids)) stop("Duplicated individual IDs in input_assignment_file") #stop if individual IDs are duplicated
vcf_lines <- readLines(structure_vcf_file, warn = FALSE) #read original vcf file used to create Structure input
vcf_header <- grep("^#CHROM", vcf_lines, value = TRUE) #extract vcf header containing sample IDs
if (length(vcf_header) != 1) stop("Could not uniquely identify #CHROM header in Structure vcf file") #stop if vcf header cannot be identified
vcf_header <- strsplit(vcf_header, "\t", fixed = TRUE)[[1]] #split vcf header into columns
if (length(vcf_header) <= 9) stop("No individual IDs found in Structure vcf file") #stop if vcf contains no individuals
vcf_ids <- vcf_header[-(1:9)] #extract individual IDs from vcf header
vcf_ids <- vcf_ids[!vcf_ids %in% outgroup] #remove outgroup individuals excluded from Structure analyses
if (anyDuplicated(vcf_ids)) stop("Duplicated individual IDs in Structure vcf file") #stop if vcf IDs are duplicated
if (length(vcf_ids) != nrow(dataset)) stop("Number of individuals in input_assignment_file does not match Structure vcf file") #check sample number
if (!setequal(vcf_ids, dataset_ids)) { #check that assignment file and vcf contain exactly the same individuals
  cat("In assignment file but not Structure vcf:\n")
  print(setdiff(dataset_ids, vcf_ids))
  cat("In Structure vcf but not assignment file:\n")
  print(setdiff(vcf_ids, dataset_ids))
  stop("Individuals in input_assignment_file do not match Structure vcf file")
}
if (!identical(vcf_ids, dataset_ids)) { #check that individual order exactly matches Structure vcf order
  mismatch <- which(vcf_ids != dataset_ids)
  print(head(data.frame(Row = mismatch, Structure_vcf_ID = vcf_ids[mismatch], Assignment_file_ID = dataset_ids[mismatch]), 20))
  stop("Individual order in input_assignment_file does not match Structure vcf order")
}
format_decimal <- function(x, dec = Excel_output_decimal) { #define function for formatting decimals
  formatted <- format(x, decimal.mark = dec, scientific = FALSE, trim = TRUE) #format numbers with specified decimal
  return(as.character(formatted)) #convert to character to ensure correct formatting
}
path_clumpak_files <- file.path(getwd(), clumpak_folder) #set path of Clumpak output files and dataset
clumpak_results_list <- list() #initialize list
for (k in kmin:kmax) { #loop through K values
  file_path <- file.path(path_clumpak_files, paste0("K=", k), "CLUMPP.files", "ClumppIndFile.output")
  if (!file.exists(file_path)) stop("File not found for K = ", k, ": ", file_path) #stop if expected Clumpak file does not exist
  clumpak_data <- tryCatch({ #read data from file
    utils::read.table(file_path, header = FALSE, stringsAsFactors = FALSE)
  }, error = function(e) {
    stop(paste("Error reading file for K=", k, ": ", e$message, sep = ""))
  })
  if (is.null(clumpak_data) || ncol(clumpak_data) == 0) stop("K=", k, ": Clumpak file contains no data") #check if data is NULL or empty
  if (ncol(clumpak_data) <= 5) stop("K=", k, ": CLUMPAK file has <= 5 columns. No assignment probability columns found") #check that CLUMPAK file contains assignment probability columns
  if (nrow(clumpak_data) != nrow(dataset)) stop("K=", k, ": CLUMPAK rows do not match dataset rows. Check input order and sample number") #check that CLUMPAK output and dataset have same number of rows
  clumpak_ids <- as.character(clumpak_data[[2]]) #extract sequential individual numbers from CLUMPAK output
  expected_clumpak_ids <- as.character(seq_len(nrow(dataset))) #define expected sequential individual numbers
  if (!identical(clumpak_ids, expected_clumpak_ids)) stop("K=", k, ": CLUMPAK individual numbers/order do not match expected sequential order") #stop if CLUMPAK individual order differs
  clumpak_data <- clumpak_data[, -(1:5), drop = F] #delete first five columns
  if (ncol(clumpak_data) != k) stop("K=", k, ": Number of CLUMPAK assignment columns does not match k") #check that number of assignment columns matches k
  clumpak_data[] <- lapply(clumpak_data, function(x) suppressWarnings(as.numeric(x))) #ensure assignment probabilities are numeric
  if (anyNA(clumpak_data)) stop("K=", k, ": CLUMPAK assignment probabilities contain missing or non-numeric values") #stop if assignment probabilities contain missing values
  if (any(as.matrix(clumpak_data) < 0 | as.matrix(clumpak_data) > 1)) stop("K=", k, ": CLUMPAK assignment probabilities contain values outside 0-1") #stop if assignment probabilities are outside valid range
  probability_sums <- rowSums(clumpak_data) #calculate sum of assignment probabilities for each individual
  if (any(abs(probability_sums - 1) > 0.01)) stop("K=", k, ": CLUMPAK assignment probabilities do not sum approximately to 1") #stop if assignment probabilities do not sum approximately to one
  colnames(clumpak_data) <- paste0(1:ncol(clumpak_data), "_k", k) #rename remaining columns based on k-value
  clumpak_data_excel <- clumpak_data #copy data frame for individual Excel output
  clumpak_data_excel$ID <- dataset[[ID]] #add ID of individual
  clumpak_data_excel$Species <- dataset[[Species]] #add Species
  clumpak_results_list[[length(clumpak_results_list) + 1]] <- clumpak_data #store only assignment probability columns in list
  wb <- openxlsx::createWorkbook() #create new Excel workbook
  openxlsx::addWorksheet(wb, sheetName = paste0("K", k))
  openxlsx::writeData(wb, sheet = 1, x = clumpak_data_excel, colNames = T, borders = "columns")
  numeric_columns_index <- which(sapply(clumpak_data_excel, is.numeric)) #identify numeric columns
  number_format <- openxlsx::createStyle(numFmt = ifelse(Excel_output_decimal == ",", "0,00", "0.00")) #format with 2 decimal places
  if (length(numeric_columns_index) > 0) openxlsx::addStyle(wb, sheet = 1, style = number_format, rows = 2:(nrow(clumpak_data_excel) + 1), cols = numeric_columns_index, gridExpand = T)
  excel_file <- file.path(results_dir, paste0(Excel_individual_file_name, "_K", k, ".xlsx")) #define Excel file name
  openxlsx::saveWorkbook(wb, file = excel_file, overwrite = T)
}
if (length(clumpak_results_list) != length(kmin:kmax)) stop("Not all requested K values were successfully imported") #check that all requested K values were imported
Clumpak_results <- do.call(cbind, clumpak_results_list) #combine all K data frames into one data frame
Clumpak_results <- Clumpak_results[ , !(names(Clumpak_results) %in% c("ID", "Species")), drop = F] #remove duplicated ID and Species columns
dataset <- cbind(dataset, Clumpak_results) #concatenate with dataset


## Create individual labels for Structure-like plots
structure_sample_labels <- dataset_ids #use original individual IDs by default
if (!is.null(structure_label_file)) { #replace individual labels if optional label file is specified
  if (!file.exists(structure_label_file)) stop("Structure label file does not exist: ", structure_label_file) #stop if specified label file does not exist
  structure_label_data <- utils::read.csv(file = structure_label_file,
                                          sep = structure_label_file_sep,
                                          dec = structure_label_file_dec,
                                          header = TRUE,
                                          stringsAsFactors = FALSE) #import label replacement file
  required_structure_label_columns <- c(structure_label_match_column, structure_label_replacement_column) #define required columns
  missing_structure_label_columns <- setdiff(required_structure_label_columns, colnames(structure_label_data)) #check required columns
  if (length(missing_structure_label_columns) > 0) stop("Missing required columns in structure_label_file: ", paste(missing_structure_label_columns, collapse = ", ")) #stop if required columns are missing
  structure_label_match_ids <- as.character(structure_label_data[[structure_label_match_column]]) #extract original individual IDs from label file
  valid_structure_label_rows <- !is.na(structure_label_match_ids) & structure_label_match_ids != "" #identify rows containing valid original individual IDs
  structure_label_data <- structure_label_data[valid_structure_label_rows, , drop = FALSE] #remove rows without original individual IDs
  structure_label_match_ids <- as.character(structure_label_data[[structure_label_match_column]]) #extract valid original individual IDs
  if (anyDuplicated(structure_label_match_ids)) stop("Duplicated individual IDs in structure_label_file column: ", structure_label_match_column) #stop if original IDs are duplicated
  structure_label_indices <- match(dataset_ids, structure_label_match_ids) #match Structure individuals to label file
  unmatched_structure_labels <- is.na(structure_label_indices) #identify Structure individuals not found in label file
  if (any(unmatched_structure_labels)) warning("Some Structure individuals were not matched in structure_label_file and retain their original labels: ", paste(dataset_ids[unmatched_structure_labels], collapse = ", ")) #warn about unmatched individuals
  replacement_structure_labels <- rep(NA_character_, length(dataset_ids)) #initialize replacement labels
  replacement_structure_labels[!unmatched_structure_labels] <- as.character(structure_label_data[[structure_label_replacement_column]][structure_label_indices[!unmatched_structure_labels]]) #extract replacement labels for matched individuals
  valid_replacement_labels <- !is.na(replacement_structure_labels) & replacement_structure_labels != "" #identify valid replacement labels
  if (any(!unmatched_structure_labels & !valid_replacement_labels)) warning("Some matched Structure individuals have missing replacement labels and retain their original labels: ", paste(dataset_ids[!unmatched_structure_labels & !valid_replacement_labels], collapse = ", ")) #warn about missing replacement labels
  structure_sample_labels[valid_replacement_labels] <- replacement_structure_labels[valid_replacement_labels] #replace original labels with matching replacement labels
}


## Initialize and process clusters
for (k in kmin:kmax) { #loop through k values
  sp_col_name <- paste0("Cluster_assignment_k", k) #create column name for current k
  prob_cols <- paste0(1:k, "_k", k) #define assignment probability columns for current k
  if (!all(prob_cols %in% colnames(dataset))) stop("Missing probability columns for k = ", k) #stop if probability columns are missing
  probs <- dataset[, prob_cols, drop = F] #extract assignment probabilities
  probs[] <- lapply(probs, as.numeric) #ensure assignment probabilities are numeric
  dataset[[sp_col_name]] <- admixed_individuals #initialize new column with high admixture
  for (row_i in seq_len(nrow(dataset))) { #loop through individuals
    row_probs <- as.numeric(probs[row_i, ]) #extract assignment probabilities for current individual
    if (any(is.na(row_probs))) stop("Missing assignment probability for k = ", k, ", row = ", row_i) #stop if assignment probabilities are missing
    high_clusters <- which(row_probs >= threshold_high) #identify high-confidence cluster assignments
    if (length(high_clusters) == 1) { #assign single cluster if one cluster is above high threshold
      dataset[[sp_col_name]][row_i] <- as.character(high_clusters)
      next
    }
    if (length(high_clusters) > 1) stop("More than one cluster exceeds threshold_high for k = ", k, ", row = ", row_i) #stop if more than one high-confidence cluster is identified
    cluster_order <- order(row_probs, decreasing = T) #order clusters from highest to lowest assignment probability
    cumulative_probs <- cumsum(row_probs[cluster_order]) #calculate cumulative assignment probabilities
    number_shared_clusters <- which(cumulative_probs >= threshold_high)[1] #identify minimum number of clusters needed to reach threshold_high
    if (is.na(number_shared_clusters)) stop("Could not determine cluster assignment for k = ", k, ", row = ", row_i) #stop if cumulative assignment does not reach threshold
    shared_clusters <- sort(cluster_order[seq_len(number_shared_clusters)]) #identify clusters contributing to shared assignment
    if (number_shared_clusters >= 2 && number_shared_clusters <= max_shared_clusters) { #assign shared cluster if two or three clusters are required
      dataset[[sp_col_name]][row_i] <- paste(shared_clusters, collapse = " + ")
    } else {
      dataset[[sp_col_name]][row_i] <- admixed_individuals #assign high admixture if more than three clusters are required
    }
  }
}
dataset_output <- dataset #copy dataset for output
numeric_columns <- sapply(dataset_output, is.numeric) #identify numeric columns for output
dataset_output[numeric_columns] <- lapply(dataset_output[numeric_columns], format_decimal) #format numeric columns for output
combined_csv_file <- file.path(results_dir, paste0(CSV_combined_file_name, ".csv")) #define path for combined CSV output file
utils::write.csv(dataset_output, file = combined_csv_file, row.names = F) #save combined dataset with cluster assignments


## Create Structure-like stacked barplots
for (k in kmin:kmax) { #loop through k values
  prob_cols <- paste0(1:k, "_k", k) #define assignment probability columns for current k
  if (!all(prob_cols %in% colnames(dataset))) stop("Missing probability columns for k = ", k) #stop if probability columns are missing
  structure_proportions <- as.matrix(dataset[, prob_cols, drop = F]) #extract assignment probabilities for current k
  storage.mode(structure_proportions) <- "numeric" #ensure assignment probabilities are numeric
  if (anyNA(structure_proportions) || any(!is.finite(structure_proportions))) stop("Invalid assignment probabilities for Structure-like plot at k = ", k) #stop if probabilities are invalid
  if (any(structure_proportions < 0)) stop("Negative assignment probabilities for Structure-like plot at k = ", k) #stop if probabilities are negative
  structure_row_sums <- rowSums(structure_proportions) #calculate row sums
  if (any(!is.finite(structure_row_sums)) || any(structure_row_sums <= 0)) stop("Invalid assignment probability sums for Structure-like plot at k = ", k) #stop if row sums are invalid
  structure_proportions <- structure_proportions / structure_row_sums #normalize probabilities to exactly sum to one
  structure_sample_names <- structure_sample_labels #extract replacement sample labels in original Structure order
  if (!is.null(structure_sort_by_cluster)) { #sort individuals by dominant cluster if requested
    if (structure_sort_by_cluster > k) stop("structure_sort_by_cluster cannot be larger than k = ", k) #stop if specified first cluster does not exist
    dominant_cluster <- max.col(structure_proportions, ties.method = "first") #identify dominant cluster for each individual
    cluster_order <- c(seq.int(structure_sort_by_cluster, k), if (structure_sort_by_cluster > 1) seq_len(structure_sort_by_cluster - 1)) #define order of dominant clusters
    dominant_cluster_order <- match(dominant_cluster, cluster_order) #match dominant clusters to requested cluster order
    dominant_assignment <- structure_proportions[cbind(seq_len(nrow(structure_proportions)), dominant_cluster)] #extract dominant assignment probability for each individual
    sample_order <- order(dominant_cluster_order, -dominant_assignment) #sort by dominant cluster and decreasing dominant assignment
    structure_proportions <- structure_proportions[sample_order, , drop = F] #reorder assignment probabilities
    structure_sample_names <- structure_sample_names[sample_order] #reorder individual IDs
  }
  cluster_colors <- viridis_palette(k, begin = 0, end = 1) #generate one color for each Structure cluster
  plotting_assignment_coefficients <- apply(cbind(0, structure_proportions), 1, cumsum) #calculate cumulative assignment coefficients for plotting
  structure_svg_filename <- file.path(results_dir, paste0("Structure_plot_K", k, ".svg")) #generate SVG filename
  grDevices::svg(structure_svg_filename, width = structure_plot_width, height = structure_plot_height) #start SVG device
  if (structure_show_individual_labels) { #set margins when individual labels are shown
    graphics::par(mfrow = c(1, 1),
                  mar = c(structure_bottom_margin, structure_left_margin, structure_top_margin, structure_right_margin),
                  oma = c(0, 0, 0, 0))
  } else { #set smaller bottom margin when individual labels are hidden
    graphics::par(mfrow = c(1, 1),
                  mar = c(2, structure_left_margin, structure_top_margin, structure_right_margin),
                  oma = c(0, 0, 0, 0))
  }
  number_of_individuals <- nrow(structure_proportions) #extract number of individuals
  graphics::plot(0,
                 xlim = c(0, number_of_individuals),
                 ylim = c(0, 1),
                 type = "n",
                 ylab = "",
                 xlab = "",
                 xaxt = "n",
                 yaxt = "n",
                 xaxs = "i",
                 yaxs = "i",
                 bty = "l") #create empty Structure-like plot
  graphics::axis(side = 2,
                 at = seq(0, 1, by = 0.2),
                 labels = seq(0, 1, by = 0.2),
                 las = 1) #add y-axis
  if (!is.null(structure_y_axis_title) && structure_y_axis_title != "") { #add y-axis title if requested
    graphics::mtext(structure_y_axis_title,
                    side = 2,
                    line = 3,
                    font = 2)
  }
  for (cluster_index in seq_len(k)) { #loop through clusters
    for (individual_index in seq_len(number_of_individuals)) { #loop through individuals
      if (tolower(structure_bar_lines) == "yes") { #draw colored bar segments with specified borders
        current_bar_border <- structure_bar_line_color
        current_bar_border_thickness <- structure_bar_line_thickness
      } else { #draw colored bar segments without borders
        current_bar_border <- NA
        current_bar_border_thickness <- 1
      }
      graphics::polygon(x = c(individual_index - 1, individual_index, individual_index, individual_index - 1),
                        y = c(plotting_assignment_coefficients[cluster_index, individual_index],
                              plotting_assignment_coefficients[cluster_index, individual_index],
                              plotting_assignment_coefficients[cluster_index + 1, individual_index],
                              plotting_assignment_coefficients[cluster_index + 1, individual_index]),
                        col = cluster_colors[cluster_index],
                        border = current_bar_border,
                        lwd = current_bar_border_thickness) #draw current ancestry segment
    }
  }
  if (tolower(structure_individual_lines) == "yes" && number_of_individuals > 1) { #add vertical separation lines between individuals
    for (individual_index in 2:number_of_individuals) {
      graphics::segments(x0 = individual_index - 1,
                         y0 = 0,
                         x1 = individual_index - 1,
                         y1 = 1,
                         col = structure_individual_line_color,
                         lwd = structure_individual_line_thickness)
    }
  }
  if (structure_show_individual_labels) { #add individual IDs if requested
    graphics::axis(side = 1,
                   at = seq_len(number_of_individuals) - 0.5,
                   labels = structure_sample_names,
                   cex.axis = structure_individual_label_size,
                   las = 2,
                   tick = TRUE)
  }
  graphics::mtext(paste("k =", k),
                  side = 3,
                  line = 0.5,
                  font = 2) #add plot title
  
  grDevices::dev.off() #close SVG device
}


## Ensure Longitude and Latitude are numeric
dataset$Longitude <- suppressWarnings(as.numeric(as.character(dataset$Longitude)))
dataset$Latitude <- suppressWarnings(as.numeric(as.character(dataset$Latitude)))
if (any(!is.finite(dataset$Longitude)) || any(!is.finite(dataset$Latitude))) stop("Longitude and/or Latitude contain missing or non-finite values")


## Create reproducibly jittered coordinates for plotting
plot_Longitude <- dataset$Longitude #copy longitude coordinates
plot_Latitude <- dataset$Latitude #copy latitude coordinates
if (jitter_setting > 0) { #apply jitter if requested
  set.seed(jitter_seed)
  plot_Longitude <- plot_Longitude + stats::runif(nrow(dataset), -jitter_setting, jitter_setting)
  plot_Latitude <- plot_Latitude + stats::runif(nrow(dataset), -jitter_setting, jitter_setting)
}



## Create function to add one admixture pie to an existing map
add_admixture_pie <- function(longitude,
                              latitude,
                              ancestry_proportions,
                              cluster_colors,
                              x_radius,
                              y_radius,
                              border_color = pie_border_color,
                              line_width = pie_border_thickness,
                              number_of_points = 80) {
  ancestry_proportions <- as.numeric(ancestry_proportions) #convert ancestry proportions to numeric
  ancestry_proportions[is.na(ancestry_proportions) | !is.finite(ancestry_proportions) | ancestry_proportions < 0] <- 0 #replace invalid proportions with zero
  if (sum(ancestry_proportions) <= 0) return(invisible(NULL)) #skip pie if no positive ancestry proportions
  ancestry_proportions <- ancestry_proportions / sum(ancestry_proportions) #normalize ancestry proportions to sum to one
  slice_start_angles <- c(0, cumsum(ancestry_proportions)[-length(ancestry_proportions)]) * 2 * pi #calculate start angle of each pie slice
  slice_end_angles <- cumsum(ancestry_proportions) * 2 * pi #calculate end angle of each pie slice
  for (slice_index in seq_along(ancestry_proportions)) { #loop through cluster slices
    if (ancestry_proportions[slice_index] <= 0) next #skip zero-sized slices
    slice_angles <- seq(slice_start_angles[slice_index], slice_end_angles[slice_index], length.out = max(3, ceiling(number_of_points * ancestry_proportions[slice_index]))) #calculate coordinates around current pie slice
    graphics::polygon(x = c(longitude, longitude + x_radius * cos(slice_angles), longitude),
                      y = c(latitude, latitude + y_radius * sin(slice_angles), latitude),
                      col = cluster_colors[slice_index],
                      border = border_color,
                      lwd = line_width) #draw current pie slice
  }
  circle_angles <- seq(0, 2 * pi, length.out = number_of_points) #calculate coordinates of outer circle
  graphics::lines(longitude + x_radius * cos(circle_angles),
                  latitude + y_radius * sin(circle_angles),
                  col = border_color,
                  lwd = line_width) #draw outer border around complete pie chart
  return(invisible(NULL))
}



## Plot maps with pie charts of cluster assignment probabilities
for (k in kmin:kmax) { #loop through k values
  prob_cols <- paste0(1:k, "_k", k) #define assignment probability columns for current k
  if (!all(prob_cols %in% colnames(dataset))) stop("Missing probability columns for k = ", k) #stop if probability columns are missing
  ancestry_proportions <- as.matrix(dataset[, prob_cols, drop = F]) #extract assignment probabilities for current k
  storage.mode(ancestry_proportions) <- "numeric" #ensure ancestry probabilities are numeric
  if (anyNA(ancestry_proportions) || any(!is.finite(ancestry_proportions))) stop("Invalid assignment probabilities for k = ", k) #stop if invalid assignment probabilities occur
  if (any(ancestry_proportions < 0)) stop("Negative assignment probabilities for k = ", k) #stop if negative assignment probabilities occur
  ancestry_row_sums <- rowSums(ancestry_proportions) #calculate probability sums
  if (any(!is.finite(ancestry_row_sums)) || any(ancestry_row_sums <= 0)) stop("Invalid assignment probability sums for k = ", k) #stop if invalid probability sums occur
  ancestry_proportions <- ancestry_proportions / ancestry_row_sums #normalize probabilities to exactly sum to one for plotting
  cluster_colors <- viridis_palette(k, begin = 0, end = 1) #generate one color for each Structure cluster
  xlim_range <- range(plot_Longitude, na.rm = T) #calculate longitude range
  ylim_range <- range(plot_Latitude, na.rm = T) #calculate latitude range
  observed_longitude_range <- diff(xlim_range) #calculate observed longitude range before adding map buffer
  observed_latitude_range <- diff(ylim_range) #calculate observed latitude range before adding map buffer
  x_buffer <- observed_longitude_range * buffer_percentage #calculate longitude buffer
  y_buffer <- observed_latitude_range * buffer_percentage #calculate latitude buffer
  if (x_buffer == 0) x_buffer <- 1 #provide longitude buffer if all samples have identical longitude
  if (y_buffer == 0) y_buffer <- 1 #provide latitude buffer if all samples have identical latitude
  xlim_range <- c(xlim_range[1] - x_buffer, xlim_range[2] + x_buffer) #apply longitude buffer
  ylim_range <- c(ylim_range[1] - y_buffer, ylim_range[2] + y_buffer) #apply latitude buffer
  pie_reference_range <- max(observed_longitude_range, observed_latitude_range) #define reference coordinate range for pie size
  if (!is.finite(pie_reference_range) || pie_reference_range <= 0) pie_reference_range <- 1 #provide fallback reference range
  pie_radius <- pie_size * 0.01 * pie_reference_range #calculate pie radius in map units
  svg_filename <- file.path(results_dir, paste0("Map_Structure_", k, ".svg")) #generate SVG filename based on current k
  grDevices::svg(svg_filename, width = width_plot, height = height_plot) #start SVG device
  graphics::par(mfrow = c(1, 1),
                oma = c(0, 0, 1.5, 0),
                mar = c(0, 0, 0, 0)) #set map margins
  current_plot_region_size_inches <- graphics::par("pin") #extract available plotting-region dimensions
  longitude_range <- diff(xlim_range) #calculate buffered longitude range
  latitude_range <- diff(ylim_range) #calculate buffered latitude range
  mean_map_latitude <- mean(ylim_range) #calculate mean map latitude
  longitude_latitude_correction <- cos(mean_map_latitude * pi / 180) #calculate geographic longitude correction
  if (!is.finite(longitude_latitude_correction) || longitude_latitude_correction <= 0) longitude_latitude_correction <- 1 #set fallback geographic correction
  target_height_width_ratio <- latitude_range / (longitude_range * longitude_latitude_correction) #calculate target geographic aspect ratio
  adjusted_plot_width_inches <- current_plot_region_size_inches[1] #set initial plot width
  adjusted_plot_height_inches <- adjusted_plot_width_inches * target_height_width_ratio #calculate required plot height
  if (adjusted_plot_height_inches > current_plot_region_size_inches[2]) { #adjust dimensions if required height exceeds available height
    adjusted_plot_height_inches <- current_plot_region_size_inches[2]
    adjusted_plot_width_inches <- adjusted_plot_height_inches / target_height_width_ratio
  }
  graphics::par(pin = c(adjusted_plot_width_inches, adjusted_plot_height_inches)) #preserve geographic aspect ratio
  graphics::plot.new() #create empty plotting window
  graphics::plot.window(xlim = xlim_range,
                        ylim = ylim_range,
                        xaxs = "i",
                        yaxs = "i") #set map plotting range
  plot_coordinate_limits <- graphics::par("usr") #extract plotting coordinate limits
  plot_region_size_inches <- graphics::par("pin") #extract final plotting-region dimensions
  x_inches_per_unit <- plot_region_size_inches[1] / diff(plot_coordinate_limits[1:2]) #calculate horizontal plotting scale
  y_inches_per_unit <- plot_region_size_inches[2] / diff(plot_coordinate_limits[3:4]) #calculate vertical plotting scale
  if (!is.finite(x_inches_per_unit) || !is.finite(y_inches_per_unit) || x_inches_per_unit <= 0 || y_inches_per_unit <= 0) { #set fallback pie radii if plotting scales cannot be calculated
    pie_radius_x <- pie_radius
    pie_radius_y <- pie_radius
  } else {
    pie_radius_x <- pie_radius * y_inches_per_unit / x_inches_per_unit #adjust longitude radius so pies remain circular
    pie_radius_y <- pie_radius #retain latitude radius
  }
  maps::map("world",
            regions = mapping_regions,
            fill = T,
            col = map_filling_color,
            border = country_border_color,
            lwd = country_border_thickness,
            xlim = xlim_range,
            ylim = ylim_range,
            add = T) #add country map
  if (tolower(states) == "yes" && "USA" %in% mapping_regions) { #add US state boundaries if requested
    maps::map("state",
              add = T,
              col = state_border_color,
              lwd = state_border_thickness,
              xlim = xlim_range,
              ylim = ylim_range)
  }
  for (sample_index in seq_len(nrow(ancestry_proportions))) { #loop through individuals
    add_admixture_pie(longitude = plot_Longitude[sample_index],
                      latitude = plot_Latitude[sample_index],
                      ancestry_proportions = ancestry_proportions[sample_index, ],
                      cluster_colors = cluster_colors,
                      x_radius = pie_radius_x,
                      y_radius = pie_radius_y) #add ancestry pie for current individual
  }
  graphics::legend(legend_position,
                   legend = paste("Cluster", seq_len(k)),
                   fill = cluster_colors,
                   border = pie_border_color,
                   title = "Cluster",
                   bg = "white",
                   box.col = "black",
                   cex = 0.8) #add cluster legend
  graphics::mtext(paste("k =", k),
                  side = 3,
                  outer = T,
                  line = 0,
                  font = 2) #add map title
  grDevices::dev.off() #close SVG device
}


## Read combined Clumpak results
Clumpak_results_combined <- utils::read.csv(file = combined_csv_file, header = T, stringsAsFactors = F) #read combined Clumpak results file
head(Clumpak_results_combined) #check combined Clumpak results
nrow(Clumpak_results_combined) #check number of individuals in combined Clumpak results
colnames(Clumpak_results_combined) #check column names in combined Clumpak results