#' Imports the data from the excel file
#'
#' @param path (Character) the path that the excel file can be read from. By default it is set to NULL so that a pop up window will ask for the path.
#' @param sheet (Character) (default: NULL) the sheet inside excel file that the data are stored. By default it gets the first one.
#' @param columns (Character) the column names as specifed in the excel.
#' @param var.names (Character) the variable names as they will be displayed in the output.
#' @param group.col.name (Character) the column name that specifies the group (control or treatment)
#' @param control.value (Character) the value that specifies the control group. (e.g. 0 or 'control')
#' @param treatment.value (Character) the value that specifies the treatment group (e.g. 1 or 'treatment')
#' @return data.frame with the data.
#'
#' @author
#' Agapios Panos <panosagapios@gmail.com>
#'
#' @importFrom readxl read_excel
#' @export

import_excel <- function(path = NULL, sheet = NULL, columns, var.names, group.col.name, control.value, treatment.value){

    # checking if the length of variable var.names and variable columns are the same
    if (length(var.names) != length(columns)){
        stop("The length of the variables 'var.names' and 'columns' is not equal. Please make sure you have entered the same number of items in these vectors.")
    }

    # check the imported file for its extension and display file choose window in case the path variable is NULL
    path <- chk_file(path)

    # importing data from the excel file
    data <- read_excel(path, sheet)

    # getting row indexes that are related to the treatment group
    treatment.ind <- which(data[group.col.name] == treatment.value)
    # getting records only for treatment group
    treatment.group <- data[treatment.ind, ]
    # counting the number of participants in the treatment group
    treat.count <- nrow(treatment.group)
    # initializing final subset of treatment group.
    treatment.group.subset <- data.frame(matrix(nrow = treat.count, ncol = length(var.names)))
    # setting column names to the final data data.frame
    names(treatment.group.subset) <- var.names
    # adding the desired data to the treatment.group.subset data.frame
    for (i in 1:length(columns)) {
        treatment.group.subset[i] <- treatment.group[columns[i]]
    }
    # getting row indexes that are related to the control group
    control.ind <- which(data[group.col.name] == control.value)
    # getting records only for treatment group
    control.group <- data[control.ind, ]
    # counting the number of participants in the treatment group
    control.count <- nrow(control.group)
    # initializing final subset of treatment group.
    control.group.subset <- data.frame(matrix(nrow = control.count, ncol = length(var.names)))
    # setting column names to the final data data.frame
    names(control.group.subset) <- var.names
    # adding the desired data to the treatment.group.subset data.frame
    for (i in 1:length(columns)) {
        control.group.subset[i] <- control.group[columns[i]]
    }

    # initializing list to return data from function and assigning data
    output <- list()
    output$treatment <- treatment.group.subset
    output$control <- control.group.subset

    return(output)
}
