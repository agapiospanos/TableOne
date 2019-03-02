#' Imports the data from the excel file
#'
#' @param path (Character) the path that the excel file can be read from. By default it is set to NULL so that a pop up window will ask for the path.
#' @param sheet (Character) (default: NULL) the sheet inside excel file that the data are stored. By default it gets the first one.
#' @param excel.col.name (Character) the column names as specifed in the excel.
#' @param output.var.names (Character) the variable names as they will be displayed in the output.
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

import_excel_data <- function(path = NULL, sheet = NULL, excel.col.name, output.var.names, group.col.name, control.value, treatment.value){

    # check the imported file for its extension and display file choose window in case the path variable is NULL
    path <- chk_file(path)

    # importing data from the excel file
    data <- read_excel(path, sheet)

    # checking if the column specified in the group.col.name variable exists in the excel
    if (!group.col.name %in% names(data)) {
        stop('The name "', group.col.name, '" that you specified in argument group.col.name does not exist in the excel file you loaded.')
    }

    # checking if the columns specified in the excel.col.name variable exist in the excel
    colnames.verification <- excel.col.name %in% names(data)
    if (!any(colnames.verification)) {
        unmatched.pos <- which(colnames.verification == F)
        stop('The columns ', excel.col.name[unmatched.pos], ' are not present in the excel file that you loaded.')
    }

    # initializing final subset which includes only the desired columns from the excel.
    subset <- data.frame(matrix(nrow = nrow(data), ncol = length(output.var.names)))
    # setting column names to the final data data.frame
    names(subset) <- excel.col.name

    # adding the desired data to subset data.frame
    for (i in 1:length(excel.col.name)) {
        subset[i] <- data[excel.col.name[i]]
    }

    # initialize a grouping column
    group <- data.frame(matrix(nrow = nrow(data), ncol = 1))

    # get values for the column that groups
    group <- data[group.col.name]

    # initializing output list
    output <- list()

    output$data <- subset
    output$group <- group

    return(output)
}
