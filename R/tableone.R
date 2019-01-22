#' Exports the table 1 in a word file
#'
#' @param columns (Character) the column names as specifed in the excel.
#' @param var.names (Character) the variable names as they will be displayed in the output.
#' @param percentages (Boolean) a vector of boolean values that specifies whether that exported variable value will be displayed in a \% format.
#' @param or.display (Boolean) a vector of boolean values that specifies whether the Odds Ratio will be calculated for each variable.
#' @param group.col.name (Character) the column name that specifies the group (control or treatment).
#' @param control.value (Character) the value that specifies the control group. (e.g. 0 or 'control').
#' @param treatment.value (Character) the value that specifies the treatment group (e.g. 1 or 'treatment').
#' @param path (Character) (optional) the path that the excel file can be read from. By default it is set to NULL so that a pop up window will ask for the path.
#' @param export.path (Character) (optional) the path that the Word Document will be exported to.
#' @param sheet (Character) (optional) (default: NULL) the sheet inside excel file that the data are stored. By default it gets the first one.
#' @param tableone.col.names (Character) (optional) a vector for the column names of the exported table. Default are: c(' ', 'Treatment Group', 'Control Group', '95\% CI', 'Z-value', 'p-value', 'OR')
#' @param export.filename (Character) (optional) the name of the file that will be exported. Do not include the .docx extension. (default filename is TableOne.docx)
#'
#' @author
#' Agapios Panos <panosagapios@gmail.com>
#'
#' @importFrom easycsv choose_dir
#' @importFrom officer read_docx body_add_blocks block_list
#' @importFrom flextable regulartable theme_vanilla
#' @importFrom stats sd t.test
#' @export
#'

tableone <- function(columns, var.names, percentages, or.display, group.col.name, control.value, treatment.value, path = NULL, export.path = NULL, sheet = NULL, tableone.col.names = NULL, export.filename = NULL) {

    # checking if the user specified an export.path argument. If not, a prompt window will be displayed to ask for a path.
    if (is.null(export.path)) {
        print('Please choose a folder to export the Word Document...')
        export.path <- choose_dir()
    }

    # checking if the export.path is specified
    if (length(export.path) == 0) {
        stop('You have to choose a folder to have the Word Document exported')
    } else if (is.na(export.path)) {
        stop('You did not specify a valid export.path argument')
    } else {
        cat('You selected the folder ', export.path, '\n')
    }

    # checking if export.filename is specified
    if (length(export.filename) == 0) {
        if (!is.null(export.filename)) {
            stop('The export filename must be of length > 0')
        } else {
            export.filename <- 'TableOne'
            cat('No export.filename is specified. The file will be saved as TableOne.docx.', '\n')
        }
    } else if (is.na(export.filename)) {
        stop('You did not specify a valid export.filename argument')
    }

    # getting data from the excel file
    excel_data <- import_excel(path, sheet, columns, var.names, group.col.name, control.value, treatment.value)

    # initializing the data.frame that will be exported to Word
    table.to.export <- data.frame( matrix( ncol = 7, nrow = length(var.names) ) )

    # checking if the user specified custom column names for the exported table
    if (is.null(tableone.col.names)) {
        tableone.col.names <- c(' ', 'Treatment Group', 'Control Group', '95% CI', 'Z-value', 'p-value', 'OR')
    } else {
        if (length(tableone.col.names) < 7) {
            stop('you must provide a vector with 7 column names for the argument tableone.col.names. Use " " inside the vector to keep empty column names')
        } else {
            tableone.col.names[which(tableone.col.names == '' | is.na(tableone.col.names))] <- ' '
        }
    }
    names(table.to.export) <- tableone.col.names

    # number of participants in the treatment group
    n.t <- length(excel_data$treatment)

    # number of participants in the control group
    n.c <- length(excel_data$control)

    # generating the table that will be exported
    for (i in 1:length(var.names)) {
        if (percentages[i]) {

            test.values <- t.test(unlist(excel_data$treatment[i]), unlist(excel_data$control[i]))

            # standard error of mean for control group
            c.p.se <- round(sqrt((test.values$estimate[2]*(1-test.values$estimate[2]))/n.c), digits = 2)

            # standard error of mean for treatment group
            t.p.se <- round(sqrt((test.values$estimate[1]*(1-test.values$estimate[1]))/n.t), digits = 2)

            # creating the table columns
            table.to.export[i,1] <- paste(var.names[i], '(%)')
            table.to.export[i,2] <- paste0(round(test.values$estimate[1]*100, digits = 2), '% (', t.p.se, ')')
            table.to.export[i,3] <- paste0(round(test.values$estimate[2]*100, digits = 2), '% (', c.p.se, ')')
            table.to.export[i,4] <- paste0('[', round(test.values$conf.int[1], digits = 2), '-', round(test.values$conf.int[2], digits = 2), ']')
            table.to.export[i,5] <- test.values$statistic
            table.to.export[i,6] <- test.values$p.value

        } else {

            test.values <- t.test(unlist(excel_data$treatment[i]), unlist(excel_data$control[i]))

            # standard deviation for control group
            c.sd <- round(sd(unlist(excel_data$control[i])), digits = 2)

            # standard deviation for the treatment group
            t.sd <- round(sd(unlist(excel_data$treatment[i])), digits = 2)

            # creating the table columns
            table.to.export[i,1] <- var.names[i]
            table.to.export[i,2] <- paste0(round(test.values$estimate[1], digits = 2), '\U00B1', t.sd)
            table.to.export[i,3] <- paste0(round(test.values$estimate[2], digits = 2), '\U00B1', c.sd)
            table.to.export[i,4] <- paste0('[', round(test.values$conf.int[1], digits = 2), '-', round(test.values$conf.int[2], digits = 2), ']')
            table.to.export[i,5] <- test.values$statistic
            table.to.export[i,6] <- test.values$p.value
        }

        # display odds ratio
        if (or.display[i]) {

            t.events <- length(which(unlist(excel_data$treatment[i]) == 1))
            t.no_events <- length(which(unlist(excel_data$treatment[i]) == 0))
            c.events <- length(which(unlist(excel_data$control[i]) == 1))
            c.no_events <- length(which(unlist(excel_data$control[i]) == 0))

            or.value <- (c.no_events * t.events) / (t.no_events * c.events)
            var.log.or <- 1/c.no_events + 1/t.no_events + 1/t.events + 1/c.events

            or.ci.lower <- log(or.value) - 1.96*sqrt(var.log.or)
            or.ci.upper <- log(or.value) + 1.96*sqrt(var.log.or)

            table.to.export[i,7] <- paste0(round(or.value, digits = 2), ' [', round(or.ci.lower, digits = 2), '-', round(or.ci.upper, digits = 2), ']')

        } else {

            table.to.export[i,7] <- ''
        }

    }

    # initializing a blank docx document
    doc <- read_docx()

    # converting the data.frame to regulartable so that we can apply styling using a theme from flextable package
    output <- regulartable(table.to.export)

    # adding the table to the document body
    body_add_blocks(doc, blocks = block_list(output))

    # applying the styling to the table
    output <- theme_vanilla(output)

    # finaly exporting the docx file
    print(doc, target = paste0(export.path, '\\', export.filename, '.docx'))

    print('The file is exported successfully! You can find it in the following directory:')
    print(paste0(export.path, export.filename, '.docx'))
}
