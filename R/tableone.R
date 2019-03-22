#' Exports the table 1 in a word file
#'
#' @param excel.col.names (Character) the column names as specifed in the excel.
#' @param output.var.names (Character) the variable names as they will be displayed in the output.
#' @param dichotomous (Character) a vector of characters values that specifies which of the excel columns contain dichotomous data.
#' @param group.col.name (Character) the column name that specifies the group (control or treatment).
#' @param control.value (Character) the value that specifies the control group. (e.g. 0 or 'control').
#' @param treatment.value (Character) the value that specifies the treatment group (e.g. 1 or 'treatment').
#' @param excel.path (Character) (optional) the path that the excel file can be read from. By default it is set to NULL so that a pop up window will ask for the path.
#' @param export.path (Character) (optional) the path that the Word Document will be exported to.
#' @param sheet (Character) (optional) (default: NULL) the sheet inside excel file that the data are stored. By default it gets the first one.
#' @param tableone.col.names (Character) (optional) a vector for the column names of the exported table. Default are: c(' ', 'Treatment Group', 'Control Group', 'Mean Difference', 'OR', 'z-value', 'p-value')
#' @param export.filename (Character) (optional) the name of the file that will be exported. Do not include the .docx extension. (default filename is TableOne.docx)
#'
#' @author
#' Agapios Panos <panosagapios@gmail.com>
#'
#' @importFrom easycsv choose_dir
#' @importFrom officer read_docx body_add_blocks block_list fp_border body_end_section_landscape
#' @importFrom flextable regulartable theme_zebra autofit vline vline_right align bold
#' @importFrom stats sd t.test
#' @export
#'

tableone <- function(excel.col.names, output.var.names, dichotomous, group.col.name, control.value, treatment.value, excel.path = NULL, export.path = NULL, sheet = NULL, tableone.col.names = NULL, export.filename = NULL) {

    # checking if the user specified an export.path argument. If not, a prompt window will be displayed to ask for a path.
    if (is.null(export.path)) {
        print('Please choose a folder to export the Word Document...')
        export.path <- choose_dir()
    }

    # checking if the user has supplied the correct amount of names for all variables
    if (length(output.var.names) != length(excel.col.names)) {
        stop("The length of the variables 'output.var.names' and 'excel.col.names' is not equal. Please make sure you have entered the same number of items in these vectors.")
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
    imported_data <- import_excel_data(excel.path, sheet, excel.col.names, output.var.names, group.col.name, control.value, treatment.value)

    excel_data <- imported_data$data
    group <- imported_data$group


    # # initializing final subset which includes only the desired columns from the excel.
    # subset <- data.frame(matrix(nrow = nrow(data), ncol = length(var.names)))
    # # setting column names to the final data data.frame
    # names(subset) <- var.names

    # initializing the data.frame that will be exported to Word
    table.to.export <- data.frame( matrix( ncol = 7, nrow = length(output.var.names) ) )

    # keeping treatment and control group values in a separate var
    treatment <- excel_data[which(group == treatment.value),]
    control <- excel_data[which(group == control.value),]

    # number of participants in the treatment group
    n.t <- nrow(treatment)

    # number of participants in the control group
    n.c <- nrow(control)

    # checking if the user specified custom column names for the exported table
    if (is.null(tableone.col.names)) {
        tableone.col.names <- c(' ', 'Treatment Group', 'Control Group', 'Mean Difference', 'OR', 'z-value', 'p-value')
    } else {
        if (length(tableone.col.names) < 7) {
            stop('you must provide a vector with 7 column names for the argument tableone.col.names. Use " " inside the vector to keep empty column names')
        } else {
            tableone.col.names[which(tableone.col.names == '' | is.na(tableone.col.names))] <- ' '
        }
    }

    # adding the number of participants in each group
    tableone.col.names[2] <- paste0(tableone.col.names[2], ' (n=', n.t, ')')
    tableone.col.names[3] <- paste0(tableone.col.names[3], ' (n=', n.c, ')')

    # adding the names in table that will be exported
    names(table.to.export) <- tableone.col.names

    # generating the table that will be exported
    for (i in 1:length(excel.col.names)) {

        if (excel.col.names[i] %in% dichotomous) {

            # get data as factors for treatment group
            t.factors <- table(treatment[i])
            t.f1name <- names(t.factors[1])
            t.f1count <- t.factors[1]
            t.f2name <- names(t.factors[2])
            t.f2count <- t.factors[2]
            # t.f2name <- names(t.factors[2]) TODO remove
            t.totalcount <- t.f1count + t.f2count
            # t.f1percent <- t.f1count / t.totalcount TODO remove
            t.f2percent <- t.f2count / t.totalcount

            # get data as factors for control group
            c.factors <- table(control[i])
            c.f1name <- names(c.factors[1])
            c.f1count <- c.factors[1]
            # c.f1name <- names(c.factors[1]) TODO remove
            c.f2name <- names(c.factors[2])
            c.f2count <- c.factors[2]
            # c.f2name <- names(c.factors[2]) TODO remove
            c.totalcount <- c.f1count + c.f2count
            # c.f1percent <- c.f1count / c.totalcount TODO remove
            c.f2percent <- c.f2count / c.totalcount

            # creating the table columns
            # column name
            # TODO IMPROVEMENT add the option to choose which factor to consider for every dichotomous var as the baseline.
            if (output.var.names[i] == '') {
                varname <- t.f2name # we get the 2nd factor's name because we use the 2nd factor as a baseline.
            } else {
                varname <- output.var.names[i]
            }


            # calculate and display OR
            t.events <- length(which(unlist(treatment[i]) == t.f2name)) # we use f2name because we have the 2nd factor as the baseline.
            t.no_events <- length(which(unlist(treatment[i]) == t.f1name))
            c.events <- length(which(unlist(control[i]) == c.f2name))
            c.no_events <- length(which(unlist(control[i]) == c.f1name))

            or.value <- (c.no_events * t.events) / (t.no_events * c.events)
            var.log.or <- 1/c.no_events + 1/t.no_events + 1/t.events + 1/c.events

            logor.ci.lower <- log(or.value) - 1.96*sqrt(var.log.or)
            logor.ci.upper <- log(or.value) + 1.96*sqrt(var.log.or)

            zval <- round(log(or.value)/sqrt(var.log.or), digits = 3)

            pval <- round(2*(1 - pnorm(zval)), digits = 3)

            table.to.export[i,1] <- paste(varname, '(%)')
            # treatment group value - percentage
            table.to.export[i,2] <- paste0(format(round(t.f2percent*100, digits = 2), nsmall = 2), '%') # we use f2percent because we have the 2nd factor as the baseline.
            # control group value - percentage
            table.to.export[i,3] <- paste0(format(round(c.f2percent*100, digits = 2), nsmall = 2), '%')
            # mean difference with 95 percent confidence interval
            table.to.export[i,4] <- ''
            # odds ratio - 95 percent OR confidence interval
            table.to.export[i,5] <- paste0(format(round(or.value, digits = 2), nsmall = 2), ' [', format(round(exp(logor.ci.lower), digits = 2), nsmall = 2), ', ', format(round(exp(logor.ci.upper), digits = 2), nsmall = 2), ']')
            # z-value
            table.to.export[i,6] <- format(zval, nsmall = 3)
            # p-value
            table.to.export[i,7] <- format(pval, nsmall = 3)


        } else { # continuous data case

            test.values <- t.test(unlist(excel_data[excel.col.names[i]])~unlist(group), excel_data)

            # standard deviation for control group
            c.sd <- round(sd(unlist(control[i])), digits = 2)

            # standard deviation for the treatment group
            t.sd <- round(sd(unlist(treatment[i])), digits = 2)

            # mean for control group
            c.mean <- test.values$estimate[1]

            # mean for treatment group
            t.mean <- test.values$estimate[2]

            # mean difference
            mean.difference <- t.mean - c.mean

            # p-value
            pval <- round(test.values$p.value, digits = 3)
            if (pval < 0.001)
                pval <- '<0.001'

            # creating the table columns
            # column name
            table.to.export[i,1] <- output.var.names[i]
            # treatment group value with standard deviation
            table.to.export[i,2] <- paste0(format(round(t.mean, digits = 2), nsmall = 2), '\U00B1', format(t.sd, nsmall = 2))
            # control group value with standard deviation
            table.to.export[i,3] <- paste0(format(round(c.mean, digits = 2), nsmall = 2), '\U00B1', format(c.sd, nsmall = 2))
            # mean difference with 95 percent confidence interval
            table.to.export[i,4] <- paste0(format(round(mean.difference, digits = 2), nsmall = 2), ' [', format(round(test.values$conf.int[1], digits = 2), nsmall = 2), ', ', format(round(test.values$conf.int[2], digits = 2), nsmall = 2), ']')
            # leaving empty the odds ratio column
            table.to.export[i,5] <- ''
            # z-value
            table.to.export[i,6] <- format(round(test.values$statistic, digits = 3), nsmall = 3)
            # p-value
            table.to.export[i,7] <- format(pval, nsmall = 3)

        }
    }

    # initializing a blank docx document
    doc <- read_docx()

    # converting the data.frame to regulartable so that we can apply styling using a theme from flextable package
    output <- regulartable(table.to.export)

    # auto-adjust width for table columns
    output <- autofit(output, add_w = 0.2, add_h = 0)

    # applying the styling to the table
    output <- theme_zebra(output, odd_header = "#CFCFCF", odd_body = "#F8F8F8",
                          even_header = "transparent", even_body = "transparent")

    # adding vertical borders between columns
    output <- vline( output, border = fp_border(color = "gray80", width = 1), part = "all" )
    # remove right table border
    output <- vline_right(output, border = fp_border(width = 0), part = "all")
    # align text to center
    output <- align(output, align = "center", part = "all")
    # align first column to right
    output <- align(output, j = 1, align = "right", part = "body")
    # make header bold
    output <- bold(output, bold = TRUE, part = "header")
    # make first column bold
    output <- bold(output, j = 1, bold = TRUE, part = "body")

    # adding the table to the document body
    body_add_blocks(doc, blocks = block_list(output))

    # make the orientation landscape
    body_end_section_landscape(doc)

    # finaly exporting the docx file
    print(doc, target = paste0(export.path, '/', export.filename, '.docx'))

    print('The file is exported successfully! You can find it in the following directory:')
    print(paste0(export.path, '/', export.filename, '.docx'))
}
