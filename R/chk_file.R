#' Evaluates that path argument so that if the specified path is NULL a pop window will be displayed to choose file and then the file will be validated for its type.
#'
#' @param path (Character) the path that the excel file can be read from or null in case the user wants to have a file choose window displayed.
#' @return the path that the user chose
#'
#' @author
#' Agapios Panos <panosagapios@gmail.com>
#'
#' @importFrom tools file_ext

chk_file <- function(path){
    if (any(is.null(path), is.na(path), !is.character(path))){
        print('Please choose the excel file to import data...')
        path <- choose_file()
    } else if (!exists(path)) { # we have to check if exists after passing the check for NULL OR NA as it produces errors if we check exists(NULL) or exists(NA)
        print('The specified path for the excel file is not valid. Please choose a valid path for the excel file...')
        path <- choose_file()
    }
    return (path)
}

choose_file <- function(){
    path <- file.choose();
    if (is.na(path))
        stop("You did not choose an xlsx file to get data from. Rerun the command and choose a file")
    if (tools::file_ext(path) != 'xlsx' & tools::file_ext(path) != 'xls')
        stop("You must choose an xlsx or xls file")
    return (path)
}
