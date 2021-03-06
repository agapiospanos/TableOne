% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/tableone.R
\name{tableone}
\alias{tableone}
\title{Exports the table 1 in a word file}
\usage{
tableone(import.col.names, output.var.names, dichotomous = c(),
  ordinal = c(), median.iqr = c(), group.col.name, control.value,
  treatment.value, data = NULL, excel.path = NULL,
  export.path = NULL, sheet = NULL, tableone.col.names = NULL,
  export.filename = NULL, show.stats = NULL, export.word = TRUE)
}
\arguments{
\item{import.col.names}{(Character) the column names as specifed in the imported data.frame or excel file.}

\item{output.var.names}{(Character) the variable names as they will be displayed in the output. If not specified we will use the excel column names.}

\item{dichotomous}{(Character) a vector of character values that specifies which of the excel columns contain dichotomous data.}

\item{ordinal}{(Character) a vector of character values that specifies which of the excel columns contain ordinal data.}

\item{median.iqr}{(Character) a vector of character values that specifies for which continuous excel columns we should calculate median and iqr instead of mean and sd.}

\item{group.col.name}{(Character) the column name that specifies the group (control or treatment).}

\item{control.value}{(Character) the value that specifies the control group. (e.g. 0 or 'control').}

\item{treatment.value}{(Character) the value that specifies the treatment group (e.g. 1 or 'treatment').}

\item{data}{(data.frame) (optional) (default: NULL) the data in a data.frame format that will be loaded. If you do not provide data, a pop up window will be displayed asking for an excel file to load data from.}

\item{excel.path}{(Character) (optional) the path that the excel file can be read from. By default it is set to NULL so that a pop up window will ask for the path.}

\item{export.path}{(Character) (optional) the path that the Word Document will be exported to.}

\item{sheet}{(Character) (optional) (default: NULL) the sheet inside excel file that the data are stored. By default it gets the first one.}

\item{tableone.col.names}{(Character) (optional) a vector for the column names of the exported table. Default are: c('Variable', 'Treatment Group', 'Control Group', 'p-value', 'Mean Difference', 'OR', 'Test Stat.')}

\item{export.filename}{(Character) (optional) the name of the file that will be exported. Do not include the .docx extension. (default filename is TableOne.docx)}

\item{show.stats}{(Character) (optional) a vector of characters that specifies which statistics will be displayed in the produced table One. By default it displays the MD, OR and Test Stat. value as additional statistics except from p-value}

\item{export.word}{(Boolean) (optional) (default: TRUE) specify if you want to have a Word Document that contains the table one exported. Whether you export the table to Word or not, you will also get the results in a data.frame format after running the package.}
}
\description{
Exports the table 1 in a word file
}
\author{
Agapios Panos <panosagapios@gmail.com>
}
