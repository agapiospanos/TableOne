TableOne
================

<img src="man/figures/tableone-logo.png" width=300 align="right" style="margin-left:20px; margin-right: 20px;"/>

Description
-----------

An R package that provides an easy way of exporting the descriptive table 1 that is commonly used in articles to a Word document for easy editing and copying to the article.

Installation
------------

The package is still under active development and has not been submitted to the CRAN yet.

You can install the **TableOne** package from GitHub repository as follows:

Installation using R package **[devtools](https://cran.r-project.org/package=devtools)**:

``` r
install.packages("devtools")
devtools::install_github("agapiospanos/TableOne")
```

Input data format
-----------------

I will soon provide an excel file as a template for the input data format.

Usage example
-------------

The basic function of this package has the following syntax:

``` r
library(TableOne)
tableone(
  dichotomous = c('gender', 'hypertension'),
  group.col.name = 'hyperglycemia',
  control.value = 0,
  treatment.value = 1,
  import.col.names = c('age', 'gender', 'systbp', 'diastbp', 'hypertension'),
  output.var.names = c('Age', 'Females', 'Systolic Blood Pressure', 'Diastolic Blood Pressure', 'Hypertension')
)
```

List of arguments:
------------------

-   import.col.names: Use a vector to specify the column names that will be loaded from the data.frame / excel
-   output.var.names: Use a vector to specify the names of the variables. i.e. how the variables will appear in the exported table
-   dichotomous: (optional) Use a vector to specify the column names that contain dichotomous data
-   ordinal: (optional) Use a vector to specify the column names that contain ordinal data
-   median.iqr: (optional) Use a vector to specify column names with continuous data for which the table will calculate median and IQR instead of mean and SD
-   group.col.name: Specify which column contains the grouping variable.
-   control.value: Specify the value that indicates the control group in the group.col.name
-   treatment.value: Specify the value that indicated the treatment group in the group.col.name
-   data: (optional) Pass a data.frame to this argument on case you want tableOne to load data from it instead of an excel file
-   excel.path: (optional) Specify the path that the excel will be read from. In case you do not specify this argument, a pop up window will be displayed to choose the path
-   export.path: (optional) Specify the path that the excel will be exported to. In case you do not specify this argument, a pop up window will be displayed to choose the path
-   sheet: (optional) Specify the excel sheet number that you need to load data from. In case you do not specify this argument, the first sheet will be used
-   tableone.col.names: (optional) A vector for the column names of the exported table. Default are: c('Variable', 'Treatment Group', 'Control Group', 'p-value', 'Mean Difference', 'OR', 'Test Stat.')
-   export.filename (optional) the name of the file that will be exported. Do not include the .docx extension. (default filename is TableOne.docx)
-   show.stats (optional) A vector that specifies which statistics will be displayed in the produced table One. By default it displays the MD, OR and Test Stat. value as additional statistics except from p-value
-   export.word (optional) (default: TRUE) specify if you want to have a Word Document that contains the table one exported. Whether you export the table to Word or not, you will also get the results in a data.frame format after running the package.
