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
  excel.col.names = c('age', 'gender', 'systbp', 'diastbp', 'hypertension'),
  output.var.names = c('Age', 'Females', 'Systolic Blood Pressure', 'Diastolic Blood Pressure', 'Hypertension')
)
```
