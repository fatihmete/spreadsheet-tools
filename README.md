
![st](https://github.com/fatihmete/spreadsheet-tools/blob/master/st/icons/sih.png)

# Spreadsheet Tools
This application allows you 
- open, view and edit small to large (more than 1,048,576 rows) Excel and CSV files,
- merge Excel/CSV file into single file,
- split Excel/CSV file into equal parts,
- to create a single data set from multiple excel files in the same template,
- to create multiple excel files in the same template from a single data set
- run any python code in Python shell.

# Contents
- [Installation](https://github.com/fatihmete/spreadsheet-tools#installation)
- [Usage](https://github.com/fatihmete/spreadsheet-tools#usage)
	- [Excel/CSV Viewer](https://github.com/fatihmete/spreadsheet-tools#excelcsv-viewer)
	- [Merge Excel/CSV](https://github.com/fatihmete/spreadsheet-tools#merge-excelcsv)
	- [Split Excel/CSV](https://github.com/fatihmete/spreadsheet-tools#split-excelcsv)
	- [Multiple Excel Reader](https://github.com/fatihmete/spreadsheet-tools#multiple-excel-reader)
	- [Multiple Excel Writer](https://github.com/fatihmete/spreadsheet-tools#multiple-excel-writer)
	- [Python Shell](https://github.com/fatihmete/spreadsheet-tools#python-shell)
- [Example usage](https://github.com/fatihmete/spreadsheet-tools#example-usage)
- [Open Source Licenses](https://github.com/fatihmete/spreadsheet-tools#open-source-licenses)
- [TODO](https://github.com/fatihmete/spreadsheet-tools#todo)


# Installation
## Python package

Install package with pip:

`pip install spreadsheet-tools --upgrade`

in shell run command:

`spreadsheet-tools`

## Windows Users Download Packages
[Click here](https://github.com/fatihmete/spreadsheet-tools/releases) to download packages. After download run st.exe

## Run from code
Clone the repository:

`git clone https://github.com/fatihmete/spreadsheet-tools`

Change directory:

`cd spreadsheet-tools`

Before running you have to install required packages:

`pip install requirements.txt`

Finally run:

`python st.py`

# Usage
## Excel/CSV Viewer

![viewer](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/viewer.png)

You can view and edit Excel/CSV file with this screen. To select Excel/CSV file please click "..." button at top right. If you open CSV file, It asks CSV seperator of file from you. You can navigate in data with **Next**, **Prev**, **First** and **Last** button at bottom. Also you can change count of rows will be shown on screen with change values of **Show rows**.
This screen allow you filter loaded file with pandas query, drop selected cols and rows, run python code (pandas functions), save result data.

### Query
You can write pandas query to filter data. To apply filter press Enter key. It uses python as engine. Below are examples that you can use in the [titanic dataset](https://www.openml.org/d/40945).

Filter only who survived data:

`survived == 1`

Filter who is survived and female data:

`survived == 1 and sex=="female"`

Filter who is survived, female and name contains "Becker" or "Wells":

`survived == 1 and sex=="female" and (name.str.contains("Becker") or name.str.contains("Wells"))`

Filter who is survived, female, name contains "Becker" or "Wells" and pclass + sibsp greater than 2:

`survived == 1 and sex=="female" and (name.str.contains("Becker") or name.str.contains("Wells")) and pclass + sibsp > 2`

You can use any pandas function (e.g. .isna(), isnull()) that is supported in pandas query.

### Drop Cols/Rows

To drop columns/rows firstly select columns/rows will be removed then press Drop Cols/Rows button at top right. This operations can't be undo so be careful. But It's not affect original data.

### Run Python Code

You can edit your data with python code. For example you can create new columns with functions (like abs(), min(), max()), fill Nan values (fillna()), split columns with seperator.
Your limit is your pandas/python knowledge. You can reach your data with predefined **df** variable (Pandas DataFrame) and pandas with predefined **pd** variable.

Below are examples that you can use in the titanic dataset.

To create new column that sum of pclass and survived:

`df["sum"] = df["pclass"] + df["survived"]`

Create last name and first name column using str.split function:

`df[["last_name","first_name"]] = df["name"].str.split(",", expand=True)`

`fare` column in the Titanic dataset contains "?" for nan values and column type is Object. To fix this:

`df["fare"] = pd.to_numeric(df["fare"], errors="coerce")`

Or you can set new value for "?" values in `home.dest` column:

`df["home.dest"] = df["home.dest"].str.replace("?","Anywhere")`

Note: You can't directly edit **df** variable.

### Saving Data

To save the filtered and edited data, please click the **Save Data** button at the bottom left.

## Merge Excel/CSV

You can merge more than one Excel/CSV files in single data file. All files will be append one after the other. You can select which files will be read in input files path. It is possible to set *.xlsx, *.csv or mix type. Also you can set csv seperator for input files.

![merger](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/merger.png)

## Split Excel/CSV

You can split a Excel/CSV file into parts containing the number of lines you want.

![merger](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/splitter.png)

You can select file that will be splitted with **...** button. After you have to select **Output Files Path** where you new files will be save.
You can change format of new files either *.xlsx or *.csv . Row number default set 1000, you can change it. Also you can set input file and out files csv seperator.

## Multiple Excel Reader

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/reader.png)

You can create a single data set from multiple excel files in the same template with this screen.

### Adding Sheets
First you have to add sheets in your excel file. 

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/add_new_sheet.png)

Please write sheet name (be sure it's correct!) then press **Add New Sheet** button. If you want to delete sheet, select related sheet after press **Delete Selected Sheet** button.

### Adding Rules
To add new rule please press **Add New Rule** button then select sheet name, set cell and column name.

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/add_new_rule.png)

While reading excel files merw uses this rules. For example, supposing our first rule is "Sheet1", "B1" and "NAME", merw will open Sheet1 of excel file and get "B1" cell value. After write this value on "NAME" column of output file. There isn't rule limit. If you want to delete rule, select related rule after press **Delete Selected Rule** button.

### Setting Input Files Path and Output File

**Input Files Path** is where your excel files at located. **Output File** can be xlsx or csv format, it is single dataset will be created from excel files at  **Input Files Path**.

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/run.png)

To set **Input Files Path** and **Output File**, please press **...** button where right of them.

Finally press **Run Rules** button, it creates single dataset (**Output File**). If file format of your ouput file is *.csv, application asks you csv seperator. You can set any value as seperator.

### Saving and Loading Rules

If you want to use rules later, you can save rules with **Save Reading Rules** button. To load rules that priorly you saved, press **Load Reading Rules** button and select rules file.

## Multiple Excel Writer

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/writer.png)

You can create multiple excel files in the same template from a single data set.

### Adding Sheets
First you have to add sheets in your excel file. 

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/add_new_sheet.png)

Please write sheet name (be sure it's correct!) then press **Add New Sheet** button. If you want to delete sheet, select related sheet after press **Delete Selected Sheet** button.

### Adding Rules
To add new rule please press **Add New Rule** button then select sheet name, set cell and column name.

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/add_new_rule_writer.png)

While writing excel files merw uses this rules. For example, supposing our first rule is "NAME", "Sheet1", "B1", merw will read value in "NAME" column of dataset and set B1 cell of Sheet1 of template excel to this value. There isn't rule limit. If you want to delete rule, select related rule after press **Delete Selected Rule** button.

### Setting Input Files Path and Output File

**Output Files Path** is where your new excel files at located. **Input File** can be xlsx or csv format, it is single dataset will be used to create new excel files. **Template File** is template. merw will create copy of this file and fill each copy of file with dataset values.

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/run_writer.png)

To set **Output Files Path** ,  **Inut File**, **Template File** please press **...** button where right of them.

Finally press **Run Rules** button, it creates multiple excel files (1.xlsx, 2.xlsx ...) compatible with template file(at **Output Files Path**). If your file format of your dataset is *.csv, application asks you csv seperator. You can set any value as seperator.

### Saving and Loading Rules

If you want to use rules later, you can save rules with **Save Writing Rules** button. To load rules that priorly you saved, press **Load Writing Rules** button and select rules file.

## Python Shell

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/code.png)

If the other functions not enough for you or you want to make different things, you can run python code inside of the application. It offer interactive python shell for you.
For example you need to get and save currency values from API:
 
```
import json
import pandas as pd
import urllib.request
with urllib.request.urlopen("https://open.exchangerate-api.com/v6/latest") as response:
   content = response.read()
content = json.loads(content)
rates = []
for rate, value in content["rates"].items():
	rates.append([rate, value])
df = pd.DataFrame(rates, columns=["rate","value"])
print(df.head())
df.to_excel(r"rates.xlsx", index=False)
```

# Example usage
## Create ticket (excel file) for titanic passangers
First [download](https://www.openml.org/d/40945) titanic dataset.
Then press Multiple Excel Writer button and load [titanic_write.json](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/titanic_write.json) file. I looks below:

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/titanic_multiple_write_01.PNG)

Then set Output Files Path, Input File (titanic.csv) and Template File ([ticket.xlsx](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/ticket.xlsx)).
Finally press Run Rules. It can creates copy of ticket.xlsx for every passanger in data set at output location.

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/titanic_multiple_write_02.PNG)

## Create dateset from titanic tickets
In previous example we create ticket.xlsx every passenger of Titanic. To create a dataset from this tickets, we use outputs of previously example. Then press Multiple Excel Writer button and load [titanic_read.json](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/titanic_read.json) file. I looks below:

![text](https://github.com/fatihmete/spreadsheet-tools/blob/master/examples/titanic_multiple_read_01.PNG)

Then set **Input Files Path** (where copies of ticket.xlsx files located) and **Output File**.
Finally press Run Rules. It can creates a single dataset from titanic tickets.

# Open Source Licenses

This software uses other software below:
- Python [licence](https://docs.python.org/3/license.html)
- Pyqt5 [license](https://github.com/baoboa/pyqt5/blob/master/LICENSE)
- Pandas [license](https://github.com/pandas-dev/pandas/blob/master/LICENSE)
- Google Material Icons [license](https://github.com/google/material-design-icons/blob/master/LICENSE/)
- Qt5 [license](https://doc.qt.io/qt-5/licensing.html)
- Openpyxl [license](https://github.com/gleeda/openpyxl/blob/master/LICENCE.rst)

# TODO
- Improve GUI
- Add *.xls and other open document files support
- Prebuild package for other OS's
