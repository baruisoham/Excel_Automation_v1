# Excel_Automation_v1
Automating excel
Automates analysis of spending by aggregating the sales in each category in a supermarket. Then further differentiating based on genders. (Done by making a pivot table)

#Usage
Convert `convert-to-exec.py` to exe using
`pyinstaller --onefile convert-to-exec.py`

Move the exe file from `dist` folder and place it in the folder where make-pivot-table.py is present.

Run the `make-pivot-table.py` on `supermarket_sales.xlsx` to generate `pivot_table.xlsx`

Then run the executable file generated, input `month` and name of the file that has pivot table, here `pivot_table.xlsx`

File generated with the name `report_monthName`

