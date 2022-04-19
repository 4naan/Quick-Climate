

from openpyxl import load_workbook;
#load in excel data
workbook_active_date= load_workbook(filename=r"C:\Users\jcass\Documents\Python\WeatherMe\IDCJDW7007.202204.xlsx");
workbook_historic= load_workbook(filename=r"C:\Users\jcass\Documents\Python\WeatherMe\IDCJCM0033_094008.xlsx");
sheet_active_date = workbook_active_date.active;
sheet_historic = workbook_historic.active;
#
input();

#python C:\Users\jcass\Documents\Python\WeatherMe\open.py