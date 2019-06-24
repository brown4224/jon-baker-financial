Install the Virtual Env
Install Virt Enviroment
virtualenv -p python3 env
source env/bin/activate
pip3 install pandas numpy xlrd xlwt xlutils datetime


These variables are hard coded 
filename = 'Work Load Forecast.xlsx'

# Excel Sheets
sheet = lambda:0
sheet.project = 'Project Log'
sheet.employee = 'Employee Times'

# Column Headers
# Lowercase
col.date = 'date'
col.project_type = 'project type'
col.planner = 'planner'
col.status = 'status'
col.employee = "employee"
col.complete = 'complete'
col.pending = 'pending'
