import pandas as pd
import matplotlib.pyplot as plt

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image

# Read Excel File
df = pd.read_excel("data/sales.xlsx")

# Total Sales
total_sales = df["Sales"].sum()

# Department Sales
dept_sales = df.groupby("Department")["Sales"].sum()

# Employee Sales
emp_sales = df.groupby("Employee")["Sales"].sum()

# Top Employee
top_employee = emp_sales.idxmax()

# Additional KPIs
total_departments = df["Department"].nunique()
total_employees = df["Employee"].nunique()

# Summary DataFrame
summary = pd.DataFrame({
    "Metric": ["Total Sales", "Top Employee"],
    "Value": [total_sales, top_employee]
})

print(summary)

# Create Bar Chart
plt.figure(figsize=(6, 4))
dept_sales.plot(kind="bar", color="skyblue")
plt.title("Department Sales")
plt.xlabel("Department")
plt.ylabel("Sales")
plt.tight_layout()
plt.savefig("charts/department_sales.png")
plt.close()

# Create Pie Chart
plt.figure(figsize=(5, 5))
dept_sales.plot(
    kind="pie",
    autopct="%1.1f%%"
)
plt.ylabel("")
plt.title("Department Contribution")
plt.savefig("charts/department_pie.png")
plt.close()

# Generate Excel Report
with pd.ExcelWriter(
    "reports/sales_report.xlsx",
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Raw Data",
        index=False
    )

    dept_sales.to_excel(
        writer,
        sheet_name="Department Sales"
    )

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    pd.DataFrame().to_excel(
        writer,
        sheet_name="Dashboard",
        index=False
    )

print("Report Generated Successfully!")

# Open Workbook
wb = load_workbook("reports/sales_report.xlsx")

dashboard_ws = wb["Dashboard"]

# Dashboard Title
dashboard_ws["D2"] = "SALES DASHBOARD"

# KPI Cards
dashboard_ws["D4"] = "Total Sales"
dashboard_ws["E4"] = total_sales

dashboard_ws["D6"] = "Top Employee"
dashboard_ws["E6"] = top_employee

dashboard_ws["D8"] = "Departments"
dashboard_ws["E8"] = total_departments

dashboard_ws["D10"] = "Employees"
dashboard_ws["E10"] = total_employees

# Formatting
dashboard_ws["D2"].font = Font(
    size=18,
    bold=True,
    color="FFFFFF"
)

dashboard_ws["D2"].fill = PatternFill(
    start_color="4F81BD",
    end_color="4F81BD",
    fill_type="solid"
)

dashboard_ws["D4"].font = Font(bold=True)
dashboard_ws["D6"].font = Font(bold=True)
dashboard_ws["D8"].font = Font(bold=True)
dashboard_ws["D10"].font = Font(bold=True)

# Insert Charts
bar_img = Image("charts/department_sales.png")
dashboard_ws.add_image(bar_img, "D14")

pie_img = Image("charts/department_pie.png")
dashboard_ws.add_image(pie_img, "M14")

# Save Workbook
wb.save("reports/sales_report.xlsx")

print("Dashboard Created Successfully!")