import os
from pathlib import Path
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt

from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4


# ==========================================
# 1. FILE SETTINGS
# ==========================================

BASE_DIR = Path(__file__).resolve().parent

EXP_FILE = BASE_DIR / "expenses.xlsx"
PDF_FILE = BASE_DIR / "Expense_Report.pdf"
CHART_FILE = BASE_DIR / "Expense_Chart.png"

BUDGET = 10000

COLUMNS = [
    "ID", "Date", "Category", "Description", "Amount"
]


# ==========================================
# 2. CREATE / LOAD EXCEL FILE
# ==========================================

def initialize_file():

    if not EXP_FILE.exists():
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXP_FILE, index=False)
        print("Excel File Created Successfully!")

    else:
        df = pd.read_excel(EXP_FILE)

        for column in COLUMNS:
            if column not in df.columns:
                df[column] = None

        df = df[COLUMNS]
        df.to_excel(EXP_FILE, index=False)


def load_data():

    df = pd.read_excel(EXP_FILE)

    for column in COLUMNS:
        if column not in df.columns:
            df[column] = None

    df = df[COLUMNS]

    df["Amount"] = pd.to_numeric(
        df["Amount"], errors="coerce"
    ).fillna(0)

    return df


def save_data(df):
    df.to_excel(EXP_FILE, index=False)


# ==========================================
# 3. INPUT VALIDATION
# ==========================================

def get_valid_date():

    while True:
        date = input("Enter Date (dd-mm-yyyy): ").strip()

        try:
            datetime.strptime(date, "%d-%m-%Y")
            return date

        except ValueError:
            print("Invalid Date! Example: 03-11-2026")


def get_valid_amount():

    while True:
        try:
            amount = float(input("Enter Amount: ₹"))

            if amount <= 0:
                print("Amount must be greater than zero.")
                continue

            return amount

        except ValueError:
            print("Enter a valid amount.")


def get_valid_id():

    while True:
        try:
            return int(input("Enter Expense ID: "))

        except ValueError:
            print("Enter a valid ID.")


# ==========================================
# 4. ADD EXPENSE
# ==========================================

def add_expense():

    df = load_data()

    if df.empty or df["ID"].dropna().empty:
        expense_id = 1
    else:
        expense_id = int(
            pd.to_numeric(df["ID"], errors="coerce").max()
        ) + 1

    date = get_valid_date()
    category = input("Enter Category: ").strip()
    description = input("Enter Description: ").strip()
    amount = get_valid_amount()

    if not category:
        print("Category cannot be empty.")
        return

    new_expense = pd.DataFrame([{
        "ID": expense_id,
        "Date": date,
        "Category": category,
        "Description": description,
        "Amount": amount
    }])

    df = pd.concat([df, new_expense], ignore_index=True)
    save_data(df)

    print("\nExpense Added Successfully!")
    print("Expense ID:", expense_id)


# ==========================================
# 5. VIEW ALL EXPENSES
# ==========================================

def view_expense():

    df = load_data()

    if df.empty:
        print("No Expenses Found.")
        return

    print("\n========== ALL EXPENSES ==========")
    print(df.to_string(index=False))


# ==========================================
# 6. UPDATE EXPENSE
# ==========================================

def update_expense():

    df = load_data()

    if df.empty:
        print("No Expenses Available.")
        return

    view_expense()

    expense_id = get_valid_id()

    if expense_id not in df["ID"].values:
        print("Expense ID Not Found.")
        return

    index = df.index[df["ID"] == expense_id][0]

    print("\nEnter New Details")

    date = get_valid_date()
    category = input("New Category: ").strip()
    description = input("New Description: ").strip()
    amount = get_valid_amount()

    if not category:
        print("Category cannot be empty.")
        return

    df.loc[index, "Date"] = date
    df.loc[index, "Category"] = category
    df.loc[index, "Description"] = description
    df.loc[index, "Amount"] = amount

    save_data(df)

    print("Expense Updated Successfully!")


# ==========================================
# 7. DELETE EXPENSE
# ==========================================

def delete_expense():

    df = load_data()

    if df.empty:
        print("No Expenses Available.")
        return

    view_expense()

    expense_id = get_valid_id()

    if expense_id not in df["ID"].values:
        print("Expense ID Not Found.")
        return

    confirm = input("Confirm Delete? (yes/no): ")

    if confirm.lower() == "yes":

        df = df[df["ID"] != expense_id]
        save_data(df)

        print("Expense Deleted Successfully!")

    else:
        print("Delete Cancelled.")


# ==========================================
# 8. SEARCH EXPENSE
# ==========================================

def search_expense():

    df = load_data()

    if df.empty:
        print("No Expenses Available.")
        return

    keyword = input(
        "Enter Category or Description to Search: "
    ).strip()

    result = df[
        df["Category"].astype(str).str.contains(
            keyword, case=False, na=False
        )
        |
        df["Description"].astype(str).str.contains(
            keyword, case=False, na=False
        )
    ]

    if result.empty:
        print("No Matching Records Found.")

    else:
        print("\nSearch Results:")
        print(result.to_string(index=False))


# ==========================================
# 9. LAST 15 DAYS EXPENSES
# ==========================================

def last_15_days():

    df = load_data()

    if df.empty:
        print("No Data Found.")
        return

    df["Date"] = pd.to_datetime(
        df["Date"], dayfirst=True, errors="coerce"
    )

    today = pd.Timestamp.today().normalize()
    start_date = today - pd.Timedelta(days=14)

    result = df[
        (df["Date"] >= start_date) &
        (df["Date"] <= today)
    ]

    if result.empty:
        print("No Expenses in the Last 15 Days.")

    else:
        print(result.to_string(index=False))

        print(
            "\nLast 15 Days Total: ₹",
            round(result["Amount"].sum(), 2)
        )


# ==========================================
# 10. TOTAL EXPENSE
# ==========================================

def total_expense():

    df = load_data()

    total = df["Amount"].sum()

    print("\nTotal Expense = ₹", round(total, 2))


# ==========================================
# 11. EXPENSE DASHBOARD
# ==========================================

def pie_chart():

    df = load_data()

    if df.empty:
        print("No Expenses Found.")
        return

    report = df.groupby("Category")["Amount"].sum()

    if report.sum() <= 0:
        print("No Valid Expense Amounts.")
        return

    plt.figure(figsize=(8, 8))

    plt.pie(
        report.values,
        labels=report.index,
        autopct="%1.1f%%",
        startangle=90
    )

    plt.title("Category-wise Expense Distribution")
    plt.tight_layout()

    plt.savefig(CHART_FILE)
    plt.show()

    print("Dashboard Chart Saved Successfully!")


# ==========================================
# 12. PDF REPORT
# ==========================================

def pdf_report():

    df = load_data()

    if df.empty:
        print("No Expenses Found.")
        return

    styles = getSampleStyleSheet()

    pdf = SimpleDocTemplate(
        str(PDF_FILE),
        pagesize=A4
    )

    elements = []

    elements.append(
        Paragraph("EXPENSE TRACKER REPORT", styles["Title"])
    )

    elements.append(Spacer(1, 20))

    total = df["Amount"].sum()

    elements.append(
        Paragraph(
            f"Total Expense: Rs. {total:.2f}",
            styles["Heading2"]
        )
    )

    elements.append(Spacer(1, 15))

    data = [COLUMNS]

    for _, row in df.iterrows():

        data.append([
            str(row["ID"]),
            str(row["Date"]),
            str(row["Category"]),
            str(row["Description"]),
            f'{float(row["Amount"]):.2f}'
        ])

    table = Table(data, repeatRows=1)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))

    elements.append(table)
    pdf.build(elements)

    print("\nPDF Report Created Successfully!")
    print("File:", PDF_FILE.name)


# ==========================================
# 13. TOP SPENDING CATEGORY
# ==========================================

def top_category():

    df = load_data()

    if df.empty:
        print("No Expenses Found.")
        return

    result = df.groupby("Category")["Amount"].sum()

    top = result.idxmax()
    amount = result.max()

    print("\nHighest Spending Category")
    print("Category:", top)
    print("Amount: ₹", round(amount, 2))


# ==========================================
# 14. RECENT TRANSACTIONS
# ==========================================

def recent_transactions():

    df = load_data()

    if df.empty:
        print("No Transactions Found.")
        return

    print("\nRecent 5 Transactions:")
    print(df.tail(5).to_string(index=False))


# ==========================================
# 15. MONTHLY BUDGET ALERT
# ==========================================

def budget_alert():

    df = load_data()

    if df.empty:
        print("No Expenses Found.")
        return

    df["Date"] = pd.to_datetime(
        df["Date"], dayfirst=True, errors="coerce"
    )

    today = pd.Timestamp.today()

    current_month = df[
        (df["Date"].dt.month == today.month) &
        (df["Date"].dt.year == today.year)
    ]

    total = current_month["Amount"].sum()

    print("\nMonthly Budget Limit: ₹", BUDGET)
    print("Current Month Expense: ₹", round(total, 2))
    print("Remaining Budget: ₹", round(BUDGET - total, 2))

    if total > BUDGET:
        print("WARNING: Monthly Budget Exceeded!")

    else:
        print("Budget Under Control!")


# ==========================================
# 16. MONTHLY SUMMARY
# ==========================================

def monthly_summary():

    df = load_data()

    if df.empty:
        print("No Data Found.")
        return

    df["Date"] = pd.to_datetime(
        df["Date"], dayfirst=True, errors="coerce"
    )

    df = df.dropna(subset=["Date"])

    if df.empty:
        print("No Valid Dates Found.")
        return

    df["Month"] = df["Date"].dt.strftime("%B %Y")

    report = df.groupby("Month")["Amount"].sum()

    print("\n========== MONTHLY SUMMARY ==========")
    print(report.to_string())

    print("\nOverall Total: ₹", round(report.sum(), 2))


# ==========================================
# 17. OPEN EXCEL - VIEW SAVED DATA
# ==========================================

def open_excel():

    try:
        os.startfile(str(EXP_FILE))
        print("\nOpening Excel File...")

    except Exception as e:
        print("Unable to open Excel:", e)


# ==========================================
# 18. MAIN MENU
# ==========================================

def main():

    initialize_file()

    print("\n================================")
    print("        EXPENSE TRACKER")
    print("================================")

    while True:

        print("\n1.  Add Expense")
        print("2.  View All Expenses")
        print("3.  Update Expense")
        print("4.  Delete Expense")
        print("5.  Search Expense")
        print("6.  Last 15 Days Expenses")
        print("7.  Total Expense")
        print("8.  Expense Dashboard")
        print("9.  Generate PDF Report")
        print("10. Top Spending Category")
        print("11. Recent Transactions")
        print("12. Monthly Budget Alert")
        print("13. Monthly Summary")
        print("14. Open Excel - View Saved Data")
        print("15. Exit")

        choice = input("\nEnter Choice: ").strip()

        try:

            if choice == "1":
                add_expense()

            elif choice == "2":
                view_expense()

            elif choice == "3":
                update_expense()

            elif choice == "4":
                delete_expense()

            elif choice == "5":
                search_expense()

            elif choice == "6":
                last_15_days()

            elif choice == "7":
                total_expense()

            elif choice == "8":
                pie_chart()

            elif choice == "9":
                pdf_report()

            elif choice == "10":
                top_category()

            elif choice == "11":
                recent_transactions()

            elif choice == "12":
                budget_alert()

            elif choice == "13":
                monthly_summary()

            elif choice == "14":
                open_excel()

            elif choice == "15":
                print("\nTHANK YOU FOR USING EXPENSE TRACKER!")
                break

            else:
                print("Invalid Choice! Enter a number from 1 to 15.")

        except Exception as e:
            print("\nSomething went wrong:", e)


# ==========================================
# 19. RUN PROJECT
# ==========================================

if __name__ == "__main__":
    main()