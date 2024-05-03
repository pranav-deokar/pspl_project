import openpyxl
import matplotlib.pyplot as plt
import os

def add_expenses():
    print("User Expenses Input:")
    # Asking for budget of each expense category
    grocery_budget = int(input('Enter budget for grocery: '))
    clothing_budget = int(input('Enter budget for clothing: '))
    travelling_budget = int(input('Enter budget for travelling: '))
    
    # Using tuples to represent individual financial transactions
    # Each tuple contains amount spent on each category
    grocery_expense = int(input('Enter actual expense on grocery: '))
    clothing_expense = int(input('Enter actual expense on clothing: '))
    travelling_expense = int(input('Enter actual expense on travelling: '))
    
    other_expenses = int(input("Enter your other expenses: "))
    
    # Asking for budget goal
    goal = int(input("Enter your overall budget goal: "))
    
    # Storing all expenses, budgets, and the overall goal in a list
    return [[grocery_budget, grocery_expense], [clothing_budget, clothing_expense], [travelling_budget, travelling_expense], other_expenses, goal]

def calculate_savings(expenses, income):
    grocery_savings = expenses[0][0] - expenses[0][1]
    clothing_savings = expenses[1][0] - expenses[1][1]
    travelling_savings = expenses[2][0] - expenses[2][1]
    other_savings = income - expenses[3] - sum([expense[1] for expense in expenses[:3]])  # Calculate savings for 'Other' category
    total_savings = grocery_savings + clothing_savings + travelling_savings + other_savings
    return grocery_savings, clothing_savings, travelling_savings, other_savings, total_savings

def generate_report(user_name, income, expenses, total_savings, month, filename):
    print('\nUser Summary:')
    print(f'User: {user_name}')  # Displaying user's name
    
    print('Expense Category\tBudget\tActual Expense\tSavings')

    categories = ['Grocery', 'Clothing', 'Travelling', 'Other']
    filtered_expenses = [expense for expense in expenses if isinstance(expense, list)]  # Filter out non-lists
    filtered_labels = [category for expense, category in zip(filtered_expenses, categories) if isinstance(expense, list)]  # Filter labels

    for i, category in enumerate(categories):
        if isinstance(expenses[i], list):
            budget = expenses[i][0]
            expense = expenses[i][1]
            savings = budget - expense
        else:
            budget = 'N/A'
            expense = expenses[i]
            savings = 'N/A'
        print(f'{category}\t\t{budget}\t{expense}\t{savings}')

    print(f'\nTotal Savings: {total_savings}')  # Print the total savings calculated in calculate_savings function
    if total_savings >= expenses[4]:
        print('Congratulations! You have achieved your overall budget goal.')
    else:
        print('You have not achieved your overall budget goal.')

    # Write data to Excel
    if os.path.isfile(filename):
        # Load existing workbook
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.active
    else:
        # Create new workbook
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = 'Expense Report'
        # Add headers to the worksheet
        headers = ['User', 'Total Savings', 'Goal Achievement', 'Month', 'Grocery Actual Expense', 'Clothing Actual Expense', 'Travelling Actual Expense', 'Other Actual Expense']
        for col, header in enumerate(headers, start=1):
            worksheet.cell(row=1, column=col, value=header)

    row = worksheet.max_row + 1
    worksheet.cell(row=row, column=1, value=user_name)  # Adding user's name to Excel
    worksheet.cell(row=row, column=2, value=total_savings)
    worksheet.cell(row=row, column=3, value='Achieved' if total_savings >= expenses[4] else 'Not Achieved')
    worksheet.cell(row=row, column=4, value=month)  # Adding month to Excel
    
    # Adding actual expenses to Excel
    for i, category in enumerate(categories):
        if isinstance(expenses[i], list):
            worksheet.cell(row=row, column=i+5, value=expenses[i][1])
        else:
            worksheet.cell(row=row, column=i+5, value=expenses[i])

    workbook.save(filename)
    print(f'\nExpense report saved to {filename}')
    plt.figure(figsize=(11, 8))  # Increase the figure size to accommodate both subplots effectively

    # First subplot for the bar chart
    plt.subplot(1, 2, 1)
    expenses = [expense[1] for expense in filtered_expenses] + [expenses[categories.index('Other')]]
    plt.bar(categories, expenses)
    plt.title('Expense Comparison')
    plt.xlabel('Expense Category')
    plt.ylabel('Expense Amount')

    # Second subplot for the pie chart
    plt.subplot(1, 2, 2) 
    if filtered_expenses:  # Check if filtered_expenses is not empty
        plt.pie([expense[1] for expense in filtered_expenses], labels=filtered_labels, autopct='%1.1f%%', startangle=140)
        plt.title('Expense Distribution')
    else:
        plt.text(0.5, 0.5, 'No data available', horizontalalignment='center', verticalalignment='center')

    plt.tight_layout()  # Adjust subplot parameters to give specified padding.
    plt.show()


# Main 
print("User 1:")

# Input user's name
user_name = input('Enter your name: ')

# Input number of income sources and their names
num_sources = int(input('Enter the number of your income sources: '))
sources = [input(f'Enter income source {i+1}: ') for i in range(num_sources)]

# Input total income
income = int(input("Enter your total income through all sources for 1 month: "))

# Input month
month = input("Enter the month: ")

# Define the filename
filename = "C:\\Users\\prana\\OneDrive\\Dokumen\\PROJECT_FINAL.xlsx"

# Function to add expenses
expenses = add_expenses()

# Function to calculate savings
grocery_savings, clothing_savings, travelling_savings, other_savings, total_savings = calculate_savings(expenses, income)

# Function to generate report
generate_report(user_name, income, expenses, total_savings, month, filename)
