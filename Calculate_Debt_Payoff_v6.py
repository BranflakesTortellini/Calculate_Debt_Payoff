import pandas as pd
from tkinter import Tk, simpledialog, filedialog, messagebox

def calculate_debt_payoff(initial_balance, apr, payment_multiples, base_min_payment_rate):
    monthly_interest_rate = apr / 12
    histories = {}
    results = pd.DataFrame()
    
    for multiple in payment_multiples:
        histories[multiple] = [initial_balance]

    month = 0
    all_zero = False

    while not all_zero:
        month += 1
        current_all_zero = True
        print(f'Calculating month: {month}')  # Debug print

        for multiple in payment_multiples:
            if len(histories[multiple]) <= month:
                last_balance = histories[multiple][-1]
                if last_balance > 0:
                    interest = last_balance * monthly_interest_rate
                    min_payment = max(last_balance * base_min_payment_rate * multiple, 50)
                    new_balance = max(last_balance + interest - min_payment, 0)
                    histories[multiple].append(new_balance)
                    if new_balance > 0 and multiple >= 1:
                        current_all_zero = False
                else:
                    histories[multiple].append(0)
        
        if month >= 1000:  # Failsafe to prevent infinite looping
            print("Reached month limit without zeroing all balances.")
            break

        print(f'Status at month {month}: {current_all_zero}')  # Debug print
        if current_all_zero:
            print('All balances are zero, breaking loop.')
            break

    max_length = max(len(hist) for hist in histories.values())
    results['Months'] = range(max_length)
    for multiple in payment_multiples:
        results[f'Balance_{multiple}x'] = histories[multiple] + [0] * (max_length - len(histories[multiple]))
    results['APR'] = [apr] * max_length
    results['Base Min Payment Rate'] = [base_min_payment_rate] * max_length
    return results

# Setup for GUI
root = Tk()
root.withdraw()

initial_balance = float(simpledialog.askstring("Input", "Enter the initial balance:"))
apr = float(simpledialog.askstring("Input", "Enter the APR (e.g., 0.20 for 20%):"))
base_min_payment_rate = float(simpledialog.askstring("Input", "Enter the base minimum payment rate (e.g., 0.03 for 3%):"))
payment_multiples_input = simpledialog.askstring("Input", "Enter payment multiples (e.g., 0.5, 1, 2) separated by commas:")
payment_multiples = [float(x.strip()) for x in payment_multiples_input.split(',')]

results = calculate_debt_payoff(initial_balance, apr, payment_multiples, base_min_payment_rate)

# File saving dialogue
default_filename = f"Debt_Payoff_{initial_balance}_apr{apr}_min_{base_min_payment_rate}.xlsx"
filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], initialfile=default_filename, title="Save the file as")

if filename:
    results.to_excel(filename, index=False)
    messagebox.showinfo("Success", f"Data saved successfully in {filename}")
else:
    messagebox.showerror("Error", "Save file operation cancelled.")

root.mainloop()
