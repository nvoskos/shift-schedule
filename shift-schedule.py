import pandas as pd
import numpy as np

def auto_generate_shift_schedule(weeks_range=(45, 52), employees=None):
    if employees is None:
        employees = [
            'PAPANASTASIOU', 'PROTOPAPA', 'PALLA', 'ARISTEIDOU', 'NIKOLAOU',
            'LEONIDOU', 'SPIROU', 'ACHILLEOS', 'AGGELI', 'NEOPHITOU',
            'CONSTANTINOU', 'ARGYROU', 'PANAYI', 'IRAKLEIDIS', 'LOUKA',
            'NARKISSOS', 'METTOU', 'CHRISTOFOROU', 'LEFTERI', 'THEODOROU',
            'CHRISTOFI', 'VOSKOS', 'KAROLIDES', 'MISIRELIDIS'
        ]
    
    days = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY']
    all_schedules = pd.DataFrame()

    # Iterate through each week to generate the schedule
    for week_num in range(weeks_range[0], weeks_range[1] + 1):
        # Create a DataFrame for the weekly schedule
        schedule = pd.DataFrame(index=employees, columns=days)

        # Assign 'SL' for IRAKLEIDIS for the entire week
        schedule.loc['IRAKLEIDIS'] = 'SL'

        # Randomly assign two days off for each employee during the weekdays
        np.random.seed(week_num)  # To ensure different output for each week
        for employee in employees:
            if employee != 'IRAKLEIDIS':
                days_off = np.random.choice(days[:5], size=2, replace=False)
                schedule.loc[employee, days_off] = 'Day Off'

        # Assign weekends off alternatively for each employee (every third employee has the weekend off)
        for i, employee in enumerate(employees):
            if employee != 'IRAKLEIDIS':
                if i % 3 == 0:  # Every third employee has the weekend off
                    schedule.loc[employee, ['SATURDAY', 'SUNDAY']] = 'Day Off'

        # Fill in remaining shifts with 'Shift' for workdays
        for employee in employees:
            for day in days:
                if pd.isna(schedule.loc[employee, day]):
                    schedule.loc[employee, day] = 'Shift'

        # Append the generated schedule for this week to the all_schedules DataFrame
        week_label = f"Week {week_num}"
        schedule.columns = pd.MultiIndex.from_product([[week_label], schedule.columns])
        all_schedules = pd.concat([all_schedules, schedule], axis=1)

    return all_schedules

# Example usage
schedule = auto_generate_shift_schedule()
# Save to Excel
schedule.to_excel("Weekly_Shift_Schedule_Weeks_45_to_52.xlsx")
print("Schedule generated and saved as 'Weekly_Shift_Schedule_Weeks_45_to_52.xlsx'")
