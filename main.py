"""
This is V2 of the Buying Target Group Finder, to return the buying Target Group and not create the Curve Data.
"""

import pandas as pd
import xlwings as xw
from target_group_list import target_group_list, target_group_CPP_dict

core_target_group = input("Enter the target group: ")
reach_goal_min = input("Enter the Min. Level reach goal: 0.0 - 1.0: ")
reach_goal_mid = input("Enter the Mid. Level reach goal: 0.0 - 1.0: ")
reach_goal_max = input("Enter the Max. Level reach goal: 0.0 - 1.0: ")

reach_goals = [reach_goal_min, reach_goal_mid, reach_goal_max]

# open the Excel file
wb = xw.Book("160922_GrossContactsCalculatorTVFY2223.xlsm")
ws = wb.sheets["Manual TV"]  # ws is the worksheet object and the second Worksheet (Manual TV) in the Workbook

# select the core target group
ws.range("B6").value = core_target_group
ws.range("B3").value = "Germany"
ws.range("D3").value = "Germany"

# Set all fields to 0 in row 10
month_columns = ["H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"]
for month in month_columns:
    ws.range(month + "10").value = 0

df = pd.DataFrame(columns=["core_target_group",
                           "buying_target_group",
                           "GRP",
                           "Reach Goal",
                           "Reach",
                           "Budget",
                           "CostPerReachPoint"])

buying_target_groups = []

# Loop through all buying target groups to calculate 3 reach points
for buying_target_group in target_group_list[:2]:
    ws.range("C10").value = buying_target_group
    cost_per_grp = target_group_CPP_dict[buying_target_group]

    # Increase GRP until 3 reach points are reached
    for reach_goal in reach_goals:
        for grp in range(10, 800, 10):
            ws.range("H10").value = grp
            current_reach = ws.range("H15").value
            if ws.range("F10").value == 0:  # Add a skip on Target Groups without Conversion Factor
                break
            elif current_reach >= reach_goal:
                target_group_performance = {"core_target_group": core_target_group,
                                            "buying_target_group": buying_target_group,
                                            "GRP": grp,
                                            "Reach Goal": reach_goal,
                                            "Reach": current_reach,
                                            "Budget": grp * cost_per_grp,
                                            "CostPerReachPoint": (grp * cost_per_grp) / current_reach}
                buying_target_groups.append(target_group_performance)
                break

df = df.append(buying_target_groups, ignore_index=True)
pd.set_option('display.max_columns', None)
print(df)

# Find the best buying target group per reach goal
best_buying_target_groups = df.groupby("Reach Goal")
best_buying_target_groups = best_buying_target_groups.apply(lambda x: x.sort_values("CostPerReachPoint").iloc[0])

# Print the results
for index, row in best_buying_target_groups.iterrows():
    print(f"Best buying target group for {row['Reach Goal']} reach goal is {row['buying_target_group']} "
          f"with {row['GRP']} GRP, {row['Reach']} reach,"
          f" {row['CostPerReachPoint']} Cost per Reach at "
          f"{row['Budget']} budget")
