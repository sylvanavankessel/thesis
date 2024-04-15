# -*- coding: utf-8 -*-
"""
Created on Sun Feb 25 14:07:36 2024

@author: Sylvana
"""
from Preferendus_model_functions_2 import *

def settings():
    """
    Function to run the continuous model optimalisation through Excel
    """ 
        
    """
    Step 1: Define Excel worksheets input for the model
    """ 
    wb = xw.Book('Preferendus_model_2.xlsm')
    wks = xw.sheets
    # Name the worksheets, the first sheet 'Input' is inputws, etc. wks[0] is called explanation and doesn't need a python link
    preferences_input_ws = wks[1]
    preference_curves_ws = wks[2]
    preferendus_tool_ws = wks[3]
    multiple_runs_tool_ws = wks[4]
    assistance_tab_ws = wks[5]

    # Give the starting cell of elements in 'preferences input' tab (the first cell)
    start_objective = "C19" # First (numerical) value of preference tables (no headers)
    start_weights = "G11" # First (numerical) value of Weights table (no headers or x1)
    
    # Give the starting cell of elements in 'preferences curves' tab (the first cell)
    start_plot = "A4" # First (left and top) cell to plot the first function

    # Give the starting cell of elements in 'preferendus tool' tab (the first cell)
    white_box_output_optimum = "G9" # First (numerical) value to print the variable outcomes (no headers)
    black_box_output_optimum = "H9" # First (numerical) value to print the variable outcomes (no headers)

    # Give the starting cell of elements in 'preferendus tool' tab (the first cell)
    number_of_runs = "I49" # Cell value that reports the number of runs to be performed
    multiple_runs_print = "A11" # First (numerical) value to print the variable outcomes (no headers)
    
    # Give the starting cell of elements in the 'assitance tab' tab (the first cell)
    start_bounds = "D6" # First (numerical) value of variable bounds table (no headers or x1)
    start_surface = "B35" # First (numerical) value of surface values (no headers or x1)
    start_functions_categorical = "C474" # First (numerical) value of variable functions table for the categorical objectives (no headers or x1)
    start_functions_categorical_new = "C474" # First (numerical) value of variable functions table for the categorical objectives (no headers or x1)
    start_functions_balance = "D260" # First (numerical) value of variable functions table for the saldo objective (no headers or x1)
    start_categorical_balance_sum = "C376" # First (numerical) value of sum of the categoric preferendus calculation table for the optimalisation (no headers or x1)
    start_stakeholder_weights = "B456" # First (numerical) value of stakeholder weights table (no headers or x1)
    
    # Define the number of stakeholders and objectives per stakeholder
    num_stakeholders = 4
    num_objectives = 5
    num_house_variables = 5 # Only take house types variables, not the categorical 
    num_categorical_variables = 4
    continuous_num_objectives = num_stakeholders * num_objectives
    categorical_num_objectives = num_stakeholders
    
    # Give the name of the folder for the preference curves
    folder_location = 'preference_curves'

    """
    Step 2: Setup for the model
    """
    # Set the number of runs for multiple runs
    number_of_runs = int(preferendus_tool_ws.range(number_of_runs).value)
    ### Make lists of all weights for all different stakeholders and put them in a list, and filter weights for active objectives ###
    s1_weights = weights_values(preferences_input_ws, start_weights, 0)
    s2_weights = weights_values(preferences_input_ws, start_weights, 42)
    s3_weights = weights_values(preferences_input_ws, start_weights, 84)
    s4_weights = weights_values(preferences_input_ws, start_weights, 126)

    # Make a flat list of all weights from all stakeholders, and one of the stakeholder weights
    weights_list = s1_weights + s2_weights + s3_weights + s4_weights
    stakeholder_weights_list = weights_values(assistance_tab_ws, start_stakeholder_weights, 0)
    
    ### Make lists of all objectives for all different stakeholders and put them in a list ###
    # Initialize an empty list to store all objective lists, 
    # one with all elements for the continuous run, and one with only the saldo for the categorical
    continuous_objective_list = []
    categorical_objective_list = []
    # Loop through stakeholders and objectives 
    for s in range(num_stakeholders):
        for o in range(num_objectives):
            # Calculate the start index (the offset) for each objective
            # 42 is offset between stakeholders, 6 is offset between objectives
            offset = s*42 + o*6
            # Get objective values and append to the continuous objective list
            obj = objective_values(preferences_input_ws, start_objective, offset)
            continuous_objective_list.append(obj)
            # Add the saldo objective values to the categorical objective list 
            if o == num_objectives-1:
                categorical_objective_list.append(obj)
                    
    ### Make lists of all functions of all objectives ###
    # Make one list of all categorical functions (4 times) and then insert an empty list (needed for looping in objective function)
    vars_ = var_values(assistance_tab_ws, start_functions_categorical)
    functions_list = vars_ #[item for _ in range(4) for item in var_values(assistance_tab_ws, start_functions_categorical)]
    for i in range(1,5):
        functions_list.insert(i * 4 + (i - 1), [])
    functions_list_new = var_values_new(assistance_tab_ws, start_functions_categorical_new)
    # Make one list of the balance values in the null scenario
    functions_null_scenario_balance_list = assistance_tab_ws.range(start_functions_balance).expand('right').options(numbers=float).value
    functions_null_scenario_balance_list = [float(val) for val in functions_null_scenario_balance_list]
    # Make a list of all functions of the delta balances for the scenarios of each objective
    functions_delta_balance_list = assistance_tab_ws.range(start_functions_balance).offset(row_offset=1).expand().value
    functions_delta_balance_list = [[float(val) for val in function] for function in functions_delta_balance_list]
    functions_delta_balance_list = [list(row) for row in zip(*functions_delta_balance_list)]
    scenario_delta_balance_list = [functions_delta_balance_list[i][j:j+4] for i in range(len(functions_delta_balance_list)) for j in range(0, len(functions_delta_balance_list[i]), 4)]
    categorical_balance_sum_list = assistance_tab_ws.range(start_categorical_balance_sum).expand('right').options(numbers=float).value

    ### Set other lists ###
    # Set surface list
    surface_list = assistance_tab_ws.range(start_surface).expand('down').value
    # Set the bounds of all variables
    black_box_bounds_list = assistance_tab_ws.range(start_bounds).expand().value
    white_box_bounds_list = black_box_bounds_list[:num_house_variables]

    # make dictionary with parameter settings for the GA run with the IMAP solver, see chapter 4 for more information
    options = {
            'n_bits': 8,
            'n_iter': 400,
            'n_pop': 500,
            'r_cross': 0.8,
            'max_stall': 8,
            'aggregation': 'tetra',
            'var_type': 'int'
            }
    return (preference_curves_ws, preferendus_tool_ws, multiple_runs_tool_ws, folder_location, number_of_runs, multiple_runs_print, white_box_output_optimum, black_box_output_optimum, 
            num_stakeholders, num_objectives, num_house_variables, num_categorical_variables, continuous_num_objectives, categorical_num_objectives, weights_list, 
            stakeholder_weights_list, continuous_objective_list, categorical_objective_list, functions_list, functions_list_new, functions_null_scenario_balance_list, 
            scenario_delta_balance_list, categorical_balance_sum_list, surface_list,white_box_bounds_list, black_box_bounds_list, options, start_plot)

if __name__ == "__main__":
    (preference_curves_ws, preferendus_tool_ws, multiple_runs_tool_ws, folder_location, number_of_runs, multiple_runs_print, white_box_output_optimum, black_box_output_optimum, 
     num_stakeholders, num_objectives, num_house_variables, num_categorical_variables, continuous_num_objectives, categorical_num_objectives, weights_list, 
     stakeholder_weights_list, continuous_objective_list, categorical_objective_list, functions_list, functions_list_new, functions_null_scenario_balance_list, 
     scenario_delta_balance_list, categorical_balance_sum_list, surface_list,white_box_bounds_list, black_box_bounds_list, options, start_plot) = settings()