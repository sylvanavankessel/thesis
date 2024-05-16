# -*- coding: utf-8 -*-
"""
This file executes the optimalisation for the complex (continuous) model.
It gathers the input of the Excel file and creates the variables and formulas for the model

@author: Sylvana
"""
import sys
import os

# Get the parent directory
parent_dir = os.path.dirname(os.path.realpath(__file__))

# Add the parent directory to sys.path
sys.path.append(parent_dir)

import matplotlib.pyplot as plt
import numpy as np
from scipy.interpolate import pchip_interpolate
import xlwings as xw
#from openpyxl.utils.cell import column_index_from_string, get_column_letter, coordinate_from_string
import pandas as pd
from decimal import Decimal

from genetic_algorithm_pfm import GeneticAlgorithm
from genetic_algorithm_pfm.tetra_pfm import TetraSolver
from Preferendus_model_functions import *
from Preferendus_model_settings import settings


@xw.func
def black_box_run_model():
    """
    Function to run the continuous model optimalisation through Excel
    """ 
        
    """
    Step 1: Define Excel worksheets input for the model, these are defined in the settings file
    """ 
    (preference_curves_ws, preferendus_tool_ws, multiple_runs_tool_ws, folder_location, number_of_runs, multiple_runs_print, black_box_output_optimum, 
            num_stakeholders, num_objectives, num_house_variables, num_categorical_variables, continuous_num_objectives, categorical_num_objectives, weights_list, 
            stakeholder_weights_list, continuous_objective_list, categorical_objective_list, functions_list, functions_list_new, functions_null_scenario_balance_list, 
            scenario_delta_balance_list, categorical_balance_sum_list, surface_list, black_box_bounds_list, options, start_plot) = settings()        
    """
    Step 2: We have to put the objective and a constraint function itself in this part,
    else you get an error with loading the variables during optimalisation.
    """
    def objective(variables):
       """      
       Objective function that is fed to the GA. Calls the separate preference functions that are declared above.

       :param variables: array with design variable values per member of the population. Can be split by using array
       slicing
       :return: 1D-array with aggregated preference scores for the members of the population.
       """
       # extract 1D design variable arrays from full 'variables' array
       x1 = variables[:, 0]
       x2 = variables[:, 1]
       x3 = variables[:, 2]
       x4 = variables[:, 3]
       x5 = variables[:, 4]

       # calculate the preference scores
       active_p = []
       categorical_scores = []
       
       i=0
       for s in range(num_stakeholders):
           for o in range(num_objectives):
               if (o+1)%5!=0:
                   x1_categorical = variables[:, o+5]
                   x2_categorical = variables[:, o+9]
                   x3_categorical = variables[:, o+13]
                   x4_categorical = variables[:, o+17]
                   x5_categorical = variables[:, o+21]
                   categorical_score, formula = objective_formula_categorical_differentiate(surface_list, continuous_objective_list[i], functions_list_new[o], 
                                                                                  x1, x2, x3, x4, x5, 
                                                                                  x1_categorical, x2_categorical, x3_categorical, x4_categorical, x5_categorical)
                   active_p.append(formula)
                   categorical_scores.append(categorical_score)
                   i+=1
               if (o+1)%5==0:
                   formula = continuous_objective_formula_balance(i, continuous_objective_list, categorical_scores, functions_null_scenario_balance_list, scenario_delta_balance_list, x1, x2, x3, x4, x5)
                   active_p.append(formula)
                   categorical_scores = []
                   i+=1            
       # aggregate preference scores and return this to the GA
       # score_tetra = TetraSolver().request(weights_list,active_p)
       # print(score_tetra)
       return weights_list, active_p
    
    def constraint_space(variables):
        """Constraint that checks if the sum of the build area is not higher than the maximum space.

        :param variables: ndarray of n-by-m, with n the population size of the GA and m the number of variables.
        :return: list with scores of the constraint
        """
        x1 = variables[:, 0] 
        x2 = variables[:, 1]
        x3 = variables[:, 2]
        x4 = variables[:, 3]
        x5 = variables[:, 4]
        
        return (x1*surface_list[0])+(x2*surface_list[1])+(x3*surface_list[2])+(x4*surface_list[3])+(x5*surface_list[4]) - surface_list[5] # < 0
    
    cons = [['ineq', constraint_max], ['ineq', constraint_min], ['ineq', constraint_affordable], ['ineq', constraint_space]]

    """
    Step 3: Now we have everything for the optimization, we can run it. For more information about the different options to 
    configure the GA, see the docstring of GeneticAlgorithm (via help()) or chapter 4 of the reader.
    """
    # run the GA and print its result
    print('Run GA with IMAP')
    ga = GeneticAlgorithm(objective=objective, constraints=cons, bounds=black_box_bounds_list, options=options)
    score_IMAP, design_variables_IMAP, _ = ga.run()
    print(f'Optimal result for x1 = {round(design_variables_IMAP[0], 0)},  x2 = {round(design_variables_IMAP[1], 0)}, '
          f'x3 = {round(design_variables_IMAP[2], 0)}, x4 = {round(design_variables_IMAP[3], 0)}, '
          f'x5 = {round(design_variables_IMAP[4], 0)}, (with sum = {round(sum(design_variables_IMAP[0:5]))} houses)')
  
   
    """
   Step 4: Now we have the results, we want to return them to Excel. First preference curves are returned as pictures. Then, final outcomes are returned.
   """
    # Create preference curve plots for in Excel
    linspaces = create_linspaces(continuous_objective_list)
    preference_arrays = create_preference_arrays(continuous_objective_list, linspaces)
    
    # Some settings for creating the plots
    # First some lists to be filled
    formula_values = []
    formula_list = []
    individual_preference_scores = []
    # Define the horizontal and vertical offsets between plots
    horizontal_offset = 3  # Number of columns between plots
    vertical_offset = 19   # Number of rows between plots
    # Lastly a counter
    i=0
    # Looping and creating all plots
    for s in range(num_stakeholders):
        for o in range(num_objectives):
            if (o+1)%5!=0:
                formula_value = create_formula_value_categorical(o, functions_list_new[o], design_variables_IMAP)
                formula_values.append(formula_value)
                formula_list.append(formula_value)
                individual_preference_score = create_individual_preference_score(continuous_objective_list[i], formula_value)
                individual_preference_scores.append(individual_preference_score)
                create_save_categorical_plot(folder_location, linspaces[i], preference_arrays[i], formula_value, individual_preference_score, horizontal_offset, vertical_offset, o, s, i, preference_curves_ws, start_plot)
                i+=1
            if (o+1)%5==0:
                formula_value = create_formula_value_saldo(i, continuous_objective_list, scenario_delta_balance_list, formula_list, functions_null_scenario_balance_list, design_variables_IMAP)
                formula_values.append(formula_value)
                individual_preference_score = create_individual_preference_score(continuous_objective_list[i], formula_value)
                individual_preference_scores.append(individual_preference_score)
                create_save_saldo_plot(folder_location, linspaces[i], preference_arrays[i], formula_value, individual_preference_score, horizontal_offset, vertical_offset, o, s, i, preference_curves_ws, start_plot)
                formula_list = []
                i+=1 
    
    # Fill in optimum values
    categorical_slicer = num_categorical_variables*5
    design_variables_IMAP[-categorical_slicer:] = transform_list(list(map(float, design_variables_IMAP[-num_categorical_variables:])))
    preferendus_tool_ws.range(black_box_output_optimum).options(transpose=True).value = design_variables_IMAP
    
    return design_variables_IMAP

def black_box_run_model_multiple_runs():
    """
    Function to run the continuous model optimalisation through Excel
    """ 
        
    """
    Step 1: Define Excel worksheets input for the model, these are defined in the settings file
    """ 
    (preference_curves_ws, preferendus_tool_ws, multiple_runs_tool_ws, folder_location, number_of_runs, multiple_runs_print, black_box_output_optimum, 
            num_stakeholders, num_objectives, num_house_variables, num_categorical_variables, continuous_num_objectives, categorical_num_objectives, weights_list, 
            stakeholder_weights_list, continuous_objective_list, categorical_objective_list, functions_list, functions_list_new, functions_null_scenario_balance_list, 
            scenario_delta_balance_list, categorical_balance_sum_list, surface_list, black_box_bounds_list, options, start_plot) = settings()         
    """
    Step 2: We have to put the objective and a constraint function itself in this part,
    else you get an error with loading the variables during optimalisation.
    """
    def objective(variables):
       """      
       Objective function that is fed to the GA. Calls the separate preference functions that are declared above.

       :param variables: array with design variable values per member of the population. Can be split by using array
       slicing
       :return: 1D-array with aggregated preference scores for the members of the population.
       """
       # extract 1D design variable arrays from full 'variables' array
       x1 = variables[:, 0]
       x2 = variables[:, 1]
       x3 = variables[:, 2]
       x4 = variables[:, 3]
       x5 = variables[:, 4]

       # calculate the preference scores
       active_p = []
       categorical_scores = []
       
       i=0
       for s in range(num_stakeholders):
           for o in range(num_objectives):
               if (o+1)%5!=0:
                   x1_categorical = variables[:, o+5]
                   x2_categorical = variables[:, o+9]
                   x3_categorical = variables[:, o+13]
                   x4_categorical = variables[:, o+17]
                   x5_categorical = variables[:, o+21]
                   categorical_score, formula = objective_formula_categorical_differentiate(surface_list, continuous_objective_list[i], functions_list_new[o], 
                                                                                  x1, x2, x3, x4, x5, 
                                                                                  x1_categorical, x2_categorical, x3_categorical, x4_categorical, x5_categorical)
                   active_p.append(formula)
                   categorical_scores.append(categorical_score)
                   i+=1
               if (o+1)%5==0:
                   formula = continuous_objective_formula_balance(i, continuous_objective_list, categorical_scores, functions_null_scenario_balance_list, scenario_delta_balance_list, x1, x2, x3, x4, x5)
                   active_p.append(formula)
                   categorical_scores = []
                   i+=1            
       # aggregate preference scores and return this to the GA
       # score_tetra = TetraSolver().request(weights_list,active_p)
       # print(score_tetra)
       return weights_list, active_p
    
    def constraint_space(variables):
        """Constraint that checks if the sum of the build area is not higher than the maximum space.

        :param variables: ndarray of n-by-m, with n the population size of the GA and m the number of variables.
        :return: list with scores of the constraint
        """
        x1 = variables[:, 0] 
        x2 = variables[:, 1]
        x3 = variables[:, 2]
        x4 = variables[:, 3]
        x5 = variables[:, 4]
        
        return (x1*surface_list[0])+(x2*surface_list[1])+(x3*surface_list[2])+(x4*surface_list[3])+(x5*surface_list[4]) - surface_list[5] # < 0
    
    cons = [['ineq', constraint_max], ['ineq', constraint_min], ['ineq', constraint_affordable], ['ineq', constraint_space]]

    """
    Step 3: Now we have everything for the optimization, we can run it. For more information about the different options to 
    configure the GA, see the docstring of GeneticAlgorithm (via help()) or chapter 4 of the reader.
    """
    # run the GA and print its result
    print('Run GA with IMAP')
    ga = GeneticAlgorithm(objective=objective, constraints=cons, bounds=black_box_bounds_list, options=options)
    score_IMAP, design_variables_IMAP, _ = ga.run()
    
    # Fill in optimum values
    # design_variables_IMAP[-num_categorical_variables:] = transform_list(list(map(float, design_variables_IMAP[-num_categorical_variables:])))
    
    return design_variables_IMAP

@xw.func
def multiple_runs():
    (preference_curves_ws, preferendus_tool_ws, multiple_runs_tool_ws, folder_location, number_of_runs, multiple_runs_print, black_box_output_optimum, 
            num_stakeholders, num_objectives, num_house_variables, num_categorical_variables, continuous_num_objectives, categorical_num_objectives, weights_list, 
            stakeholder_weights_list, continuous_objective_list, categorical_objective_list, functions_list, functions_list_new, functions_null_scenario_balance_list, 
            scenario_delta_balance_list, categorical_balance_sum_list, surface_list, black_box_bounds_list, options, start_plot) = settings()  
    column_length = ['x'+str(i) for i in range (1,(num_house_variables+num_house_variables*num_categorical_variables)+1)]
    df = pd.DataFrame(columns=column_length)
    print(df)
    for i in range(1, number_of_runs+1):
        design_variables_run = black_box_run_model_multiple_runs()
        print(design_variables_run)
        df.loc[i] = design_variables_run
        print(df)
    
    multiple_runs_tool_ws.range(multiple_runs_print).value = df
    return df
        
# Function that compares all alternatives
@xw.func
def tetra_function(weight,preferences):
    score_tetra = TetraSolver().request(weight,preferences)
    return score_tetra

# Function that interpolates with the preferences, values and solutions of the objective formulas
@xw.func
def interpolate_excel(xi,yi,x):
    try:
        pchip_interpolate(xi,yi,x)
    except:
        print('Invalid input')
    return pchip_interpolate(xi,yi,x)

if __name__ == "__main__":
    #design_variables_IMAP = white_box_run_model()
    design_variables_IMAP = black_box_run_model()
    #df_multiple_runs = multiple_runs()


