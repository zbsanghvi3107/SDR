####################################################
#                  Revision: 1.1                   #
#              Updated on: 12/03/2015              #
#                                                  #
# Whats new:                                       #
#           Fixed  Calculation of HOST lines       #
#           excluded. SASMap data is grabed        #
#           between from first SASMap string       #
#           and first CNTRLR string.               #
#                                                  #
####################################################

####################################################
#                  Revision: 1.0                   #
#              Updated on: 11/06/2015              #
####################################################

####################################################
#                                                  #
#   This file checks Read & Write errors, extracts #
#   data for F2 & F3 Menu, creates 2 DPSLD files.  #
#   Which are useful to generate report.           #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################
import os
import pandas
import csv
import numpy as np


###################################
#  Importing from other Directory
###################################
os.chdir('..')
c_path = os.getcwd()
import sys
sys.path.insert(0, r''+str(c_path)+'/IO Stress')
import log_functions

class Extract_F2_F3_SDR:

    def generate_f2_f3_sdr(csv_file, log_file):

        #################################################
        #   This step will check Read and Write Errors
        #   in .csv file. It will also raise an error_flag
        #   if it is present.
        #################################################

        lf=log_functions.Log_Functions

        [write_sum, read_sum] = lf.check_errors(csv_file)

        
        #########################
        #
        # For Iterations flag
        #
        #########################
        iter_flag = lf.find_iterations_SDR(log_file)
        
  
        ####################################################
        #   User Decision is by default is Yes, it will only 
        #   change if there is an error. 
        ####################################################
        

        # if error_flag> 0 & iter_flag >= 100:
            # print('.csv file has errors!')
            # usr_dec = input ('Do you want to continue(Y/N) ? :')
            
        # elif error_flag== 0 & iter_flag < 100:
            # print('Number of Iterations are less than 100!')
            # usr_dec = input ('Do you want to continue(Y/N) ? :')
                
        # elif error_flag> 0 & iter_flag < 100:
            # print('.csv file has errors! \nNumber of Iterations are less than 100!')
            # usr_dec = input ('Do you want to continue(Y/N) ? :')        

        # if error_flag==0:
            # if str(usr_dec)=='Y':
                            
        #########################
        #
        # For Hardware Info
        #
        #########################
        hw_info='--- Current Hardware Information ---'
        int_info= 'Internal RAIDIO SN'
        
        hw_index = lf.find_index(log_file,hw_info) 
        # Finding index where this string
        
        int_index = lf.find_index(log_file,int_info)
        
        hw_data_list = log_file[hw_index[0]:int_index[0]+1]
        
        #########################
        #
        # For Software Info
        #
        #########################
        sw_info='--- Current Software Configuration ---'
        
        sw_index = lf.find_index(log_file, sw_info) 
        # Finding index where this string
            
        cntrlr_info='CNTRLR'
        cntrlr_index = lf.find_index(log_file, cntrlr_info)
            
        search_string_h='HOST'
        host_list = []
        #host_list = lf.str_lister(search_string_h,
        #                            sw_index,log_file)
        
        search_string_s='SASMap'
        sasmap_index = lf.find_index(log_file, search_string_s)
        #try:
        #    sasmap_list = lf.str_lister(search_string_s,
        #                            sw_index,log_file)
        #except:
        log_file = np.array(log_file)
        sasmap_list = log_file[sasmap_index[0] : cntrlr_index[0]]
        
            # sasmap_list = []
            # for i in range(len(sasmap_list_temp)):
                # if search_string_h not in str(sasmap_list_temp[i]):
                    # sasmap_list.append(sasmap_list_temp[i])
                  
        return [write_sum, read_sum, iter_flag, hw_data_list, host_list,
                sasmap_list]

        
#####################################
#              END                  #
#####################################