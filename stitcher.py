# -*- coding: utf-8 -*-
"""
This script combines together all Excel files in a directory into one CSV output file.

Initially created on Tue Jun  8 14:23:00 2021 by @author: uh-sheesh.
"""

import glob
import pandas as pd
import os
import sys
import getopt
from pathlib import Path
from datetime import datetime

def usage():
    print('To use this program, please follow this example: \'python stitcher.py output\', where output refers to the desired filename. This file will be created in an \'output\' folder in the current directory.')

def divider():
    print('-------------------------')

def printer(df, sheet_name):
    #add the location of the filename for the sheet
    df['Filename'] = sheet_name
    print(sheet_name + ' inlcuded.')

def main(output_name):
    try:

        #checks if a outname has been entered
        if len(output_name) < 1:
            usage()

        #initialize the dataframe
        df = pd.DataFrame()
        
        
        #save item to a folder        
        cwd = os.getcwd()    
        
        #saves all the files in the current directoy as a list
        files = os.listdir(cwd) 
        
        #prints a divder
        divider()
        
        #verify that each item ends in the appropriate format, read the item and then append it to the dataframe 
        for item in files:
            
            #create (or recreate) a temporary dataframe to store the new sheet
            temp_df = pd.DataFrame(columns=['Filename'])
            
            if item.endswith('.xlsx'):
                temp_df = temp_df.append(pd.read_excel(item, index_col=[0]), ignore_index=True)
                printer(temp_df, item)
                df = df.append(temp_df)
                
            elif item.endswith('.csv'):
                temp_df = temp_df.append(pd.read_csv(item, index_col=[0]), ignore_index=True)
                printer(temp_df, item)
                df = df.append(temp_df)  
            
        # get current date and time for the name of the output file
        now = datetime.now()
        current_date = now.strftime("%m-%d-%y")
        current_time = now.strftime('%I-%M-%S')
        
        #format the output file's name
        output_file = output_name + '_' + current_date + '_' + current_time + '.csv'

        #format the output file directory
        output_dir = Path(cwd + '/output')
        
        #check if the folder exists, if not create the output folder
        output_dir.mkdir(parents=True, exist_ok=True)
        
        #save the dataframe to a csv file in the folder, remove the index
        df.to_csv(output_dir / output_file, index=False)  # can join path elements with / operator
        
        #print success message
        divider()
        print('Combine successful!', len(files), 'files combined.')
        
    except getopt.GetoptError:
        usage()

if __name__ == '__main__':
    try:
        main(sys.argv[1])
    except IndexError:
        usage()
    except KeyboardInterrupt:
        print('\nProgram Interrupted!')
        try:
            sys.exit(0)
        except SystemExit:
            os._exit(0)