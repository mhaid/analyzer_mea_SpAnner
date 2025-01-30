#!/usr/bin/env python
"""Analysis of a Synopsis output from SpAnner

Description:
This python script analyses relevant output data from all synopsis files generated
with SpAnner placed in the input directory. The user has to enter some further
information. Then the analysis takes place and an output xlsx file is saved in the
output directory.

Further information:
pandas Website  -> https://pandas.pydata.org/
scypy Website   -> https://scipy.org/


Tested with:
Python      -> Version 3.10.12
pandas      -> Version 2.2.2
openpyxl    -> Version 3.1.5
scipy       -> Version 1.12.0

Licence
MIT License

Copyright (c) 2024 Morris Haid

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import os
import pandas as pd
import scipy


__author__ = "Morris Haid"
__copyright__ = "Copyright 2024"
__credits__ = ["Morris Haid"]
__license__ = "MIT License"
__version__ = "0.1.1"
__maintainer__ = "Morris Haid"
__email__ = "morris.haid@hhu.de"
__status__ = "Prototype"

# Constants for File Directory
INPUT_DIR = os.path.abspath(os.path.dirname(__file__))+"/input/" # Use directory of script
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__))+"/output/" # Use directory of script


# Setting Constants
EXCEL_SHEET_NAME = "P2PAmplitudes2Plot"
EXCEL_SHEET_HEADERLINES = 2
PVAL_TRESHOLD = 0.06
# below setting contants can be made dynamic by adding a # at the start of the line. The user will be asked for input
BASELINE_COUNT = 7
DURATION_COUNT = 5
# INDEX_START = 23
# INDEX_END = 28 # Will only take effect if DURATION_COUNT not defined



print("-----------------------------------------------------------------------")
print("-----------------------------------------------------------------------")
print( '---{: ^65}---'.format(">>> Analyzer for SpAnner Synopsis <<<") )
print( '---{: ^65}---'.format("analyzes .xlsx ") )
print( '---{: ^65}---'.format("Version "+__version__) )
print( '---{: ^65}---'.format("by "+__author__) )
print( '---{: ^65}---'.format("Licence: "+__license__) )
print( '---{: ^65}---'.format("Status: "+__status__) )
print("-----------------------------------------------------------------------")
print("-----------------------------------------------------------------------")
print("")
print("")

# --------------- Helper Functions ---------------
def do_xlsx_conversion(filename):

    # Open Excel File
    excel_file = pd.ExcelFile(INPUT_DIR+filename)
    print("Opening excel file succeeded.")

    # Open correct sheet
    excel_sheet = excel_file.parse(EXCEL_SHEET_NAME)
    print("Fetching sheet '" + EXCEL_SHEET_NAME + "' succeeded.")

    # Fetch user settings
    settings = fetch_settings(filename, excel_sheet)

    # Create result object
    results = {}

    # Calculate index of periods pre, during, post
    pre_index_start = settings["index_start"] - settings["baseline_count"]
    pre_index_end = settings["index_start"] - 1
    dur_index_start = settings["index_start"]
    dur_index_end = settings["index_end"]
    post_index_start = settings["index_end"] + 1
    post_index_end = len(excel_sheet.index) - 1

    # Add dataframe for each group
    for group_name in settings["groups"]:

        # Create result object for group
        results[group_name + "_Raw"] = pd.DataFrame()
        results[group_name + "_Rel"] = pd.DataFrame()

        # Add label in the first column
        add_label_df(results[group_name + "_Raw"], (pre_index_end - pre_index_start + 1), (dur_index_end - dur_index_start + 1), (post_index_end - post_index_start + 1), True, False)
        add_label_df(results[group_name + "_Rel"], (pre_index_end - pre_index_start + 1), (dur_index_end - dur_index_start + 1), (post_index_end - post_index_start + 1), False, True)


    # Filter columns for channels
    cols_channel = [col for col in excel_sheet.columns if col.startswith('Ch')]

    # Iterate trough all channels
    for col_channel in cols_channel:

        # Find the index of the column
        col_index = excel_sheet.columns.get_loc(col_channel)

        # Get the column before
        col_before_name = excel_sheet.columns[col_index - 1]

        # Select this column from the DataFrame
        col_before = excel_sheet[col_before_name]

        # Check if the column header is correct
        if col_before[0] == "NoS/Minute":

            # Extract values
            values_pre_raw = extract_period(col_before, pre_index_start, pre_index_end)
            values_dur_raw = extract_period(col_before, dur_index_start, dur_index_end)
            values_post_raw = extract_period(col_before, post_index_start, post_index_end)



            # Calc statistics for channel
            print("-----------------------------------------------------------------------")
            print("Calculating statistics for " + col_channel)
            stat = calc_statistic(values_pre_raw, values_dur_raw)


            # Check for valid return
            if stat is not None:
            
                # Check for significance
                if stat["tteqvar"]["p"] <= PVAL_TRESHOLD or stat["ttwelch"]["p"] <= PVAL_TRESHOLD or stat["manwhitu"]["p"] <= PVAL_TRESHOLD:
                    print("Baseline-Period is significant from Application-Period")

                    # Calculate averages (pre/baseline and dur)
                    average_pre_raw = calc_average(values_pre_raw)
                    average_dur_raw = calc_average(values_dur_raw)
                    print("-> Calculated averages for baseline-/application-period: " + str(average_pre_raw) + " / " + str(average_dur_raw))


                    # Combine values
                    values_raw = combine_values([values_pre_raw,values_dur_raw,values_post_raw])


                    # Calc relative values
                    values_rel = calc_period_rel(values_raw, average_pre_raw)

                    # Add baseline
                    values_rel = append_baseline(values_rel, average_pre_raw)


                    # Add statistics
                    values_raw_stat = values_raw
                    values_raw_stat = append_statistics(values_raw_stat, stat)


                    # Check for excitation or inhibition
                    if average_pre_raw < average_dur_raw:
                        print("Excitation was detected")

                        # Add to dataframe
                        add_dataframe_column(results[settings["groups"][0] + "_Raw"], col_channel, values_raw_stat)
                        add_dataframe_column(results[settings["groups"][0] + "_Rel"], col_channel, values_rel)

                    else:
                        print("Inhibition was detected")

                        # Add to dataframe
                        add_dataframe_column(results[settings["groups"][1] + "_Raw"], col_channel, values_raw_stat)
                        add_dataframe_column(results[settings["groups"][1] + "_Rel"], col_channel, values_rel)



                else:
                    print("Baseline-Period is NOT significant from Application-Period")
                    print("Channel will be ignored")
            
            else:
                print("Statistics could not be calculated")
                print("Channel will be ignored")
            


    # create a excel writer object
    with pd.ExcelWriter(OUTPUT_DIR + "ANALYSIS_" + filename) as writer:
        
        for dataframe in results:
            results[dataframe].to_excel(writer, sheet_name=dataframe, index=False)




def fetch_settings(filename, data):

    settings = {
        "baseline_count": 0,
        "index_start": 0,
        "index_end": 0,
        "groups": [
            "Excited",
            "Inhibited"
        ]
    }

    print("Fetching settings for Synopsis '" + filename + "'")

    # --- Setting for 'baseline_count'
    # Static setting
    if 'BASELINE_COUNT' in globals():
        settings["baseline_count"] = int(BASELINE_COUNT)
        print("Baseline: Average of " + str(settings["baseline_count"]) + " measurements before substance-application.")
    # Dynamic setting
    else:
        settings["baseline_count"] = 0
        range_min = 1
        range_max = len(data["Time"]) - 3
        while not int(settings["baseline_count"]) in range(range_min, range_max + 1):
            settings["baseline_count"] = int(input("Number of measurements before substance application to use for calculating the baseline (min. " + str(range_min) + ", max. " + str(range_max) + "): "))
    


    # --- Setting for 'index_start'
    # Static setting
    if 'INDEX_START' in globals():
        settings["index_start"] = int(INDEX_START)
        print("Substance application start set to data index " + str(settings["index_start"]) + ".")
    # Dynamic setting
    else:
        data_corrected = data["Time"]
        data_corrected.index = data_corrected.index + 2
        data_corrected = pd.concat([pd.Series(["*","*"]), data_corrected])
        # Print time column
        print(data_corrected)

        # Fetch index for start of substance application
        settings["index_start"] = 0
        range_min = settings["baseline_count"] + 1 + EXCEL_SHEET_HEADERLINES
        range_max = len(data["Time"]) - 2 + EXCEL_SHEET_HEADERLINES
        while not int(settings["index_start"]) in range(range_min, range_max + 1):
            settings["index_start"] = int(input("Data index for the start of substance application (min. " + str(range_min) + ", max. " + str(range_max) + "): "))

        settings["index_start"] = settings["index_start"] - EXCEL_SHEET_HEADERLINES



    # --- Setting for 'index_end'
    # Static setting
    if 'DURATION_COUNT' in globals():
        settings["index_end"] = settings["index_start"] + int(DURATION_COUNT) - 1
        print("Substance application end set to data index " + str(settings["index_end"]) + ".")
    # Static setting
    elif 'INDEX_END' in globals():
        settings["index_end"] = int(INDEX_END)
        print("Substance application end set to data index " + str(settings["index_end"]) + ".")
    # Dynamic setting
    else:
        # Fetch index for end of substance application
        settings["index_end"] = 0
        range_min = settings["index_start"] + EXCEL_SHEET_HEADERLINES
        range_max = len(data["Time"]) - 2 + EXCEL_SHEET_HEADERLINES
        while not int(settings["index_end"]) in range(range_min, range_max + 1):
            settings["index_end"] = int(input("Data index for the end of substance application (min. " + str(range_min) + ", max. " + str(range_max) + "): "))

        settings["index_end"] = settings["index_end"] - EXCEL_SHEET_HEADERLINES

    return settings







def add_label_df(data, len_pre, len_dur, len_post, stats, baseline):

    if data.empty:
        states = []

        for i in range(len_pre):
            states.append("Pre")

        for i in range(len_dur):
            states.append("During")

        for i in range(len_post):
            states.append("Post")

        if stats:
            # states.append("Shapiro t-stat pre")
            # states.append("Shapiro t-stat during")
            states.append("Shapiro p-value pre")
            states.append("Shapiro p-value during")
            states.append("Normal distribution")
            
            # states.append("Levene t-stat")
            states.append("Levene p-value")
            states.append("Equal Variance")

            # states.append("TTest eqVar t-stat")
            states.append("TTest eqVar p-value")

            # states.append("TTest Welch t-stat")
            states.append("TTest Welch p-value")
            
            # states.append("Man-Whit-U t-stat")
            states.append("Man-Whit-U p-value")

            states.append("Applicable test")
            # states.append("Applicable t-stat")
            states.append("Applicable p-value")

        if baseline:
            states.append("Baseline avg")

        data["State"] = states


def extract_period(data, start, end):

    values = []

    for i in range(int(start),int(end)+1):

        # Check if entry is a valid integer
        if isinstance(data[i], (int,float)):
            # Save raw value as return value
            values.append(data[i])

        # Otherwise replace value with None
        else:
            values.append(None)


    return values


def rem_empty_values(x):
    z = []

    # Check every element for int or float value
    for y in x:
        if y is not None:
            z.append(y)

    return z


def combine_values(values):
    
    values_comb = []
    for value in values:
        values_comb.extend(value)

    return values_comb


def append_statistics(values, stat):

    if stat is not None:

        # Shapiro
        if stat["shapiro"]["x"] is not None and stat["shapiro"]["y"] is not None:
            # values.append(stat["shapiro"]["x"]["w"])
            # values.append(stat["shapiro"]["y"]["w"])
            values.append(stat["shapiro"]["x"]["p"])
            values.append(stat["shapiro"]["y"]["p"])
            values.append(stat["shapiro"]["normal"])
        else:
            # values.append("-")
            # values.append("-")
            values.append("-")
            values.append("-")
            values.append("-")

        # Levene
        if stat["levene"] is not None:
            # values.append(stat["levene"]["w"])
            values.append(stat["levene"]["p"])
            values.append(stat["levene"]["eqVar"])
        else:
            # values.append("-")
            # values.append("-")
            values.append("-")

        # T-Test eqVar
        if stat["tteqvar"] is not None:
            # values.append(stat["tteqvar"]["w"])
            values.append(stat["tteqvar"]["p"])
        else:
            # values.append("-")
            values.append("-")

        # T-Test Welch
        if stat["ttwelch"] is not None:
            # values.append(stat["ttwelch"]["w"])
            values.append(stat["ttwelch"]["p"])
        else:
            # values.append("-")
            values.append("-")

        # Mann-Whitney-U
        if stat["manwhitu"] is not None:
            # values.append(stat["manwhitu"]["w"])
            values.append(stat["manwhitu"]["p"])
        else:
            # values.append("-")
            values.append("-")

        # Comparison
        if stat["comp"] is not None:
            values.append(stat["comp"])
            # values.append(stat[stat["comp"]]["w"])
            values.append(stat[stat["comp"]]["p"])
        else:
            values.append("-")
            # values.append("-")
            values.append("-")



    return values


def append_baseline(values, baseline):

    values.append(baseline)

    return values


def add_dataframe_column(df, col_name, col_values):

    df[col_name] = col_values





def calc_statistic(x,y):

    results = {
        "shapiro": {
            "x": None,
            "y": None
        },
        "levene": None,
        "comp": None,
        "tteqvar": None,
        "ttwelch": None,
        "manwhitu": None
    }


    # Remove empty values from datasets
    x = rem_empty_values(x)
    y = rem_empty_values(y)

    try:
        # Check for normal distribution
        results["shapiro"]["x"] = stat_shapiro(x)
        results["shapiro"]["y"] = stat_shapiro(y)
        results["shapiro"]["normal"] = check_shapiro(results["shapiro"])

        # Check for difference in variance
        results["levene"] = stat_levene(x,y)
        results["levene"]["eqVar"] = check_levene(results["levene"])
        
        # Perform T-Test with equal variance
        results["tteqvar"] = stat_ttest(x,y,True)
        print("Results of independent T-Test with equal variance: " + str(results["tteqvar"]["w"]) + " / " + str(results["tteqvar"]["p"]))

        # Perform T-Test with welch correction
        results["ttwelch"] = stat_ttest(x,y,False)
        print("Results of independent T-Test with welch correction: " + str(results["ttwelch"]["w"]) + " / " + str(results["ttwelch"]["p"]))

        # Perform Mann-Whitney U
        results["manwhitu"] = stat_mannwhitneyu(x,y)
        print("Results of Mann-Whitney-U: " + str(results["manwhitu"]["w"]) + " / " + str(results["manwhitu"]["p"]))



        # The value of this statistic tends to be high (close to 1) for samples drawn from a normal distribution.
        print("Results of Shapiro-Wilk: " + str(results["shapiro"]["x"]["p"]) + " / " + str(results["shapiro"]["y"]["p"]))
        if results["shapiro"]["normal"]:

            print("-> Normal distribution")
            
            # The value of the statistic tends to be high when there is a large difference in variances.
            print("Results of Levene-Test: " + str(results["levene"]["w"]) + " / " + str(results["levene"]["p"]))
            # Variances are equal
            if results["levene"]["eqVar"]:
                print("-> Equal variance")
                results["comp"] = "tteqvar"

            # Variances are not equal
            else:
                print("-> Not equal variance")
                results["comp"] = "ttwelch"


        else:
            print("-> No normal distributionion")

            results["comp"] = "manwhitu"



        # Return statistics
        return results

    except Exception as e:
        print("An exception occurred: ", e)


    return None


def stat_shapiro(x):

    results = {
        "w": None,
        "p": None
    }

    # Check for normal distribution
    stat = scipy.stats.shapiro(x)

    results["w"] = stat.statistic
    results["p"] = stat.pvalue

    return results


def check_shapiro(stat):
    normal = None

    if stat["x"]["p"] > 0.05 and stat["y"]["p"] > 0.05:
        normal = True
    else:
        normal = False

    return normal


def stat_levene(x,y):

    results = {
        "w": None,
        "p": None
    }

    # Check for normal distribution
    stat = scipy.stats.levene(x,y)

    results["w"] = stat.statistic
    results["p"] = stat.pvalue

    return results


def check_levene(stat):
    eqVar = None

    if stat["p"] <= 0.05:
        eqVar = True
    else:
        eqVar = False

    return eqVar


def stat_ttest(x,y,equal_var):
    
    results = {
        "w": None,
        "p": None
    }

    # Check for differences between datasets
    # -> Independend
    # -> Normal distributed
    # -> Equal Variance True/False
    stat = scipy.stats.ttest_ind(x,y,equal_var=equal_var)

    results["w"] = stat.statistic
    results["p"] = stat.pvalue

    return results


def stat_mannwhitneyu(x,y):
    
    results = {
        "w": None,
        "p": None
    }

    # Check for differences between datasets
    # -> Independend
    # -> Not normal distributed
    stat = scipy.stats.mannwhitneyu(x,y)

    results["w"] = stat.statistic
    results["p"] = stat.pvalue

    return results





def calc_average(x):

    # Remove empty values from datasets
    x = rem_empty_values(x)

    # Calculate averages
    avg = sum(x)/len(x)

    return avg


def calc_period_rel(data, baseline):

    data_rel = []

    for i in range(0,len(data)):

        # Check if entry is a valid integer
        if isinstance(data[i], (int,float)):
            # Calc relative value and save as return value
            measurement_rel = data[i] / baseline * 100
            data_rel.append(measurement_rel)

        # Otherwise replace value with None
        else:
            data_rel.append(None)



    return data_rel







# --------------- STEP 0: Fetch user input ---------------
# check if neccessary directories exist

# --------------- STEP 1: Fetch available regions in space and parcellation ---------------
print("Start analysis of SpAnner Synopsis")
print("Locating files in "+INPUT_DIR)
filesconverted_cnt = 0

for filename in os.listdir( INPUT_DIR ):
    
    print("----------")

    if filename.endswith(".xlsx"):
        print("---------- " + filename + " ----------")

        do_xlsx_conversion(filename)

        print("Done converting " + filename)
        filesconverted_cnt += 1
    else:
        print("Skipping conversion of "+filename+" due to incompatible format.")


print("----------")
print("")
print("Processed all files in the directory '"+OUTPUT_DIR+"'.")
print("Analyzed "+str(filesconverted_cnt)+" files.")
print("")
print("Thanks for using.")
print("Have a good day!")
print("")
print('{: ^15}'.format("/\\_/\\") )
print('{: ^15}'.format("( o.o )") )
print('{: ^15}'.format("> ^ <") )
print("")