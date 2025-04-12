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
from datetime import datetime
import re


__author__ = "Morris Haid"
__copyright__ = "Copyright 2024"
__credits__ = ["Morris Haid"]
__license__ = "MIT License"
__version__ = "0.4.1"
__maintainer__ = "Morris Haid"
__email__ = "morris.haid@hhu.de"
__status__ = "Prototype"

# Constants for File Directory
INPUT_DIR = os.path.abspath(os.path.dirname(__file__))+"/input/" # Use directory of script
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__))+"/output/" # Use directory of script


# Setting Constants
EXCEL_SHEET_SPIKE_NAME = "P2PAmplitudes2Plot"       # Name of excel sheet containing spike information
EXCEL_SHEET_SPIKE_ADDITIONALHEADERS = 1             # Additional rows of (sub)headers
EXCEL_SHEET_SPIKE_COL_TIME = "Time"                 # Column name for time

EXCEL_SHEET_BURST_NAME = "Synopsis"                 # Name of excel sheet containing burst information
EXCEL_SHEET_BURSTS_ADDITIONALHEADERS = 0            # Additional rows of (sub)headers
EXCEL_SHEET_BURST_COL_FILENAME = "Filename"         # Column name for time
EXCEL_SHEET_BURST_COL_TIME = "Time"                 # Column name for time
EXCEL_SHEET_BURST_COL_FIELDTYPE = "FieldType"       # Column name for fieldtype
EXCEL_SHEET_BURST_FIELDTYPE_BURST = "NoB/Minute"  # FieldType of Burst per Minute


PVAL_TRESHOLD = 0.06
CONSTBASEL_TRESHOLD = 0.5
# below setting contants can be made dynamic by adding a # at the start of the line. The user will be asked for input
BASELINE_COUNT = 7
DURATION_COUNT = 5
WASHOUT_COUNT_MIN = 3
WASHOUT_COUNT_MAX = 10
# INDEX_START = 23
AUTOFILTER_CHANNELS = True



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
    excel_sheet_spike = excel_file.parse(EXCEL_SHEET_SPIKE_NAME)
    print("Fetching sheet '" + EXCEL_SHEET_SPIKE_NAME + "' succeeded.")


    # Open correct sheet
    excel_sheet_burst = excel_file.parse(EXCEL_SHEET_BURST_NAME)
    print("Fetching sheet '" + EXCEL_SHEET_BURST_NAME + "' succeeded.")

    # Fetch user settings
    settings = fetch_settings(filename, excel_sheet_spike)

    # Create result object
    results = {}

    # Create overview output
    overview_out = {
        "ch_included": 0,
        "ch_valid": 0,
        "ch_invalid": 0,
        "ch_pError": 0,
        "ch_pError": 0,
        "ch_unstable": 0,
        "ch_excluded": 0
    }

    # Calculate index of periods pre, during, post
    pre_index_start = settings["index_start"] - settings["baseline_count"]
    pre_index_end = settings["index_start"] - 1
    dur_index_start = settings["index_start"]
    dur_index_end = settings["index_end"]
    post_index_start = settings["index_end"] + 1
    post_index_end = post_index_start + settings["washout_count_max"] - 1
    if post_index_end >= len(excel_sheet_spike.index):
        post_index_end = len(excel_sheet_spike.index) - 1

    # Add dataframe for each group
    for group_name in settings["groups"]:

        # Create result object for group
        results[group_name + "_NoS_Abs"] = pd.DataFrame()
        results[group_name + "_NoS_Rel"] = pd.DataFrame()
        results[group_name + "_NoB_Abs"] = pd.DataFrame()

        # Add label in the first column
        add_label_df(results[group_name + "_NoS_Abs"], (pre_index_end - pre_index_start + 1), (dur_index_end - dur_index_start + 1), (post_index_end - post_index_start + 1), True, True, False)
        add_label_df(results[group_name + "_NoS_Rel"], (pre_index_end - pre_index_start + 1), (dur_index_end - dur_index_start + 1), (post_index_end - post_index_start + 1), False, False, True)
        add_label_df(results[group_name + "_NoB_Abs"], (pre_index_end - pre_index_start + 1), (dur_index_end - dur_index_start + 1), (post_index_end - post_index_start + 1), False, False, True)


    # Filter columns for channels
    cols_channel = [col for col in excel_sheet_spike.columns if col.startswith('Ch')]

    # Iterate trough all channels
    for col_channel in cols_channel:

        try:

            # Find the index of the column
            col_index = excel_sheet_spike.columns.get_loc(col_channel)

            # Get the columnm "2P Amplitude"
            col_P2PAmp_name = excel_sheet_spike.columns[col_index]
            # Select this column from the DataFrame
            col_P2PAmp = excel_sheet_spike[col_P2PAmp_name]
            # Create output var
            average_p2p = []

            # Check if the column header is correct
            if col_P2PAmp[0].startswith("P2PAmp"):

                # Extract spikes
                p2p_pre = extract_period(col_P2PAmp, pre_index_start, pre_index_end, EXCEL_SHEET_SPIKE_ADDITIONALHEADERS)
                p2p_dur = extract_period(col_P2PAmp, dur_index_start, dur_index_end, EXCEL_SHEET_SPIKE_ADDITIONALHEADERS)

                average_p2p_pre = calc_average(p2p_pre)
                average_p2p_dur = calc_average(p2p_dur)

                average_p2p = append_line(average_p2p, average_p2p_pre)
                average_p2p = append_line(average_p2p, average_p2p_dur)



            # Get the column before
            col_NoS_name = excel_sheet_spike.columns[col_index - 1]

            # Select this column from the DataFrame
            col_NoS = excel_sheet_spike[col_NoS_name]

            # Check if the column header is correct
            if col_NoS[0] == "NoS/Minute":


                # Extract spikes
                spikes_pre_raw = extract_period(col_NoS, pre_index_start, pre_index_end, EXCEL_SHEET_SPIKE_ADDITIONALHEADERS)
                spikes_dur_raw = extract_period(col_NoS, dur_index_start, dur_index_end, EXCEL_SHEET_SPIKE_ADDITIONALHEADERS)
                spikes_post_raw = extract_period(col_NoS, post_index_start, post_index_end, EXCEL_SHEET_SPIKE_ADDITIONALHEADERS)


                # Calc statistics for channel
                print("-----------------------------------------------------------------------")
                print("Calculating statistics for " + col_channel)
                stat = calc_statistic(spikes_pre_raw, spikes_dur_raw)


                # Check for valid return
                if stat is not None:
                
                    # Check for significance
                    if AUTOFILTER_CHANNELS is False or stat["tteqvar"]["p"] <= PVAL_TRESHOLD or stat["ttwelch"]["p"] <= PVAL_TRESHOLD or stat["manwhitu"]["p"] <= PVAL_TRESHOLD:
                        print("Baseline-Period is significant from Application-Period")


                        # Check statistic
                        overview_out["ch_included"] += 1
                        valid = True
                        if stat[stat["comp"]]["p"] > PVAL_TRESHOLD:
                            valid = False
                            # Mark channel as invalid pError
                            overview_out["ch_pError"] += 1
                        if stat["baseline_const"] == False:
                            valid = False
                            # Mark channel as invalid unstable Baseline
                            overview_out["ch_unstable"] += 1

                        if valid:
                            # Mark channel as valid
                            overview_out["ch_valid"] += 1
                        else:
                            # Mark channel as invalid
                            overview_out["ch_invalid"] += 1


                        # Calculate averages (pre/baseline and dur)
                        average_spikes_pre_raw = calc_average(spikes_pre_raw)
                        average_spikes_dur_raw = calc_average(spikes_dur_raw)
                        print("-> Calculated spike averages for baseline-/application-period: " + str(average_spikes_pre_raw) + " / " + str(average_spikes_dur_raw))


                        # Combine spikes
                        spikes_raw = combine_values([spikes_pre_raw,spikes_dur_raw,spikes_post_raw])


                        # Calc relative spikes
                        spikes_rel = calc_period_rel(spikes_raw, average_spikes_pre_raw)

                        # Calc %change baseline/application-period
                        if average_spikes_pre_raw != 0:
                            spikes_change = average_spikes_dur_raw / average_spikes_pre_raw * 100
                        else:
                            spikes_change = 0

                        # Add lines to dataframe
                        spikes_rel = append_line(spikes_rel, average_spikes_pre_raw)
                        spikes_rel = append_line(spikes_rel, average_spikes_dur_raw)
                        spikes_rel = append_line(spikes_rel, spikes_change)


                        # Add statistics
                        spikes_raw_stat = combine_values([average_p2p,[spikes_change],spikes_raw])
                        spikes_raw_stat = append_statistics(spikes_raw_stat, stat)


                        # ---- Compute Bursts per Minute
                        # Merge time and NoB into one dataframe
                        burst_df = merge_timeburst(excel_sheet_burst)

                        # Fetch burst for channel
                        burst_channel = burst_df[col_channel]
                        print("Merged Burst per Minute in dataframe (" + str(len(burst_channel)) + " entries for channel)")


                        # Extract bursts
                        bursts_pre_raw = extract_period(burst_channel, pre_index_start, pre_index_end, EXCEL_SHEET_BURSTS_ADDITIONALHEADERS)
                        bursts_dur_raw = extract_period(burst_channel, dur_index_start, dur_index_end, EXCEL_SHEET_BURSTS_ADDITIONALHEADERS)
                        bursts_post_raw = extract_period(burst_channel, post_index_start, post_index_end, EXCEL_SHEET_BURSTS_ADDITIONALHEADERS)


                        # Calculate averages (pre/baseline and dur)
                        average_bursts_pre_raw = calc_average(bursts_pre_raw)
                        average_bursts_dur_raw = calc_average(bursts_dur_raw)
                        print("-> Calculated burst averages for baseline-/application-period: " + str(average_bursts_pre_raw) + " / " + str(average_bursts_dur_raw))
                        
                        # Combine bursts
                        bursts_raw = combine_values([bursts_pre_raw,bursts_dur_raw,bursts_post_raw])

                        # Calc %change baseline/application-period
                        if average_bursts_pre_raw != 0:
                            bursts_change = average_bursts_dur_raw / average_bursts_pre_raw * 100
                        else:
                            bursts_change = 0

                        # Add lines to dataframe
                        bursts_raw = append_line(bursts_raw, average_bursts_pre_raw)
                        bursts_raw = append_line(bursts_raw, average_bursts_dur_raw)
                        bursts_raw = append_line(bursts_raw, bursts_change)



                        # Check for excitation or inhibition
                        appl_group = 0
                        if average_spikes_pre_raw < average_spikes_dur_raw:
                            print("Excitation was detected")
                            appl_group = 0

                        else:
                            print("Inhibition was detected")
                            appl_group = 1


                        # Add to dataframe
                        add_dataframe_column(results[settings["groups"][appl_group] + "_NoS_Abs"], col_channel, spikes_raw_stat)
                        add_dataframe_column(results[settings["groups"][appl_group] + "_NoS_Rel"], col_channel, spikes_rel)
                        add_dataframe_column(results[settings["groups"][appl_group] + "_NoB_Abs"], col_channel, bursts_raw)

                    else:
                        # Mark channel as excluded
                        overview_out["ch_excluded"] += 1

                        print("Baseline-Period is NOT significant from Application-Period\nChannel will be ignored")
                
                else:
                    # Mark channel as excluded
                    overview_out["ch_excluded"] += 1

                    print(spikes_pre_raw)
                    print(spikes_dur_raw)
                    raise Exception("Baseline-Period is NOT significant from Application-Period\nChannel will be ignored\nStatistics could not be calculated")
            
            else:
                raise Exception("Correct column could not be found\nChannel will be ignored")

        except Exception as e:
            print(e)
            input("Press any key to continue")



    # Create about page
    about = pd.DataFrame()
    about["Desc"] = [
        "Filename",
        "Pre-Period Lines",
        "During-Period Lines",
        "Post-Period Lines",
        "Autofilter Channels",
        "Treshold for significant P-Value",
        "Treshold for constant baseline (%change)",
        "Analysis Date",
        "Channels detected",
        "-> included ch",
        "--> valid ch",
        "--> invalid but not excluded ch",
        "---> pError",
        "---> unstable baseline",
        "-> excluded ch",
        "Tool",
        "Tool Version",
        "Tool Author",
        "Tool Licence",
    ]

    # create values
    try:
        value_ch_valid = " " + str(overview_out["ch_valid"]) + " (" + str(round(overview_out["ch_valid"] / overview_out["ch_included"] * 100,2)) + "%)"
    except:
        value_ch_valid = "N/A (Calc Error)"

    try:
        value_ch_invalid = " " + str(overview_out["ch_invalid"]) + " (" + str(round(overview_out["ch_invalid"] / overview_out["ch_included"] * 100,2)) + "%)"
    except:
        value_ch_invalid = "N/A (Calc Error)"



    about["Value"] = [
        filename,
        str(pre_index_end - pre_index_start + 1)   + " (Line# " + str(convert_indexToLine(pre_index_start))  + " to " + str(convert_indexToLine(pre_index_end))  + ")",
        str(dur_index_end - dur_index_start + 1)   + " (Line# " + str(convert_indexToLine(dur_index_start))  + " to " + str(convert_indexToLine(dur_index_end))  + ")",
        str(post_index_end - post_index_start + 1) + " (Line# " + str(convert_indexToLine(post_index_start)) + " to " + str(convert_indexToLine(post_index_end)) + ")",
        str(AUTOFILTER_CHANNELS),
        str(PVAL_TRESHOLD),
        str(CONSTBASEL_TRESHOLD),
        datetime.today().strftime('%Y-%m-%d'),
        str(len(cols_channel)),
        str(overview_out["ch_included"]),
        value_ch_valid,
        value_ch_invalid,
        "  " + str(overview_out["ch_pError"]),
        "  " + str(overview_out["ch_unstable"]),
        str(overview_out["ch_excluded"]),
        "Analyzer for SpAnner Synopsis",
        __version__,
        __author__,
        __license__
    ]
    result_complete = {"About": about}
    result_complete.update(results)


    # create a excel writer object
    with pd.ExcelWriter(OUTPUT_DIR + "ANALYSIS_" + filename) as writer:
        
        for dataframe in result_complete:
            result_complete[dataframe].to_excel(writer, sheet_name=dataframe, index=False)





def fetch_settings(filename, data):

    settings = {
        "baseline_count": 0,
        "index_start": 0,   	    # Index of row for application-start starting from 1
        "index_end": 0,             # Index of row for application-end starting from 1
        "washout_count_min": 0,
        "washout_count_max": 0,
        "groups": [
            "Excited",
            "Inhibited"
        ]
    }

    print("Fetching settings for Synopsis '" + filename + "'")

    # --- Setting for 'baseline_count'
    # Static setting
    settings["baseline_count"] = int(BASELINE_COUNT)
    print("Baseline: Average of " + str(settings["baseline_count"]) + " measurements before substance-application.")

    
    # --- Setting for 'baseline_count'
    # Static setting
    settings["washout_count_min"] = int(WASHOUT_COUNT_MIN)
    settings["washout_count_max"] = int(WASHOUT_COUNT_MAX)
    print("Washout: Min. " + str(settings["washout_count_min"]) + " / Max. " + str(settings["washout_count_max"]) + " measurements before substance-application.")


    # Fetch / Autodetect line start in name
    index_start_autodetect = fetch_startline(filename)


    # --- Setting for 'index_start'
    # Static setting
    if index_start_autodetect is not False:
        settings["index_start"] = convert_lineToIndex(int(index_start_autodetect))
        print("Substance application start set to data index " + str(settings["index_start"]) + ".")
    elif 'INDEX_START' in globals():
        settings["index_start"] = int(INDEX_START)
        print("Substance application start set to data index " + str(settings["index_start"]) + ".")
    # Dynamic setting
    else:
        data_corrected = data[EXCEL_SHEET_SPIKE_COL_TIME]
        data_corrected.index = data_corrected.index + 2
        data_corrected = pd.concat([pd.Series(["*","*"]), data_corrected])
        # Print time column
        print(data_corrected)
        print("File: " + filename)

        # Fetch index for start of substance application
        settings["index_start"] = 0
        range_min = convert_indexToLine(settings["baseline_count"] + 1)
        range_max = len(data[EXCEL_SHEET_SPIKE_COL_TIME]) - DURATION_COUNT - convert_indexToLine(settings["washout_count_min"])
        while not int(settings["index_start"]) in range(range_min, range_max + 1):
            settings["index_start"] = int(input("Data index for the start of substance application (min. " + str(range_min) + ", max. " + str(range_max) + "): "))

        settings["index_start"] = convert_lineToIndex(settings["index_start"])



    # --- Setting for 'index_end'
    # Static setting
    settings["index_end"] = settings["index_start"] + int(DURATION_COUNT) - 1
    print("Substance application end set to data index " + str(settings["index_end"]) + ".")




    # --- Do Checks for Settings
    if settings["index_start"] - settings["baseline_count"] < 0:
        settings["baseline_count"] = settings["index_start"] - 1
        print("Baseline: Corrected period length to " + str(settings["baseline_count"]) + ", this is the max. datapoints available before 'index_start'.")



    return settings


def fetch_startline(filename):
    x = re.search(r"^[L|l]ine\s?([0-9]*)[_\s]+.*$", filename)

    if x:
        print("Autodetected start of application period at line " + str(x.group(1)) + ".")
        return x.group(1)
    else:
        return False






def add_label_df(data, len_pre, len_dur, len_post, perc_change, stats, averages):

    if data.empty:
        states = []

        if perc_change:
            states.append("2P_Amp Pre")
            states.append("2P_Amp Dur")
            states.append("%change")

        for i in range(len_pre):
            states.append("Pre")

        for i in range(len_dur):
            states.append("During")

        for i in range(len_post):
            states.append("Post")

        if stats:
            states.append("Baseline Constant")

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

        if averages:
            states.append("Baseline avg")
            states.append("Application avg")
            states.append("%change")

        data["State"] = states


def extract_period(data, start, end, row_offset):

    values = []

    start = int(start) + row_offset - 1
    end = int(end) + row_offset

    for i in range(start, end):

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

        # Const Baseline
        if stat["baseline_const"] is not None:
            values.append(stat["baseline_const"])
        else:
            values.append("-")

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


def append_line(values, line):

    values.append(line)

    return values


def add_dataframe_column(df, col_name, col_values):

    df[col_name] = col_values





def calc_rowOffset():
    return 1 + EXCEL_SHEET_SPIKE_ADDITIONALHEADERS


def convert_lineToIndex(line):

    # Calculate the offset needed to compensate for all rows of header
    row_offset_header = calc_rowOffset()

    return line - row_offset_header


def convert_indexToLine(index):
    # Calculate the offset needed to compensate for all rows of header
    row_offset_header = calc_rowOffset()

    return index + row_offset_header



def calc_statistic(x,y):

    results = {
        "baseline_const": None,
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
        # Check for consant baseline
        results["baseline_const"] = calc_constantBaseline(x)
        print("Results of baseline constant check: " + str(results["baseline_const"]))

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




def calc_constantBaseline(x, treshold=CONSTBASEL_TRESHOLD):
    start = (x[0] + x[1]) / 2
    end = (x[-2] + x[-1]) / 2

    change = start / end

    if change > (1+treshold) or change < (1-treshold):
        return False
    
    return True


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





def merge_timeburst(df):
    
    time_df = df[[EXCEL_SHEET_BURST_COL_FILENAME, EXCEL_SHEET_BURST_COL_TIME]]
    nob_df = df.drop([EXCEL_SHEET_BURST_COL_FILENAME, EXCEL_SHEET_BURST_COL_TIME], axis=1)

    time = time_df[time_df[EXCEL_SHEET_BURST_COL_TIME] != '*'].reset_index(drop=True)
    nob_minute = nob_df[nob_df[EXCEL_SHEET_BURST_COL_FIELDTYPE] == EXCEL_SHEET_BURST_FIELDTYPE_BURST].reset_index(drop=True)
    
    merged_df = time.merge(nob_minute, left_index=True, right_index=True)

    return merged_df






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