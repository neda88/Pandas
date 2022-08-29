import pandas as pd
import os
folder_path = r"C:\Desktop\DriveTimeResults"


#change directory to where excel files are
os.chdir(os.path.join(os.getcwd(),folder_path))

# drop all these columns from each dataframe
df_drop= ['OriginID','DestinationID','DestinationRank','Total_Miles','Total_Kilometers','Total_TimeAt1KPH','Total_WalkTime','Total_TruckMinutes','Total_TruckTravelTime','Shape_Length']

# create an empty list to store dataframes (df_tmp) from each Excel file
dataframes = []
# os.listdir(folder_path) provides a list of all files in the given folder (folder_path)
for f in os.listdir(folder_path):
    # create a temporary dataframe (df_tmp) from each file in folder_path
    df_tmp = pd.read_excel(f)
    # for each df_tmp delete all columns in df_drop list above
    for item in df_drop:
        del df_tmp[item]
    # convert the Total_Minutes column in df_tmp to hours
    df_tmp['Total_Minutes'] = df_tmp['Total_Minutes']/60.00
    # convert the Total_TravelTime column in df_tmp to hours
    df_tmp['Total_TravelTime'] = df_tmp['Total_TravelTime']/60.00
    # create a list of words from the current filename (f): The 3rd values in the list is day and time togather as a string
    run_day_time = (f.split("_")[3])  #[3]: "day time"
    # rename
    #   Total_Minutes to Total_Hours_run_day_time
    #   Total_TravelTime to Total_TravelTime(h)_run_day_time
    # and overwrite the df_tmp with the new column names
    df_tmp = df_tmp.rename({'Total_Minutes': str('Total_Hours' + '_' + run_day_time), 'Total_TravelTime': str("Total_TravelTime(h)" + "_" + run_day_time)}, axis=1)
    # append the current df_tmp after all changes above to the dataframes list
    dataframes.append(df_tmp)

# create one big dataframe from all df_tmps and call it df. axis=1 creates it Horizontally
df = pd.concat(dataframes, axis=1)

# write the resulting dataframe to the path and desired filename in Excel format
df.to_excel(r"C:\Documents\points_driveTime_results.xlsx")
    
#change durectory back to where you were
os.chdir("..")
