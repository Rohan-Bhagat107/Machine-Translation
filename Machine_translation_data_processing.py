import os
import pandas as pd
import difflib  # For comparing
from langdetect import detect  # For Detecting the language
from deep_translator import GoogleTranslator  # For translation
import Levenshtein


# Function for calculating match % between our target and the google's translation
def calculate_match_percentage(target, translated_data):
    return difflib.SequenceMatcher(None, target, translated_data).ratio() * 100

'''Function for calculating edit dist between our target and the google's translation
'''
def calculate_edit_distance(source, target):
    return Levenshtein.distance(source, target)

'''Function for transalting our source data using google's translation
'''
def translate_text(text, source_lang="ja", target_lang="en"):
    if not isinstance(text, str) or text.strip() == "":
        return ""
    try:
        translation = GoogleTranslator(source=source_lang, target=target_lang).translate(text.strip())
        return translation if translation else text
    except Exception as e:
        print(f"Translation error: {e}")
        return text


'''Function for detecting the column language
'''
def detect_lang(column_data):
    new_text = " ".join(map(str, column_data.dropna()))  # Here we are converting the column data into string
    lang_counts = {"en": 0, "ja": 0}  # Dict for storing the occurance of english data and japanese data
    total_checked = 0  # How many times we are checking

    for value in new_text:
        try:
            lang = detect(str(value))
            if lang in lang_counts:
                lang_counts[lang] += 1  # Incrementing the language count if it is occured
                total_checked += 1
        except:
            continue

    if total_checked == 0:
        return "Unknown"
    '''Here we are calculating the languagewise %'''
    en_percent = (lang_counts["en"] / total_checked) * 100
    ja_percent = (lang_counts["ja"] / total_checked) * 100
    '''Returning that language which has higher %'''
    if en_percent > ja_percent:
        return "en"
    else:
        return "ja"


'''Function for processing the excel file and saving it to their desired locations'''
def process_excel(file_path, output_path, error_files):
    df = pd.read_excel(file_path, skiprows=8, header=None,
                       engine="xlrd")  # Reading the excel file by skipping initial 8 rows
    df = df.iloc[:, :2].copy()  # Creating copy of dataframe by considering initial 2 columns

    if df.empty:
        print(f"Warning: {file_path} is empty after skipping rows.")
        error_files.append(file_path)  # Appending the file if it is empty
        return
    if df.shape[1] < 2:
        print(f"Skipping {file_path} - Less than 2 valid columns available.")
        return
    '''calling a lang_detect function for language detection'''
    lang_col1 = detect_lang(df.iloc[:25, 0])
    lang_col2 = detect_lang(df.iloc[:25, 1])

    '''Naming columns based on their language'''
    if lang_col1 == "ja" and lang_col2 != "ja":
        df.columns = ["Source", "Target"]
    elif lang_col1 != "ja" and lang_col2 == "ja":
        df.columns = ["Target", "Source"]
    elif lang_col1 == lang_col2:
        lang_col1 = detect_lang(df.iloc[:50, 0])
        lang_col2 = detect_lang(df.iloc[:50, 1])
    else:
        print(f"Skipping {os.path.basename(file_path)} - Can't detect the language for this file.")
        error_files.append(os.path.basename(file_path))
        return True
    if lang_col1==lang_col2:
        print(f"We are skipping the file {file_path}....Same language detected for both columns")
        error_files.append(os.path.basename(file_path))

    else:
        match_percentages = []  # For storing match %
        edit_distances = []  # For storing edit distance
        translated_targets = []  # For storing the translated data
        print("It takes some time to process..... so wait!")
        for index, row in df.iterrows():
            source_text = str(row.get("Source", "")).strip()
            target_text = str(row.get("Target", "")).strip()
            # print(source_text,target_text)

    #We have to skip it if its not present
            if not source_text or source_text=="nan" and not target_text or target_text=="nan":
                translated_targets.append("")
                match_percentages.append("")
                edit_distances.append("")
                continue


            elif not source_text or source_text=="nan":
                translated_targets.append("")
                match_percentages.append("")
                edit_distances.append("")
                continue

            elif not target_text or target_text=="nan":
                # If target is missing, translating source and leaving the  match % & edit distance blank
                translated_value = translate_text(source_text)
                translated_targets.append(translated_value)
                match_percentages.append("")
                edit_distances.append("")
                continue
            else:
                # If both source and target exist, translating and calculating the  match & edit distance
                translated_value = translate_text(source_text)
                match_percentage = calculate_match_percentage(target_text, translated_value)
                edit_distance = calculate_edit_distance(target_text, translated_value)

                translated_targets.append(translated_value)
                match_percentages.append(round(match_percentage, 2))
                edit_distances.append(edit_distance)
        '''Here we are adding the relevant data into their containers'''
        df["Google Translation"] = translated_targets
        df["Match Percentage"] = match_percentages
        df["Edit Distance"] = edit_distances
        '''Storing the o/p file into the o/p folder'''
        output_file = os.path.join(output_path, f"processed_{os.path.basename(file_path)}")
        df.to_excel(output_file, index=False, engine="openpyxl")

        #Here we are filtering the records based on src and tgt weather they both are present or not
        filtered_df=df[df["Source"].notna() & df["Target"].notna()]

        '''Creating two folders for saving filtered data based
            on the edit distance calculated'''
        Result_for_below_20 = os.path.join(output_path, "Below_20_Edit_Distance")
        Results_for_above_20 = os.path.join(output_path, "Above_20_Edit_Distance")
        os.makedirs(Result_for_below_20, exist_ok=True)
        os.makedirs(Results_for_above_20, exist_ok=True)

        #  Filtering the records based on the edit Distance as above 20 and below 20
        below_20_df = filtered_df[filtered_df["Edit Distance"] <= 20]
        above_20_df = filtered_df[filtered_df["Edit Distance"] > 20]

        below_20_df = below_20_df[["Source", "Target", "Edit Distance"]]
        above_20_df = above_20_df[["Source", "Target", "Edit Distance"]]

            #---------------------------------------------------------------------------------------------------
        '''saving the filtered data '''
        below_20_excel = os.path.join(Result_for_below_20, f"Below_20_{os.path.basename(file_path)}")
        above_20_excel = os.path.join(Results_for_above_20, f"Above_20_{os.path.basename(file_path)}")

        below_20_df.to_excel(below_20_excel, index=False, engine="openpyxl")
        above_20_df.to_excel(above_20_excel, index=False, engine="openpyxl")
        print(f"Processing completed. Results saved to {output_file}")

'''Function for checking weather the input folder exists or not'''
def path_validator(input_path):
    return os.path.exists(input_path)

try:
    input_path = input("Please enter your folder path here: ")
    output_path = input("Please enter your folder path for output here: ")
    if not os.path.isdir(input_path):
        raise Exception("Please provide a valid folder.")
except Exception as e:
    print(e)
else:
    excel_file_count, error_files = 0, []
    try:
        if path_validator(input_path):
            print("Processing started...Please wait!")
            for n, each_file in enumerate(os.listdir(input_path)):
                if each_file.endswith(".xls"):
                    print("Working file---->", each_file)
                    excel_file_count += 1
                    file_path = os.path.join(input_path, each_file)
                    same_language=process_excel(file_path, output_path, error_files)
                    if not same_language:
                        print(" File is processed successfully")

            if excel_file_count == 0:
                print("No Excel files found in the provided folder.")
            else:
                print("Task Completed! Files processed successfully.")
                if error_files:
                    print("Files that encountered errors:", error_files)
        else:
            print("Provided path doesn't exist. Please provide a valid path.")
    except Exception as e:
        print("Error during processing:", e)
