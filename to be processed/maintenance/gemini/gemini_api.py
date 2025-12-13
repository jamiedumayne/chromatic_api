from google import genai
from google.genai import types
from API.gemini.gemini_extra_info import GET_GEMINI_INFO

def ASK_GEMINI(question, sheet_data, gemini_function_selection):
    gemini_functions = {"update_cell": "UPDATE_CELL_VALUE", "archive_file":"ARCHIVE_FILE", "add_summary":"ADD_SUMMARY"}
    if gemini_function_selection["active"] == gemini_function_selection["general"]:
        answer = ASK_GEMINI_GENERAL(question, sheet_data, gemini_function_selection)

    return answer, gemini_functions

def ASK_GEMINI_GENERAL(question, sheet_data, gemini_function_selection):
    api_key, csv_data = GET_GEMINI_INFO(gemini_function_selection, sheet_data)

    full_prompt = "Here is a request from a user:\n " \
        f"{question} \n\n" \
        "There are 2 parts to the data a report dataset and a summary dataset. The summary shows an overall summary of the report tab, and is calculated using formulas." \
        "The following is the report dataset in CSV format\n\n" \
        f"{csv_data["report_values"]} \n\n " \
        "The following is the formulas for the report dataset, in CSV format" \
        f"{csv_data["report_formulas"]} \n\n " \
        "The following is the summary dataset in CSV format\n\n" \
        f"{csv_data["summary_values"]} \n\n " \
        "The following is the formulas for the summary dataset, in CSV format" \
        f"{csv_data["summary_formulas"]}\n\n "

    client = genai.Client(api_key=api_key)
    file_path = "C:/Users/jamie.dumayne/Documents/python/repos/maintenance_agent/API_folder/gemini/prompts/general_prompt.txt"

    with open(file_path, 'r') as instruction_for_gemini_doc:
        instruction_for_gemini = instruction_for_gemini_doc.read()

    response = client.models.generate_content(
        model="gemini-2.5-pro",
        config=types.GenerateContentConfig(
            system_instruction=instruction_for_gemini),
        contents = full_prompt
    )

    answer = response.text

    return answer


def ASK_GEMINI_SUMMARY(question, sheet_data, headings_for_gemini):
    api_key = "AIzaSyCkQOckZOTzYOETPcdM6IbjhyQvqtgpTKU"

    report_values_csv = sheet_data["report_values"].to_csv(index=False)
    report_formulas_csv = sheet_data["report_formulas"].to_csv(index=False)
    summary_values_csv = sheet_data["summary_values"].to_csv(index=False)
    summary_formulas_csv = sheet_data["summary_formula_df"].to_csv(index=False)
    headings_string = ','.join(headings_for_gemini)

    full_prompt = "Here is a request from a user:\n " \
        f"{question} \n\n" \
        "There are 2 parts to the data a report dataset and a summary dataset. The summary shows an overall summary of the report tab, and is calculated using formulas." \
        "The following is the report dataset in CSV format\n\n" \
        f"{report_values_csv} \n\n " \
        "The following is the formulas for the report dataset, in CSV format" \
        f"{report_formulas_csv} \n\n " \
        "The following is the summary dataset in CSV format\n\n" \
        f"{summary_values_csv} \n\n " \
        "The following is the formulas for the summary dataset, in CSV format" \
        f"{summary_formulas_csv}\n\n " \
        f"The columns in the summary will appear in the following order: {headings_string} "
        

    client = genai.Client(api_key=api_key)
    file_path = "C:/Users/jamie.dumayne/Documents/python/repos/maintenance_agent/API_folder/gemini/prompts/summary_prompt.txt"

    with open(file_path, 'r') as instruction_for_gemini_doc:
        instruction_for_gemini = instruction_for_gemini_doc.read()

    response = client.models.generate_content(
        model="gemini-2.5-pro",
        config=types.GenerateContentConfig(
            system_instruction=instruction_for_gemini),
        contents = full_prompt
    )

    answer = response.text

    return answer


