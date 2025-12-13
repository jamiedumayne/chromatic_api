def GET_GEMINI_INFO(gemini_function_selection, sheet_data):
    api_key = "AIzaSyCkQOckZOTzYOETPcdM6IbjhyQvqtgpTKU"

    if gemini_function_selection["active"] == gemini_function_selection["general"]:
        report_values_csv = sheet_data["report_values"].to_csv(index=False)
        report_formulas_csv = sheet_data["report_formulas"].to_csv(index=False)
        summary_values_csv = sheet_data["summary_values"].to_csv(index=False)
        summary_formulas_csv = sheet_data["summary_formula_df"].to_csv(index=False)

        csv_data = {"report_values":report_values_csv, "report_formulas": report_formulas_csv, "summary_values":summary_values_csv, "summary_formulas":summary_formulas_csv}

    return api_key, csv_data