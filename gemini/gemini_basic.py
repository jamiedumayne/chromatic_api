from google import genai
from google.genai import types

def ask_gemini(prompt, api_info, model="gemini-2.5-pro"):

    # fix api key bug, and module import
    client = genai.Client(api_key=api_key)
    file_path = "C:/Users/jamie.dumayne/Documents/python/repos/maintenance_agent/API_folder/gemini/prompts/general_prompt.txt"

    with open(file_path, 'r') as instruction_for_gemini_doc:
        instruction_for_gemini = instruction_for_gemini_doc.read()

    response = client.models.generate_content(
        model=model,
        config=types.GenerateContentConfig(
            system_instruction=instruction_for_gemini),
        contents = prompt
    )

    answer = response.text

    return answer