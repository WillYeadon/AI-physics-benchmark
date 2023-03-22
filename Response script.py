import openai
import openpyxl
from openpyxl import load_workbook

# Set up the OpenAI API
openai.api_key = "Use your own!"

# Load the Excel workbook
workbook = load_workbook("Question-Answer-Bank.xlsx")
worksheet = workbook.active

# Define the prompt prefix
prompt_prefix = "Answer the following question as correctly and concisely as possible ideally using numbers alone including units where appropriate: '"

prompt_suffix = "'? Please answer below."

# Iterate through the rows, starting from row 2 (skipping the header)
for row in range(2, worksheet.max_row + 1):
    question = worksheet.cell(row=row, column=6).value  # Column 6 is the 'Question' column

    if question is not None:
        prompt = prompt_prefix + question + prompt_suffix

        # Call the OpenAI API with the prompt
        response = openai.Completion.create(
            engine="text-davinci-002",
            prompt=prompt,
            max_tokens=50,
            n=1,
            stop=None,
            temperature=0.5,
        )

        # Extract the answer from the API response
        answer = response.choices[0].text.strip()

        # Store the answer in the 'Response' column (column 8)
        worksheet.cell(row=row, column=8).value = answer

        # Rinse and repeat for the different models

# Save the updated workbook
workbook.save("Question-Answer-Bank-Updated.xlsx")
