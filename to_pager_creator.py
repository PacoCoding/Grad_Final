import openai
import streamlit as st
from docx import Document
import pandas as pd
import to_pager_functions as fc

# Streamlit App Title
st.title("Custom Assistant: PDF Document Analyzer")

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.write(
    """
    Upload a PDF file to be analyzed by the custom assistants. The assistants will process
    the content of the PDF and provide responses based on predefined prompts. Ensure your 
    OpenAI API key is set in Streamlit Secrets.
    """
)

# Access API Key from Streamlit Secrets
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Sidebar File Upload
st.sidebar.header("Upload PDF")
uploaded_file = st.sidebar.file_uploader("Upload a PDF File", type=["pdf"])

if not uploaded_file:
    st.warning("Please upload a file to begin processing.")
    st.stop()

# Provide Feedback About the Uploaded File
st.sidebar.success(f"Uploaded File: {uploaded_file.name}")

# Preloaded Files for Prompts and Document Template
xlsx_file = "prompt_db.xlsx"
docx_file = "to_pager_template.docx"

try:
    prompt_db = pd.ExcelFile(xlsx_file)
    doc_copy = Document(docx_file)
except Exception as e:
    st.error(f"Error loading preloaded files: {e}")
    st.stop()

# Define the sections and their respective assistants
sections = {
    "A. BUSINESS OPPORTUNITY AND GROUP OVERVIEW": {
        "sheet_names": ["BO_Prompts", "BO_Format_add"],
        "assistant_id": "asst_ZwYHPxoqquAdDHmVyZrr8SgC",
    },
    "B. REFERENCE MARKET": {
        "sheet_names": ["RM_Prompts", "RM_Format_add"],
        "assistant_id": "asst_vy2MqKVgrmjCecSTRgg0y6oO",
    },
}

# Process Sections
st.subheader("Document Generation in Progress")
temp_responses = []
answers_dict = {}

for section_name, section_data in sections.items():
    st.write(f"Processing section: {section_name}")
    sheet_names = section_data["sheet_names"]
    assistant_id = section_data["assistant_id"]

    try:
        # Retrieve prompts and formatting requirements
        prompt_list, additional_formatting_requirements = fc.prompts_retriever(
            xlsx_file, sheet_names
        )
    except Exception as e:
        st.error(f"Error processing prompts for {section_name}: {e}")
        continue

    for prompt_name, prompt_message in prompt_list:
        st.write(f"Processing prompt: {prompt_name}")

        try:
            # Send the uploaded PDF directly to the assistant for analysis
            assistant_response = openai.ChatCompletion.create(
                model="gpt-4",  # Use your specific assistant model or API
                messages=[
                    {
                        "role": "system",
                        "content": "You are an assistant that analyzes PDF documents.",
                    },
                    {
                        "role": "user",
                        "content": f"Analyze the PDF document for the following prompt: {prompt_message}",
                        "attachments": [{"file": uploaded_file}],
                    },
                ],
            )

            # Extract the response from the assistant
            assistant_response_content = assistant_response["choices"][0]["message"]["content"]

            if assistant_response_content:
                temp_responses.append(assistant_response_content)
                # Clean the response using `remove_source_patterns`
                assistant_response_cleaned = fc.remove_source_patterns(assistant_response_content)
                answers_dict[prompt_name] = assistant_response_cleaned

                # Fill the document with the assistant's response
                fc.document_filler(doc_copy, prompt_name, assistant_response_cleaned)

        except Exception as e:
            st.error(f"Error generating response for {prompt_name}: {e}")
            continue

# Save the Generated Document
new_file_path = "to_pager_official.docx"
doc_copy.save(new_file_path)

# Provide Download Link
st.success(f"Document successfully generated!")
with open(new_file_path, "rb") as file:
    st.download_button(
        label="Download Generated Document",
        data=file,
        file_name="to_pager_official.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
