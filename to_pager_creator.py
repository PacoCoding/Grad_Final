import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import pandas as pd
import to_pager_functions as fc

# Streamlit App Title
st.title("OpenAI Assistant: Document Generator with PDF Support")

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.write(
    """
    This app uses the preloaded prompt database (`prompt_db.xlsx`) and Word template 
    (`to_pager_template.docx`) to generate a document. Additionally, you can upload a PDF
    for analysis by your preconfigured assistants.
    """
)

# Access API Key from Streamlit Secrets
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Sidebar PDF Upload
# Sidebar PDF Upload
st.sidebar.header("Upload PDF")
uploaded_pdf = st.sidebar.file_uploader("Upload a PDF File", type=["pdf"])
pdf_uploaded = False
file_id = None

if uploaded_pdf:
    st.sidebar.success(f"Uploaded PDF: {uploaded_pdf.name}")
    pdf_uploaded = True

    try:
        # Use OpenAI File API to upload the PDF
        st.sidebar.write("Uploading PDF to OpenAI...")
        response = openai.File.create(
            file=uploaded_pdf,
            purpose="answers",  # The purpose must match your assistant's functionality
        )
        file_id = response["id"]
        st.sidebar.success(f"PDF uploaded successfully. File ID: {file_id}")

    except Exception as e:
        st.sidebar.error(f"Error uploading PDF: {e}")
        pdf_uploaded = False

# Preloaded Files
xlsx_file = "prompt_db.xlsx"
docx_file = "to_pager_template.docx"

try:
    prompt_db = pd.ExcelFile(xlsx_file)
    doc_copy = Document(docx_file)
except Exception as e:
    st.error(f"Error loading preloaded files: {e}")
    st.stop()

# Assistants for Each Section
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

# Process PDF Content with Assistants
if pdf_uploaded and st.button('Process'):
    with st.spinner('Processing...'):
        temp_responses = []
        answers_dict = {}

        for section_name, section_data in sections.items():
            st.write(f"Processing section: {section_name}")
            sheet_names = section_data["sheet_names"]
            assistant_id = section_data["assistant_id"]

            try:
                # Retrieve prompts and additional formatting requirements
                prompt_list, additional_formatting_requirements = fc.prompts_retriever(
                    xlsx_file, sheet_names
                )
            except Exception as e:
                st.error(f"Error processing prompts for {section_name}: {e}")
                continue

            # Use the assistant's `file_search` tool to analyze the uploaded PDF
            try:
                st.write(f"Analyzing PDF with assistant {assistant_id}...")
                file_analysis = openai.File.search(
                    assistant_id=assistant_id,
                    file_id=file_id,
                )

                # Extract relevant information from the file analysis
                pdf_analysis_result = file_analysis["choices"][0]["text"]
                st.write(f"Assistant Analysis Result: {pdf_analysis_result}")

            except Exception as e:
                st.error(f"Error analyzing PDF with assistant {assistant_id}: {e}")
                continue

            # Process each prompt with the enriched PDF analysis
            try:
                for prompt_name, prompt_message in prompt_list:
                    st.write(f"Processing prompt: {prompt_name}")

                    # Enrich the prompt with PDF analysis results
                    enriched_prompt = f"{prompt_message}\n\nInsights from PDF Analysis:\n{pdf_analysis_result}"

                    # Use the assistant to process the enriched prompt
                    assistant_response = fc.separate_thread_answers(
                        openai, enriched_prompt, additional_formatting_requirements, assistant_id
                    )

                    if assistant_response:
                        temp_responses.append(assistant_response)
                        assistant_response = fc.remove_source_patterns(assistant_response)
                        answers_dict[prompt_name] = assistant_response

                        # Fill the Word document with the assistant's response
                        fc.document_filler(doc_copy, prompt_name, assistant_response)

            except Exception as e:
                st.error(f"Error generating response for {section_name}: {e}")
                continue

    # Save the final document
    new_file_path = "to_pager_official.docx"
    doc_copy.save(new_file_path)

    st.success(f"Document successfully generated!")
    with open(new_file_path, "rb") as file:
        st.download_button(
            label="Download Generated Document",
            data=file,
            file_name="to_pager_official.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
