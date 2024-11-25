import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import pandas as pd
import to_pager_functions as fc
from PyPDF2 import PdfReader  # For handling PDFs

# Streamlit App Title
st.title("OpenAI Assistant: Document Generator with PDF Support")

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.write(
    """
    This app uses the preloaded prompt database (`prompt_db.xlsx`) and Word template 
    (`to_pager_template.docx`) to generate a document. Additionally, you can upload a PDF
    to extract text for further processing.
    """
)

# Access API Key from Streamlit Secrets
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Sidebar PDF Upload
st.sidebar.header("Upload PDF")
uploaded_pdf = st.sidebar.file_uploader("Upload a PDF File", type=["pdf"])
pdf_text = ""

if uploaded_pdf:
    st.sidebar.success(f"Uploaded PDF: {uploaded_pdf.name}")
    # Extract text from the uploaded PDF
    try:
        reader = PdfReader(uploaded_pdf)
        for page in reader.pages:
            pdf_text += page.extract_text()
        st.sidebar.write("PDF Extracted Text Preview:")
        st.sidebar.text_area("PDF Content", pdf_text, height=300)
    except Exception as e:
        st.sidebar.error(f"Error processing PDF: {e}")

# Preloaded Files
xlsx_file = "prompt_db.xlsx"
docx_file = "to_pager_template.docx"

try:
    prompt_db = pd.ExcelFile(xlsx_file)
    doc_copy = Document(docx_file)
except Exception as e:
    st.error(f"Error loading preloaded files: {e}")
    st.stop()

# Process Sections
st.subheader("Document Generation in Progress")
temp_responses = []
answers_dict = {}

sections = {
    "A. BUSINESS OPPORTUNITY AND GROUP OVERVIEW": {
        "sheet_names": ["BO_Prompts", "BO_Format_add"],
        "assistant_id": "asst_SRPZaRBdVx0Wv89dHUGU7z7u",
    },
    "B. REFERENCE MARKET": {
        "sheet_names": ["RM_Prompts", "RM_Format_add"],
        "assistant_id": "asst_vy2MqKVgrmjCecSTRgg0y6oO",
    },
}

if st.button('Process'):
    with st.spinner('Processing...'):
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
        
            for prompt_name, prompt_message in prompt_list:
                st.write(f"Processing prompt: {prompt_name}")
                
                # Include PDF content in the prompt
                if pdf_text:
                    prompt_message += f"\n\nAdditional Context from PDF:\n{pdf_text}"

                try:
                    assistant_response = fc.separate_thread_answers(
                        openai, prompt_message, additional_formatting_requirements, assistant_id
                    )
        
                    if assistant_response:
                        temp_responses.append(assistant_response)
                        assistant_response = fc.remove_source_patterns(assistant_response)
                        answers_dict[prompt_name] = assistant_response
        
                        fc.document_filler(doc_copy, prompt_name, assistant_response)
        
                except Exception as e:
                    st.error(f"Error generating response for {prompt_name}: {e}")
                    continue

    # Save and Provide Download Link
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
