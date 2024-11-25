import openai
import streamlit as st
from docx import Document
import pandas as pd
from PyPDF2 import PdfReader
import to_pager_functions as fc

# Streamlit App Title
st.title("OpenAI Assistant: Document Generator with PDF Support")

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.write(
    """
    This app uses the preloaded prompt database (`prompt_db.xlsx`) and Word template 
    (`to_pager_template.docx`) to generate a document. You can upload a PDF file to extract
    its content and analyze it using predefined assistants.
    Ensure your OpenAI API key is set in Streamlit Secrets.
    """
)

# Access API Key from Streamlit Secrets
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Sidebar File Upload
# If using Streamlit's uploaded file
uploaded_file = st.sidebar.file_uploader("Upload a PDF File", type=["pdf"])

if not uploaded_file:
    st.warning("Please upload a file to begin.")
    st.stop()

# Confirm File Upload
st.sidebar.success(f"Uploaded File: {uploaded_file.name}")
# Preloaded Files
xlsx_file = "prompt_db.xlsx"
docx_file = "to_pager_template.docx"

try:
    prompt_db = pd.ExcelFile(xlsx_file)
    doc_copy = Document(docx_file)
except Exception as e:
    st.error(f"Error loading preloaded files: {e}")
    st.stop()

# Define the sections and assistants
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
try:
    client = openai.Client()  # Initialize OpenAI client
    message_file = client.files.create(file=uploaded_file, purpose="assistants")
    st.sidebar.write(f"File uploaded successfully with ID: {message_file.id}")
except Exception as e:
    st.error(f"Error uploading the file: {e}")
    st.stop()
try:
    # Create a vector store or use an existing one
    vector_store = client.beta.vector_stores.create(name="My Vector Store")
    st.sidebar.write(f"Vector store created with ID: {vector_store.id}")

    # Add the file to the vector store
    file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
        vector_store_id=vector_store.id,
        files=[uploaded_file]
    )
    st.sidebar.write(f"File indexed in vector store successfully: {file_batch.status}")

except Exception as e:
    st.error(f"Error adding file to vector store: {e}")
    st.stop()
# Process Sections
st.subheader("Document Generation in Progress")
temp_responses = []
answers_dict = {}

for section_name, section_data in sections.items():
    st.write(f"Processing section: {section_name}")
    sheet_names = section_data["sheet_names"]
    assistant_id = section_data["assistant_id"]

    try:
        prompt_list, additional_formatting_requirements = fc.prompts_retriever(
            xlsx_file, sheet_names
        )
    except Exception as e:
        st.error(f"Error processing prompts for {section_name}: {e}")
        continue

    for prompt_name, prompt_message in prompt_list:
        st.write(f"Processing prompt: {prompt_name}")

        try:
            # Include extracted text in the input
            input_message = (
                f"{prompt_message}"
            )

            # Query the assistant
            assistant_response = fc.separate_thread_answers(
                openai,
                prompt_message=input_message,
                additional_formatting_requirements=additional_formatting_requirements,
                assistant_identifier=assistant_id
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
