import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import pandas as pd
import to_pager_functions as fc
# Function to chunk large text into manageable sizes
def chunk_text(text, max_length=1000):
    """Split text into chunks that are within the API's max length."""
    return [text[i:i + max_length] for i in range(0, len(text), max_length)]

# Streamlit App Title
st.title("OpenAI Assistant: Document Generator with PDF Support")

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.write(
    """
    This app uses the preloaded prompt database (`prompt_db.xlsx`) and Word template 
    (`to_pager_template.docx`) to generate a document. Additionally, you can upload a PDF
    to process its content via AI assistants.
    Ensure your OpenAI API key is set in Streamlit Secrets.
    """
)

# Access API Key from Streamlit Secrets
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Sidebar PDF Upload
st.sidebar.header("Upload PDF")
uploaded_pdf = st.sidebar.file_uploader("Upload a PDF File", type=["pdf"])

if not uploaded_pdf:
    st.warning("Please upload a PDF file to begin processing.")
    st.stop()

# If PDF is uploaded
st.sidebar.success(f"Uploaded PDF: {uploaded_pdf.name}")
pdf_content = uploaded_pdf.read().decode("utf-8", errors="ignore")  # Read as binary and decode
st.sidebar.write("PDF uploaded successfully!")

# Split the PDF content into chunks
chunks = chunk_text(pdf_content)

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
        "assistant_id": "asst_ZwYHPxoqquAdDHmVyZrr8SgC",
    },
    "B. REFERENCE MARKET": {
        "sheet_names": ["RM_Prompts", "RM_Format_add"],
        "assistant_id": "asst_vy2MqKVgrmjCecSTRgg0y6oO",
    },
}

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
            responses = []
            for i, chunk in enumerate(chunks):
                input_message = f"{prompt_message}\n\nPDF Chunk {i+1}: {chunk}"

                # Generate response for each chunk
                assistant_response = fc.separate_thread_answers(
                    openai, input_message, additional_formatting_requirements, assistant_id
                )

                if assistant_response:
                    assistant_response = fc.remove_source_patterns(assistant_response)
                    responses.append(assistant_response)

            # Combine responses from all chunks
            combined_response = "\n\n".join(responses)
            temp_responses.append(combined_response)
            answers_dict[prompt_name] = combined_response

            # Fill document with the combined response
            fc.document_filler(doc_copy, prompt_name, combined_response)

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
