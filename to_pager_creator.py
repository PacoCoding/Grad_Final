import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import pandas as pd
import to_pager_functions as fc
from PyPDF2 import PdfReader
# Streamlit App Title
st.title("Custom Assistant with File Upload and Vector Store")

# Sidebar Instructions
st.sidebar.header("Instructions")
st.sidebar.write(
    """
    Upload a PDF file for analysis by your pre-configured assistants. The file will be
    indexed in a vector store, allowing the assistants to search and analyze its content.
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

# Define Pre-Created Assistants
assistants = {
    "Business Analysis": "asst_ZwYHPxoqquAdDHmVyZrr8SgC",
    "Market Analysis": "asst_vy2MqKVgrmjCecSTRgg0y6oO",
}

# Upload the File to OpenAI
try:
    client = openai.Client()  # Initialize OpenAI client
    message_file = client.files.create(file=uploaded_file, purpose="assistants")
    st.sidebar.write(f"File uploaded successfully with ID: {message_file.id}")
except Exception as e:
    st.error(f"Error uploading the file: {e}")
    st.stop()

# Add File to Vector Store
try:
    # Create a vector store or use an existing one
    vector_store = client.beta.vector_stores.create(name="My Vector Store")
    st.sidebar.write(f"Vector store created with ID: {vector_store.id}")

    # Add the file to the vector store
    file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
        vector_store_id=vector_store.id,
        files=[uploaded_file],
    )
    st.sidebar.write(f"File indexed in vector store successfully: {file_batch.status}")

except Exception as e:
    st.error(f"Error adding file to vector store: {e}")
    st.stop()

# Query Each Assistant with the Indexed File
results = {}

for assistant_name, assistant_id in assistants.items():
    st.write(f"Querying {assistant_name} Assistant...")
    query = st.text_input(f"Enter your question for {assistant_name}:", key=assistant_id)

    if st.button(f"Submit Query to {assistant_name}", key=f"submit_{assistant_id}"):
        try:
            with st.spinner(f"Querying {assistant_name} Assistant..."):
                # Query the assistant with the vector store
                response = client.beta.assistants.query(
                    assistant_id=assistant_id,
                    vector_store_id=vector_store.id,
                    query=query,
                )
                results[assistant_name] = response["answer"]
                st.write(f"Response from {assistant_name} Assistant:")
                st.write(response["answer"])

        except Exception as e:
            st.error(f"Error querying {assistant_name} Assistant: {e}")

# Option to Download Results
if results:
    st.subheader("Download Responses")
    all_results = "\n\n".join([f"{name}:\n{response}" for name, response in results.items()])
    st.download_button(
        label="Download All Responses",
        data=all_results,
        file_name="assistant_responses.txt",
        mime="text/plain",
    )
