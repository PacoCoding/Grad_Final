#%%


"""
Packages for environment selection 
"""

import os  # Missing import for 'os'
from dotenv import find_dotenv, load_dotenv

"""
Packages for document writing
"""
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

import openai

import pandas as pd
import re
import time
import pickle

"""
==================================================================================================================
Assistant Question and Answering functions
==================================================================================================================
"""


def get_answer(client, run, thread):
    while not run.status == "completed":
        #print("Waiting for answer...")
        run = client.beta.threads.runs.retrieve(
            thread_id=thread.id,
            run_id=run.id
        )

"""
The following function is about loading the prompts we will use to fill the document.

This retrieves, from a .xlsx file, both the prompt and the placeholder metadata.

The placeholders corresponds to the ones in the .docx document and will be used 
select the appropriate place in which the assistant answer will be placed in 
the final document.
"""

def prompts_retriever(file_name, sheet_list):

    prompt_sheet = sheet_list[0]
    formatting_sheet = sheet_list[1]
        
    prompt_df = pd.read_excel(file_name, sheet_name=prompt_sheet)
    prompt_list = list(zip(prompt_df['Placeholder'], prompt_df['Prompt']))

    temp_df = pd.read_excel(file_name,sheet_name=formatting_sheet)

    additional_formatting_requirements = temp_df.iloc[0,0]

    return prompt_list, additional_formatting_requirements


def separate_thread_answers(client,
                            prompt_message, additional_formatting_requirements,
                            assistant_identifier):

    """
    We will iterate through all the prompts that are present in the .xlsx file.

    The prompt_list object is a list of tuples with:

        1) prompt_name = The placeholder in the .docx file that is associated with the 
        current prompt.

        2) prompt_message = The prompt itself that will be used to ask the assistant a question.
    """

    print(prompt_message)

    prompt_message_format = prompt_message + additional_formatting_requirements
    print(prompt_message_format)
    
    thread = client.beta.threads.create()

    """
    We essentially append our message to the current thread, to query the assistant
    """

    user_message = client.beta.threads.messages.create(
        thread_id=thread.id,
        role="user",
        content=prompt_message_format,
    )

    """
    This is the actual interaction with the OpenAI assistant
    """

    run = client.beta.threads.runs.create(
        thread_id=thread.id,
        assistant_id= assistant_identifier  
    )

    """
    In order to achieve a sequential workflow in which we move to the next prompt 
    only when the previous one was answered, we added this while loop to prevent
    moving forward until prompt completion.

    """
    run = get_answer(client, run, thread)

    """
    while run.status != 'completed':
        run = wait_for_run_completion(run, thread)
    """
    
    
    """
    We retrieve the entire list of messages that are part of the thread.

    By looping through the data attribute we are moving from the last message 
    upwards to the first.

!!!!!!!!!!!!!!!!!!!
    We will retrieve the first message that the answer from the assistant
    whose content is textual.
!!!!!!!!!!!!!!!!!!!
    """

    messages = client.beta.threads.messages.list(thread_id=thread.id)
    assistant_response = None

    for message in messages.data:  
        if message.role == "assistant":
            for content_block in message.content:
                if content_block.type == "text":
                    assistant_response = content_block.text.value
                    break
            if assistant_response:
                break

    return assistant_response


"""
==================================================================================================================
DOCUMENT FORMATTING
==================================================================================================================
"""

def remove_source_patterns(text):
    """
    Removes patterns like   from the     
    Args:
        text (str): The input string containing potential patterns to remove.
    
    Returns:
        str: The cleaned text without the specified patterns.
    """
    # Define the regular expression to match the pattern
    pattern = r"【\d+:\d+†source】"
    
    # Use re.sub to remove all occurrences of the pattern
    cleaned_text = re.sub(pattern, "", text)
    
    # Return the cleaned text
    return cleaned_text


def document_filler(doc_copy, prompt_name, assistant_response):
    #First we loop through all the paragraphs.
    for paragraph in doc_copy.paragraphs:

        #If the prompt_name correspond to the placeholder making up the paragraph
        #we move to the filling part
        if prompt_name in paragraph.text:
            
            #This is for formatting reasons to avoid alignment problems
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            #Then, we move to the run objects inside the paragraph.
            #The reason is that in this way, when we replace the placeholder 
            #we will keep the placeholder's formatting
            for run in paragraph.runs:
                if prompt_name in run.text:
                    run.text = run.text.replace(prompt_name, assistant_response)
# %%
