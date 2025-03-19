import os
import email
from email import policy
from email.parser import BytesParser
import tempfile
import io
import extract_msg  # Added for .msg file processing

import pandas as pd
import streamlit as st

# Update imports to use the newer client approach
from google import genai
from google.genai import types
from dotenv import load_dotenv

load_dotenv()

# Initialize the client
client = genai.Client(api_key=os.environ.get("GOOGLE_API_KEY"))

def process_eml_file(eml_file_path):
    """
    Processes a single .eml file:
      - Parses the email to extract header and body.
      - Extracts attachments and returns their data.
    
    Returns:
      header: string with key header fields.
      body: string containing the plain text body.
      attachments_data: list of dictionaries with attachment data.
    """
    # Open and parse the email file using the default email policy
    with open(eml_file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    # Extract some header fields
    header_info = (
        f"From: {msg.get('from', '')}\n"
        f"To: {msg.get('to', '')}\n"
        f"Subject: {msg.get('subject', '')}\n"
        f"Date: {msg.get('date', '')}\n"
    )
    
    # Extract the email body
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and not part.get_filename():
                body += part.get_content() + "\n"
    else:
        body = msg.get_content()
    
    # Process attachments
    attachments_data = []
    for part in msg.iter_attachments():
        filename = part.get_filename()
        if filename:
            attachment_data = {
                'filename': filename,
                'content': part.get_payload(decode=True)
            }
            attachments_data.append(attachment_data)
    
    return header_info, body, attachments_data

def process_pdf_with_gemini(pdf_content, filename):
    """
    Process PDF content using Gemini's File API
    
    Returns:
        text: Extracted text and information from the PDF
    """
    try:
        # Create a temporary file to use with Gemini
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_file.write(pdf_content)
            temp_file_path = temp_file.name
        
        # Create a prompt to extract text and information from the PDF
        prompt = "Please extract all text content from this PDF document, including text from tables, diagrams, and charts."
        
        # Process the PDF with Gemini using the client approach
        with open(temp_file_path, 'rb') as f:
            pdf_data = f.read()
            
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=[
                    types.Part.from_bytes(
                        data=pdf_data,
                        mime_type='application/pdf',
                    ),
                    prompt
                ]
            )
            
        return response.text
    
    except Exception as e:
        st.error(f"Error processing PDF with Gemini: {e}")
        return f"Error processing PDF: {str(e)}"
    finally:
        # Clean up the temporary file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

def process_multiple_pdfs(pdf_files):
    """
    Process multiple PDF files with Gemini
    
    Args:
        pdf_files: List of PDF file data (content and filename)
        
    Returns:
        combined_text: Combined text from all PDFs
    """
    combined_text = ""
    
    for pdf_file in pdf_files:
        filename = pdf_file['filename']
        content = pdf_file['content']
        
        # Create a temporary file
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_file.write(content)
            temp_file_path = temp_file.name
        
        try:
            # Process the PDF with Gemini using the client approach
            with open(temp_file_path, 'rb') as f:
                pdf_data = f.read()
                
                response = client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=[
                        types.Part.from_bytes(
                            data=pdf_data,
                            mime_type='application/pdf',
                        ),
                        "Extract all text and information from this PDF document."
                    ]
                )
                
            combined_text += f"\nPDF ATTACHMENT ({filename}):\n{response.text}\n\n"
        
        except Exception as e:
            st.error(f"Error processing PDF {filename} with Gemini: {e}")
            combined_text += f"\nPDF ATTACHMENT ({filename}) [Error: {str(e)}]\n\n"
        
        finally:
            # Clean up the temporary file
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
    
    return combined_text

def extract_text_from_email(email_text, attachments_data):
    """
    Extracts all text from email and attachments and returns as a single string.
    Uses Gemini for PDF processing.
    """
    combined_text = f"EMAIL CONTENT:\n{email_text}\n\n"
    
    # Collect PDF attachments
    pdf_attachments = []
    for attachment in attachments_data:
        filename = attachment['filename']
        content = attachment['content']
        
        if filename.lower().endswith(".pdf"):
            pdf_attachments.append({
                'filename': filename,
                'content': content
            })
        else:
            # For non-PDF attachments, just note they exist but weren't processed
            combined_text += f"\nATTACHMENT ({filename}) [Not processed - not a PDF]\n\n"
    
    # Process PDF attachments with Gemini
    if pdf_attachments:
        pdf_text = process_multiple_pdfs(pdf_attachments)
        combined_text += pdf_text
    
    return combined_text

def query_llm(all_text, query):
    """
    Sends the extracted text and query to Gemini.
    """
    # Construct a prompt that includes both the context and the query
    prompt = f"""
    Please analyze the following information extracted from emails and PDF documents:
    
    {all_text}
    
    QUESTION: {query}
    """
    
    # Get response from Gemini using the client approach
    with st.spinner("Analyzing Results..."):
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt,
            #generation_config={"temperature": 0}
        )
    
    return response.text

def process_msg_file(msg_file_path):
    """
    Processes a single .msg file (Outlook email format):
      - Parses the email to extract header and body.
      - Extracts attachments and returns their data.
    
    Returns:
      header: string with key header fields.
      body: string containing the plain text body.
      attachments_data: list of dictionaries with attachment data.
    """
    # Open and parse the Outlook message file
    msg = extract_msg.Message(msg_file_path)
    
    try:
        # Extract header fields
        header_info = (
            f"From: {msg.sender}\n"
            f"To: {msg.to}\n"
            f"Subject: {msg.subject}\n"
            f"Date: {msg.date}\n"
        )
        
        # Extract the email body
        body = msg.body
        
        # Process attachments
        attachments_data = []
        for attachment in msg.attachments:
            filename = attachment.longFilename or attachment.shortFilename
            if filename:
                attachment_data = {
                    'filename': filename,
                    'content': attachment.data
                }
                attachments_data.append(attachment_data)
        
        return header_info, body, attachments_data
    
    finally:
        # Close the msg file to release the file handle
        msg.close()

# Streamlit app
def main():
    st.title("📧 Design Parameters Extractor")
    st.write("Upload email (.eml or .msg) files and/or PDFs to extract design parameters.")
    
    # File uploader - updated to include .msg files
    uploaded_files = st.file_uploader(
        "Drag and drop email (.eml, .msg) or PDF files here", 
        type=["eml", "msg", "pdf"],
        accept_multiple_files=True
    )
    
    # Query input
    st.subheader("Query Parameters")
    default_query = """Extract the following design parameters from the documents: 
        - Post Code of Project Location: (Mostly found in the title block of the drawing attached to emails), 
        - Drawing Reference: (Either 'Amendment' or 'New Enquiry' depending on the context of the email), 
        - Drawing Title, 
        - Revision, 
        - Date Received: (Date initial email was sent by customer), 
        - Company: (Client company requesting technical drawings or services),
        - Contact: (Contact Person of the client company), 
        - Reason for Change: (if applicable), 
        - Surveyor: (Name of the surveyor if provided), 
        - Target U-Value, 
        - Target Min U-Value, 
        - Fall of Tapered,
        - Tapered Insulation, 
        - Decking."""
    
    query = st.text_area("Enter your query for the AI analysis:", value=default_query, height=200)
    
    # Process button
    process_button = st.button("Process Files")
    
    if process_button and uploaded_files:
        with st.spinner("Processing files..."):
            all_extracted_text = ""
            results_data = []
            
            for uploaded_file in uploaded_files:
                # Create a temporary file to store the uploaded content
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as temp_file:
                    temp_file.write(uploaded_file.getvalue())
                    temp_file_path = temp_file.name
                
                try:
                    if uploaded_file.name.lower().endswith(".eml"):
                        # Process email file
                        header, body, attachments_data = process_eml_file(temp_file_path)
                        email_text = header + "\n" + body
                        
                        # Extract text from email and attachments
                        extracted_text = extract_text_from_email(email_text, attachments_data)
                        all_extracted_text += f"\n\nEMAIL FILE: {uploaded_file.name}\n{extracted_text}\n{'='*50}\n"
                    
                    elif uploaded_file.name.lower().endswith(".msg"):
                        # Process Outlook .msg file
                        header, body, attachments_data = process_msg_file(temp_file_path)
                        email_text = header + "\n" + body
                        
                        # Extract text from email and attachments
                        extracted_text = extract_text_from_email(email_text, attachments_data)
                        all_extracted_text += f"\n\nOUTLOOK EMAIL FILE: {uploaded_file.name}\n{extracted_text}\n{'='*50}\n"
                        
                    elif uploaded_file.name.lower().endswith(".pdf"):
                        # Process PDF file directly with Gemini
                        pdf_text = process_pdf_with_gemini(uploaded_file.getvalue(), uploaded_file.name)
                        all_extracted_text += f"\n\nPDF FILE: {uploaded_file.name}\n{pdf_text}\n{'='*50}\n"
                
                finally:
                    # Clean up the temporary file
                    if os.path.exists(temp_file_path):
                        os.remove(temp_file_path)
            
            # Send text to Gemini for analysis
            if all_extracted_text:
                llm_response = query_llm(all_extracted_text, query)
                
                # Parse the LLM response into a structured format for the dataframe
                parameters = [
                    "Post Code", "Drawing Reference", "Drawing Title", "Revision", 
                    "Date Received", "Company", "Contact", "Reason for Change", 
                    "Surveyor", "Target U-Value", "Target Min U-Value", 
                    "Fall of Tapered", "Tapered Insulation", "Decking"
                ]
                
                # Simple parsing logic - this could be improved
                result_dict = {}
                for param in parameters:
                    # Look for the parameter in the response
                    pattern = rf"{param}:?\s*(.*?)(?:\n|$)"
                    import re
                    match = re.search(pattern, llm_response, re.IGNORECASE)
                    if match:
                        result_dict[param] = match.group(1).strip()
                    else:
                        result_dict[param] = "Not found"
                
                # Add to results data
                results_data.append(result_dict)
                
                # Create DataFrame
                df = pd.DataFrame(results_data)
                
                # Display LLM response and dataframe
                st.subheader("AI Analysis Results")
                st.markdown(llm_response)
                
                st.subheader("Extracted Parameters")
                st.dataframe(df)
                
                # Create Excel download button
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Parameters')
                    # Add the full LLM response in another sheet
                    pd.DataFrame({'Response': [llm_response]}).to_excel(
                        writer, index=False, sheet_name='Full Response')
                
                buffer.seek(0)
                
                st.download_button(
                    label="Download Results as Excel",
                    data=buffer,
                    file_name="Technical_Parameters.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("No text could be extracted from the uploaded files.")
    
    elif process_button and not uploaded_files:
        st.error("Please upload at least one file first.")

if __name__ == "__main__":
    main()