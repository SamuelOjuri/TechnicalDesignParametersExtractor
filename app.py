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
from monday_dot_com_interface import MondayDotComInterface
from dotenv import load_dotenv

# Add this after initializing the Gemini client
# Initialize Monday.com client
monday_api_token = os.environ.get("MONDAY_API_TOKEN")
monday_interface = MondayDotComInterface(monday_api_token) if monday_api_token else None

load_dotenv()

# Initialize the client
client = genai.Client(api_key=os.environ.get("GOOGLE_API_KEY"))

# Add reset function
def reset_app_state():
    """Clear all session state variables to reset the app"""
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    return

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
                # model="gemini-2.5-pro-exp-03-25",
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
                    # model="gemini-2.5-pro-exp-03-25",
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
            # model="gemini-2.5-pro-exp-03-25",
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

def extract_project_name_from_content(email_text, attachments_data):
    """
    Extract the project name from email content and attachments
    
    Returns:
        str: The extracted project name
    """
    # Extract text from email and attachments
    combined_text = extract_text_from_email(email_text, attachments_data)
    
    # Create a focused prompt for the LLM
    prompt = f"""
    Based on the following email content and attachments, extract the project name (drawing title) which is usually the project location.
    Return only the project name, nothing else.
    
    {combined_text}
    """
    
    # Send to Gemini for analysis
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        # model="gemini-2.5-pro-exp-03-25",
        contents=prompt
    )
    
    # Return the response as the project name
    return response.text.strip()

def extract_parameters_from_monday_project(project_details):
    """
    Extracts design parameters from a Monday.com project.
    
    Args:
        project_details: The project details from Monday.com
        
    Returns:
        dict: A dictionary of extracted parameters
    """
    # Initialize parameters with default values
    params = {
        "Post Code": "Not found",
        "Drawing Reference": "Not found",
        "Drawing Title": "Not found",
        "Revision": "Not found",
        "Date Received": "Not found",
        "Company": "Not found",
        "Contact": "Not found",
        "Reason for Change": "Amendment",  # Default for existing projects
        "Surveyor": "Not found",
        "Target U-Value": "Not found",
        "Target Min U-Value": "Not found", 
        "Fall of Tapered": "Not found",
        "Tapered Insulation": "Not found",
        "Decking": "Not found"
    }
    
    # First, extract data from main project item
    for col in project_details.get('column_values', []):
        # Extract Post Code from dropdown_mknfpjbt column (Zip Code)
        if col.get('id') == "dropdown_mknfpjbt" and col.get('text'):
            params["Post Code"] = col.get('text')
        
        # Extract Project Name
        elif col.get('id') == "text3__1":  # Project Name column
            if col.get('text'):
                params["Drawing Title"] = col.get('text')
            elif col.get('__typename') == "MirrorValue" and col.get('display_value'):
                params["Drawing Title"] = col.get('display_value')
    
    # Get today's date for Date Received (for amendments)
    from datetime import datetime
    params["Date Received"] = datetime.now().strftime("%Y-%m-%d")
    
    # Check if we have subitems (revisions) with more detailed information
    if project_details.get('subitems') and len(project_details['subitems']) > 0:
        # Use the most recent subitem (revision) for detailed information
        # Sort by ID in descending order to get the most recent one
        latest_subitem = sorted(project_details['subitems'], key=lambda x: x['id'], reverse=True)[0]
        
        print(f"DEBUG: Using latest subitem: {latest_subitem['name']}")
        
        # Extract Drawing Reference and Revision from subitem name (e.g., "16763_25.01 - A")
        if '_' in latest_subitem['name'] and ' - ' in latest_subitem['name']:
            parts = latest_subitem['name'].split('_')
            if len(parts) >= 2:
                ref_parts = parts[1].split(' - ')
                if len(ref_parts) >= 2:
                    params["Drawing Reference"] = f"{parts[0]}_{ref_parts[0]}"
                    params["Revision"] = ref_parts[1]
        
        # Map column IDs to parameter names for subitem values
        column_mappings = {
            # Direct mappings from Monday.com column IDs to our parameter names
            "mirror_12__1": "Company",           # Account column
            "mirror39__1": "Designer",           # Designer column
            "mirror_11__1": "Contact",           # Contact column
            "mirror92__1": "Surveyor",           # Surveyor column
            "mirror0__1": "Target U-Value",      # U-Value column
            "mirror12__1": "Target Min U-Value", # Min U-Value column
            "mirror22__1": "Fall of Tapered",    # Fall column
            "mirror875__1": "Tapered Insulation", # Product Type column
            "mirror75__1": "Decking",            # Deck Type column
            "mirror95__1": "Date Received",      # Date Received column
            "mirror03__1": "Reason for Change",  # Reason For Change column
        }
        
        # Process each column value in the subitem
        for col in latest_subitem.get('column_values', []):
            col_id = col.get('id')
            if col_id in column_mappings:
                param_name = column_mappings[col_id]
                
                # Try to get text value or display_value for MirrorValue
                if col.get('text') and col.get('text') != "None":
                    params[param_name] = col.get('text')
                elif col.get('__typename') == "MirrorValue" and col.get('display_value'):
                    params[param_name] = col.get('display_value')
        
        # Special handling for certain parameters like Target U-Value that might come from different sources
        # From the Postman response, we can see mirror034__1 is actually "% Wasteage" not "Target U-Value"
        # Let's map it correctly
        for col in latest_subitem.get('column_values', []):
            if col.get('id') == "mirror034__1" and (col.get('text') or (col.get('__typename') == "MirrorValue" and col.get('display_value'))):
                value = col.get('text') if col.get('text') else col.get('display_value')
                params["Target U-Value"] = value
    
    return params

def get_monday_column_mappings():
    """
    Returns a mapping of Monday.com column IDs to parameter names.
    This allows for easy customization of the column mappings.
    """
    return {
        "text3__1": "Drawing Title",      # Project Name column
        "text": "Post Code",              # Postcode column
        "text7": "Drawing Reference",     # Drawing Reference column
        "text79": "Revision",             # Revision column
        "date4": "Date Received",         # Date Received column
        "text2": "Company",               # Company column
        "text94": "Contact",              # Contact column
        "text5": "Surveyor",              # Surveyor column
        "text9": "Target U-Value",        # Target U-Value column
        "text6": "Target Min U-Value",    # Min U-Value column
        "text8": "Fall of Tapered",       # Fall of Tapered column
        "text4": "Tapered Insulation",    # Tapered Insulation column
        "text0": "Decking"                # Decking column
    }

# Streamlit app
def main():
    st.title("📧 Design Parameters Extractor")
    
    # Add refresh button at the top right
    col1, col2 = st.columns([5, 1])
    with col2:
        if st.button("🔄 Refresh"):
            reset_app_state()
            st.rerun()
    
    st.write("Upload email (.eml or .msg) files and/or PDFs to extract design parameters.")
    
    # Debug session state
    print("DEBUG: Session state:", st.session_state)
    
    # Initialize session state variables if they don't exist
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = False
    if 'search_results' not in st.session_state:
        st.session_state.search_results = None
    if 'email_data' not in st.session_state:
        st.session_state.email_data = None
    if 'project_name' not in st.session_state:
        st.session_state.project_name = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    
    # File uploader - updated to include .msg files
    uploaded_files = st.file_uploader(
        "Drag and drop email (.eml, .msg) or PDF files here", 
        type=["eml", "msg", "pdf"],
        accept_multiple_files=True
    )
    
    # Query input
    st.subheader("Query Parameters")
    default_query = """Extract the following design parameters from the documents: 
        - Post Code of Project Location: (Mostly found in the title block of the drawing attached to emails. If post code of drawing architect exists, ignore it and use the post code of the project location), 
        - Drawing Reference: (TaperedPlus Reference Number e.g. TP12345_00.01 - A), 
        - Drawing Title (The Project Name, usually the project location), 
        - Revision (Suffix of the drawing reference e.g. 00.01 - A), 
        - Date Received: (Date initial email was sent by customer), 
        - Company: (Client company requesting technical drawings or TaperedPlus services, usually the email sender requesting the drawings from TaperedPlus),
        - Contact: (Contact Person, usually the email sender requesting the drawings from TaperedPlus), 
        - Reason for Change: (Either 'Amendment' or 'New Enquiry' depending on the context of the email), 
        - Surveyor: (Name of the surveyor if provided), 
        - Target U-Value, 
        - Target Min U-Value, 
        - Fall of Tapered,
        - Tapered Insulation, 
        - Decking."""
    
    query = st.text_area("Enter your query for the AI analysis:", value=default_query, height=200)
    
    # Process button
    process_button = st.button("Process Files")
    
    # Only process files if:
    # 1. The process button was clicked (process_button)
    # 2. Files were uploaded (uploaded_files)
    # 3. Files haven't been processed yet (not st.session_state.processed_files)
    # This prevents re-processing on page reloads and ensures one-time processing
    if process_button and uploaded_files and not st.session_state.processed_files:
        with st.spinner("Processing files..."):
            all_extracted_text = ""
            results_data = []
            
            # First pass - process the first email file to extract the project name
            for uploaded_file in uploaded_files:
                if uploaded_file.name.lower().endswith((".eml", ".msg")):
                    # Create a temporary file to store the uploaded content
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as temp_file:
                        temp_file.write(uploaded_file.getvalue())
                        temp_file_path = temp_file.name
                    
                    try:
                        if uploaded_file.name.lower().endswith(".eml"):
                            # Process email file
                            header, body, attachments_data = process_eml_file(temp_file_path)
                        elif uploaded_file.name.lower().endswith(".msg"):
                            # Process Outlook .msg file
                            header, body, attachments_data = process_msg_file(temp_file_path)
                        
                        email_text = header + "\n" + body
                        st.session_state.email_data = {
                            'email_text': email_text,
                            'attachments_data': attachments_data
                        }
                        
                        # Break after processing the first email file
                        break
                    finally:
                        # Clean up the temporary file
                        if os.path.exists(temp_file_path):
                            os.remove(temp_file_path)
            
            # Store the uploaded files in session state for later processing
            st.session_state.uploaded_files = uploaded_files
            st.session_state.query = query
            st.session_state.processed_files = True
            # Rerun to show the next step
            st.rerun()
    
    # If files are processed but we need to get Monday.com data
    if st.session_state.processed_files and not st.session_state.processing_complete:
        # If we have email data, extract project name and search Monday.com
        if st.session_state.email_data and monday_interface:
            # Extract project name if not already done
            if not st.session_state.project_name:
                with st.spinner("Extracting project name from email..."):
                    st.session_state.project_name = extract_project_name_from_content(
                        st.session_state.email_data['email_text'], 
                        st.session_state.email_data['attachments_data']
                    )
            
            st.subheader("Email Analysis")
            st.write(f"Extracted Project Name: **{st.session_state.project_name}**")
            
            # Search for similar projects if not already done
            if not st.session_state.search_results and st.session_state.project_name:
                with st.spinner("Searching for similar projects in Monday.com..."):
                    st.session_state.search_results = monday_interface.check_project_exists(st.session_state.project_name)
            
            # Display project matches if any
            if st.session_state.search_results and st.session_state.search_results['exists'] and st.session_state.search_results['matches']:
                st.subheader("Matching Projects Found in Monday.com")
                
                # Create a radio button for user to select a project
                project_options = [f"{match['title']} (Similarity: {match['similarity']:.2f})" 
                                   for match in st.session_state.search_results['matches']]
                project_options.append("None of the above - Treat as new enquiry")
                
                selected_project_option = st.radio("Select the matching project:", project_options)
                
                st.warning("Please select a project and click 'Continue' to proceed")
                
                # Button to continue
                if st.button("Continue"):
                    if selected_project_option != "None of the above - Treat as new enquiry":
                        # User confirmed this is an amendment to an existing project
                        enquiry_type = "Amendment"
                        
                        with st.spinner("Retrieving project details from Monday.com..."):
                            # Get the selected project
                            selected_index = project_options.index(selected_project_option)
                            selected_project_id = st.session_state.search_results['matches'][selected_index]['id']
                            
                            # Get detailed project information
                            board_id = "1825117125"  # Board ID for Tapered Enquiry Maintenance
                            print(f"DEBUG: Searching for project with index {selected_index}")
                            print(f"DEBUG: Project name to search: {st.session_state.search_results['matches'][selected_index]['name']}")
                            
                            project_details, error = monday_interface.get_item_by_name_on_board(
                                board_id, st.session_state.search_results['matches'][selected_index]['name'])
                            
                            print("DEBUG: Retrieved Project Details: ", project_details)
                            print("DEBUG: Error: ", error)
                        
                        # Store results in session state
                        st.session_state.project_details = project_details
                        st.session_state.project_error = error
                        st.session_state.enquiry_type = enquiry_type
                        st.session_state.processing_complete = True
                        
                        # Rerun to display results
                        st.rerun()
                    else:
                        # User selected "None of the above" - treat as new enquiry
                        st.session_state.enquiry_type = "New Enquiry"
                        st.session_state.processing_complete = True
                        st.rerun()
            else:
                st.info(f"No matching projects found in Monday.com. Treating as a new enquiry.")
                st.session_state.enquiry_type = "New Enquiry"
                st.session_state.processing_complete = True
                st.rerun()
        else:
            # No email data or Monday integration not available
            st.session_state.enquiry_type = "New Enquiry"
            st.session_state.processing_complete = True
            st.rerun()
    
    # Display results after project selection/processing is complete
    if st.session_state.processing_complete:
        results_data = []
        
        # If we have project details from Monday.com, display them
        if hasattr(st.session_state, 'project_details') and st.session_state.project_details:
            with st.spinner("Extracting parameters from Monday.com project..."):
                # Parse the project details to extract parameters
                print("DEBUG: Extracting parameters from project details")
                params = extract_parameters_from_monday_project(st.session_state.project_details)
                print("DEBUG: Extracted params: ", params)
            
            # Display the extracted parameters
            st.write("The following parameters were extracted from Monday.com:")
            for key, value in params.items():
                if value and value != "Not found":
                    st.write(f"**{key}:** {value}")
            
            # Add to results data
            results_data.append(params)
            
            # Create DataFrame
            with st.spinner("Creating results data..."):
                df = pd.DataFrame(results_data)
            
            # Display dataframe
            st.subheader("Extracted Parameters")
            st.dataframe(df)
            
            # Create Excel download button
            with st.spinner("Generating Excel file..."):
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Parameters')
                    # For the AI analysis path, add full response sheet
                    if 'llm_response' in locals():
                        pd.DataFrame({'Response': [llm_response]}).to_excel(
                            writer, index=False, sheet_name='Full Response')
                
                buffer.seek(0)
            
            st.download_button(
                label="Download Results as Excel",
                data=buffer,
                file_name="Technical_Parameters.xlsx",
                mime="application/vnd.ms-excel"
            )
            
            # Replace the existing Process New Files button with this
            st.button("Process New Files", on_click=reset_app_state, key="process_new_files_button")
                
            return  # Exit early as we've loaded the data from Monday.com
        else:
            # Process files for new enquiry
            all_extracted_text = ""
            
            # Process all uploaded files
            if hasattr(st.session_state, 'uploaded_files'):
                for uploaded_file in st.session_state.uploaded_files:
                    # Create a temporary file to store the uploaded content
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as temp_file:
                        temp_file.write(uploaded_file.getvalue())
                        temp_file_path = temp_file.name
                    
                    try:
                        if uploaded_file.name.lower().endswith(".eml"):
                            # Process email file
                            with st.spinner(f"Processing email file: {uploaded_file.name}..."):
                                header, body, attachments_data = process_eml_file(temp_file_path)
                                email_text = header + "\n" + body
                                
                                # Extract text from email and attachments
                                extracted_text = extract_text_from_email(email_text, attachments_data)
                                all_extracted_text += f"\n\nEMAIL FILE: {uploaded_file.name}\n{extracted_text}\n{'='*50}\n"
                        
                        elif uploaded_file.name.lower().endswith(".msg"):
                            # Process Outlook .msg file
                            with st.spinner(f"Processing Outlook email file: {uploaded_file.name}..."):
                                header, body, attachments_data = process_msg_file(temp_file_path)
                                email_text = header + "\n" + body
                                
                                # Extract text from email and attachments
                                extracted_text = extract_text_from_email(email_text, attachments_data)
                                all_extracted_text += f"\n\nOUTLOOK EMAIL FILE: {uploaded_file.name}\n{extracted_text}\n{'='*50}\n"
                            
                        elif uploaded_file.name.lower().endswith(".pdf"):
                            # Process PDF file directly with Gemini
                            with st.spinner(f"Processing PDF file: {uploaded_file.name}..."):
                                pdf_text = process_pdf_with_gemini(uploaded_file.getvalue(), uploaded_file.name)
                                all_extracted_text += f"\n\nPDF FILE: {uploaded_file.name}\n{pdf_text}\n{'='*50}\n"
                    
                    finally:
                        # Clean up the temporary file
                        if os.path.exists(temp_file_path):
                            os.remove(temp_file_path)
                
                # Update the query to include the determined enquiry type
                if hasattr(st.session_state, 'enquiry_type') and st.session_state.enquiry_type:
                    # Make sure the query contains instructions to find the Reason for Change
                    query = st.session_state.query
                    if "Reason for Change" in query:
                        # Update the query to specify the determined enquiry type
                        query = query.replace("Reason for Change: (Either 'Amendment' or 'New Enquiry' depending on the context of the email)", 
                                            f"Reason for Change: ({st.session_state.enquiry_type})")
                
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
                    with st.spinner("Creating results data..."):
                        df = pd.DataFrame(results_data)
                    
                    # Display LLM response and dataframe
                    st.subheader("AI Analysis Results")
                    st.markdown(llm_response)
                    
                    st.subheader("Extracted Parameters")
                    st.dataframe(df)
                    
                    # Create Excel download button
                    with st.spinner("Generating Excel file..."):
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False, sheet_name='Parameters')
                            # For the AI analysis path, add full response sheet
                            pd.DataFrame({'Response': [llm_response]}).to_excel(
                                writer, index=False, sheet_name='Full Response')
                        
                        buffer.seek(0)
                    
                    st.download_button(
                        label="Download Results as Excel",
                        data=buffer,
                        file_name="Technical_Parameters.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                    
                    # Replace the existing Process New Files button with this
                    st.button("Process New Files", on_click=reset_app_state, key="process_new_files_button")
                else:
                    st.error("No text could be extracted from the uploaded files.")
            else:
                st.error("No files found in session state.")
    
    elif process_button and not uploaded_files:
        st.error("Please upload at least one file first.")

if __name__ == "__main__":
    main()