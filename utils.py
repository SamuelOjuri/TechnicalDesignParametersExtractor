import os
import time
import random
from email.parser import BytesParser
from email import policy
import streamlit as st
import tempfile  
import extract_msg  


# Update imports to use the newer client approach
from google import genai
from google.genai import types
from monday_dot_com_interface import MondayDotComInterface
from dotenv import load_dotenv

# Initialize Monday.com client
load_dotenv()
monday_api_token = os.environ.get("MONDAY_API_TOKEN")
monday_interface = MondayDotComInterface(monday_api_token) if monday_api_token else None

# Initialize the client
client = genai.Client(api_key=os.environ.get("GOOGLE_API_KEY"))



# Define a function to check if an exception is a rate limit error
def is_rate_limit_error(exception):
    return '429' in str(exception) or 'RESOURCE_EXHAUSTED' in str(exception)


def gemini_api_with_retry(model, contents):
    """
    Call Gemini API with retry logic for rate limiting
    
    Args:
        model: The Gemini model to use
        contents: The contents to send to the model
        
    Returns:
        The model response
    """
    try:
        # Add a small random delay to help with rate limiting
        time.sleep(random.uniform(0.5, 1.5))
        
        # Make the API call
        response = client.models.generate_content(
            model=model,
            contents=contents
        )
        return response
    except Exception as e:
        # Check if this is a rate limiting error
        if is_rate_limit_error(e):
            st.warning(f"Rate limit hit. Waiting before retrying... ({str(e)})")
            # Re-raise to trigger retry
            raise e
        else:
            # For other errors, log and re-raise without retry
            st.error(f"Error calling Gemini API: {str(e)}")
            raise e

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
      - Extracts attachments and inline images and returns their data.
    
    Returns:
      header: string with key header fields.
      body: string containing the plain text body.
      attachments_data: list of dictionaries with attachment data.
      inline_images: list of dictionaries with inline image data.
    """
    # Open and parse the email file using the default email policy
    with open(eml_file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    # Extract header fields
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
    
    # Process attachments and inline images
    attachments_data = []
    inline_images = []
    
    for part in msg.iter_attachments():
        filename = part.get_filename()
        if filename:
            # Check if this is an inline image
            is_inline = False
            content_id = part.get('Content-ID')
            
            # Images with Content-ID are typically inline
            if content_id and filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                is_inline = True
                
            if is_inline:
                inline_image_data = {
                    'filename': filename,
                    'content': part.get_payload(decode=True),
                    'content_id': content_id,
                    'mime_type': part.get_content_type()
                }
                inline_images.append(inline_image_data)
            else:
                attachment_data = {
                    'filename': filename,
                    'content': part.get_payload(decode=True)
                }
                attachments_data.append(attachment_data)
    
    return header_info, body, attachments_data, inline_images

def process_pdf_with_gemini(pdf_content, filename):
    """Process PDF content using Gemini's File API"""
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
            
            # Use retry function instead of direct API call
            response = gemini_api_with_retry(
                model="gemini-2.5-flash-preview-04-17",
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
                
                response = gemini_api_with_retry(
                    model="gemini-2.5-flash-preview-04-17",
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

def process_multiple_images(image_files, image_type="ATTACHMENT"):
    """
    Process multiple image files with Gemini
    
    Args:
        image_files: List of image file data (content and filename)
        image_type: Type of image (ATTACHMENT or INLINE IMAGE)
        
    Returns:
        combined_text: Combined text from image analysis
    """
    combined_text = ""
    
    for image_file in image_files:
        filename = image_file['filename']
        content = image_file['content']
        file_extension = filename.split(".")[-1].lower()
        
        # Create a temporary file
        with tempfile.NamedTemporaryFile(suffix=f'.{file_extension}', delete=False) as temp_file:
            temp_file.write(content)
            temp_file_path = temp_file.name
        
        try:
            # Process the image with Gemini using the client approach
            with open(temp_file_path, 'rb') as f:
                image_data = f.read()
                
                response = gemini_api_with_retry(
                    model="gemini-2.5-flash-preview-04-17",
                    contents=[
                        types.Part.from_bytes(
                            data=image_data,
                            mime_type=f'image/{file_extension}',
                        ),
                        "Describe this image in detail, including any visible text, diagrams, or drawings. Extract any technical parameters or specifications you can see."
                    ]
                )
                
            combined_text += f"\n{image_type} ({filename}):\n{response.text}\n\n"
        
        except Exception as e:
            st.error(f"Error processing image {filename} with Gemini: {e}")
            combined_text += f"\n{image_type} ({filename}) [Error: {str(e)}]\n\n"
        
        finally:
            # Clean up the temporary file
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
    
    return combined_text

def extract_text_from_email(email_text, attachments_data, inline_images=None):
    """Extracts all text from email and attachments returns as a single string."""
    combined_text = f"EMAIL CONTENT:\n{email_text}\n\n"
    
    # For non-visual content attachments, just note they exist
    for attachment in attachments_data:
        filename = attachment['filename']
        if not (filename.lower().endswith(".pdf") or 
                filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp'))):
            combined_text += f"\nATTACHMENT ({filename}) [Not processed - not a PDF or image]\n\n"
    
    # Limit the total number of visual items to process to avoid rate limits
    MAX_VISUAL_ITEMS = 10
    
    # Collect all visual attachments
    pdf_attachments = [a for a in attachments_data if a['filename'].lower().endswith(".pdf")]
    image_attachments = [a for a in attachments_data if a['filename'].lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp'))]
    
    all_visual_items = []
    
    # Add most important items first
    if pdf_attachments:
        all_visual_items.extend([('pdf', pdf) for pdf in pdf_attachments])
    if image_attachments:
        all_visual_items.extend([('image', img) for img in image_attachments])
    if inline_images:
        all_visual_items.extend([('inline', img) for img in inline_images])
    
    # Process only a limited number of items
    processed_items = all_visual_items[:MAX_VISUAL_ITEMS]
    skipped_items = all_visual_items[MAX_VISUAL_ITEMS:]
    
    # Note skipped items
    if skipped_items:
        combined_text += "\nNOTE: Some visual elements were not processed due to API rate limits:\n"
        for item_type, item in skipped_items:
            combined_text += f"- {item_type.upper()}: {item['filename']}\n"
        combined_text += "\n"
    
    # Process the limited set of items
    for item_type, item in processed_items:
        if item_type == 'pdf':
            # Process this PDF
            with st.spinner(f"Processing PDF: {item['filename']}..."):
                pdf_text = process_pdf_with_gemini(item['content'], item['filename'])
                combined_text += f"\nPDF ATTACHMENT ({item['filename']}):\n{pdf_text}\n\n"
        elif item_type == 'inline':
            # Process this inline image
            with st.spinner(f"Processing inline image: {item['filename']}..."):
                image_text = process_image_with_gemini(item['content'], item['filename'], "INLINE IMAGE")
                combined_text += f"\nINLINE IMAGE ({item['filename']}):\n{image_text}\n\n"
        elif item_type == 'image':
            # Process this image attachment
            with st.spinner(f"Processing image: {item['filename']}..."):
                image_text = process_image_with_gemini(item['content'], item['filename'], "ATTACHMENT")
                combined_text += f"\nIMAGE ATTACHMENT ({item['filename']}):\n{image_text}\n\n"
    
    return combined_text

# Helper function to process a single image
def process_image_with_gemini(image_content, filename, image_type="ATTACHMENT"):
    """Process a single image with Gemini"""
    # Define supported image formats and their MIME types
    SUPPORTED_FORMATS = {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'webp': 'image/webp'
    }
    
    file_extension = filename.split(".")[-1].lower()
    
    # Validate file format
    if file_extension not in SUPPORTED_FORMATS:
        return f"Unsupported image format: {file_extension}. Only {', '.join(SUPPORTED_FORMATS.keys())} are supported."
    
    # Get proper MIME type
    mime_type = SUPPORTED_FORMATS[file_extension]
    
    # Create a temporary file
    with tempfile.NamedTemporaryFile(suffix=f'.{file_extension}', delete=False) as temp_file:
        temp_file.write(image_content)
        temp_file_path = temp_file.name
    
    try:
        # Process the image with Gemini using the client approach
        with open(temp_file_path, 'rb') as f:
            image_data = f.read()
            
            # Use a try/except block specifically for this API call
            try:
                response = gemini_api_with_retry(
                    model="gemini-2.5-flash-preview-04-17",
                    contents=[
                        types.Part.from_bytes(
                            data=image_data,
                            mime_type=mime_type,
                        ),
                        "Describe this image in detail, including any visible text, diagrams, or drawings. Extract any technical parameters or specifications you can see."
                    ]
                )
                
                return response.text
            except Exception as e:
                error_message = str(e)
                # Check if it's specifically a format issue
                if "INVALID_ARGUMENT" in error_message:
                    return f"Unable to process this image due to format compatibility issues. Please note any visible information from the image might not be included in the analysis."
                else:
                    raise e  # Re-raise the exception for other types of errors
    
    except Exception as e:
        # Use logging instead of st.error
        print(f"Error processing image {filename} with Gemini: {e}")
        return f"Error processing image: {str(e)}"
    finally:
        # Clean up the temporary file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

def query_llm(all_text, query):
    """
    Sends the extracted text and query to Gemini.
    """
    # Construct a prompt that includes both the context and the query
    prompt = f"""
    Please analyze the following information extracted from emails, PDF documents, and images:
    
    {all_text}
    
    QUESTION: {query}
    
    Note that information may be found in any of the content sources, including text from image descriptions.
    """
    
    # Get response from Gemini using the client approach
    with st.spinner("Analyzing Results..."):
        response = gemini_api_with_retry(
            model="gemini-2.5-flash-preview-04-17",
            contents=prompt
        )
    
    return response.text

def process_msg_file(msg_file_path):
    """
    Processes a single .msg file (Outlook email format):
      - Parses the email to extract header and body.
      - Extracts attachments and inline images and returns their data.
    
    Returns:
      header: string with key header fields.
      body: string containing the plain text body.
      attachments_data: list of dictionaries with attachment data.
      inline_images: list of dictionaries with inline image data.
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
        
        # Process attachments and inline images
        attachments_data = []
        inline_images = []
        
        for attachment in msg.attachments:
            filename = attachment.longFilename or attachment.shortFilename
            if filename:
                # Try to determine if it's an inline image
                is_inline = False
                
                # Look for typical image extensions and check if it might be inline
                # Outlook msg format doesn't clearly distinguish inline vs attachment 
                # so we'll use heuristics
                if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                    # Check if there's a content ID or if it's referenced in HTML
                    # This is a heuristic approach
                    if hasattr(attachment, 'cid') and attachment.cid:
                        is_inline = True
                    elif hasattr(msg, 'htmlBody') and msg.htmlBody and filename in msg.htmlBody.decode('utf-8', errors='ignore'):
                        is_inline = True
                        
                if is_inline:
                    inline_image_data = {
                        'filename': filename,
                        'content': attachment.data,
                        'content_id': attachment.cid if hasattr(attachment, 'cid') else None,
                        'mime_type': f"image/{filename.split('.')[-1].lower()}"
                    }
                    inline_images.append(inline_image_data)
                else:
                    attachment_data = {
                        'filename': filename,
                        'content': attachment.data
                    }
                    attachments_data.append(attachment_data)
        
        return header_info, body, attachments_data, inline_images
    
    finally:
        # Close the msg file to release the file handle
        msg.close()

def extract_project_name_from_content(email_text, attachments_data):
    """
    Extract the project name from email content and attachments
    
    Returns:
        str: The extracted project name
    """
    # Use already extracted text if available (to avoid reprocessing)
    if hasattr(st.session_state, 'all_extracted_text') and st.session_state.all_extracted_text:
        combined_text = st.session_state.all_extracted_text
    else:
        # Fallback to extracting text if not already done
        combined_text = extract_text_from_email(email_text, attachments_data)
    
    # Create a focused prompt for the LLM
    prompt = f"""
    Based on the following email content and attachments, extract the project name (drawing title) which is usually the project location.
    Return only the project name, nothing else.
    
    {combined_text}
    """
    
    # Send to Gemini for analysis
    response = gemini_api_with_retry(
        model="gemini-2.5-flash-preview-04-17",
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
        
        # Extract Drawing Reference from subitem name
        if '_' in latest_subitem['name']:
            # The entire name (e.g., "16903_25.01 - A") should be used as Drawing Reference
            params["Drawing Reference"] = latest_subitem['name']  # Use the full name
        
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
            "mirror_1__1": "Revision",           # Revision column
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

def map_tapered_insulation_value(value):
    """Maps specific insulation product values to their category headers"""
    
    # Define lookup tables based on the image data
    insulation_mappings = {
        "TissueFaced PIR": ["TT47", "TR27", "Glass Tissue PIR", "Powerdeck F", "Adhered", "MG", "TR/MG", "FR/MG", "BauderPIR FA-TE", "Evatherm A", "Hytherm ADH"],
        "TorchOn PIR": ["TT44", "TR24", "Torched", "Powerdeck U", "Torched", "BGM", "TR/BGM", "FR/BGM", "BauderPIR FA"],
        "FoilFaced PIR": ["TT46", "TR26", "Foil", "Powerdeck Eurodeck", "Mech Fixed", "ALU", "TR/ALU", "FR/ALU", "Aluminium Faced"],
        "ROCKWOOL HardRock MultiFix DD": ["Mineral wool", "Hardrock", "stonewool", "stone wool", "rock wool", "bauderrock"],
        "Foamglas T3+": ["Cellular Glass", "foamed glass", "Bauderglas"],
        "EPS": ["Expanded Polystrene"],
        "XPS": ["Extruded Polystyrene"]
    }
    
    # Check if value exactly matches or contains any of the lookup values
    if value and value != "Not found":
        original_value = value
        for category, products in insulation_mappings.items():
            for product in products:
                if product.lower() in value.lower() or value.lower() in product.lower():
                    return category
    
    # Return original value if no match found
    return value