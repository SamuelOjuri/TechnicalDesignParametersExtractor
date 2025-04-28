"""
app.py  –  Design Parameters Extractor + Chat Copilot
====================================================
All heavy‑lifting helpers live in **utils.py**; `DEFAULT_QUERY` lives in **constants.py**.
This file focuses on the UI orchestration and column layout.
"""

# ---------------------------------------------------------------------------
# Imports & initialisation
# ---------------------------------------------------------------------------
import os, tempfile, io, re, time
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google import genai
from monday_dot_com_interface import MondayDotComInterface

from utils import (
    reset_app_state,
    is_rate_limit_error,
    gemini_api_with_retry,
    query_llm,
    extract_text_from_email,
    process_eml_file,
    process_msg_file,
    process_pdf_with_gemini,
    extract_project_name_from_content,
    extract_parameters_from_monday_project,
    map_tapered_insulation_value,
)
from constants import DEFAULT_QUERY

# ── env / clients ───────────────────────────────────────────────────────────
load_dotenv()
client               = genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))
MONDAY_API_TOKEN     = os.getenv("MONDAY_API_TOKEN")
monday_interface     = MondayDotComInterface(MONDAY_API_TOKEN) if MONDAY_API_TOKEN else None

# ---------------------------------------------------------------------------
# CHAT PANEL (right column)
# ---------------------------------------------------------------------------

def show_chat() -> None:
    """Assistant copilot for the parameters."""
    st.session_state.setdefault("chat_history", [])
    st.session_state.setdefault("extracted_params_dict", None)

    # replay history
    for m in st.session_state["chat_history"]:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    disabled = st.session_state["extracted_params_dict"] is None
    prompt = st.chat_input("Type a question…", disabled=disabled)

    if disabled:
        st.info("Upload and process files first, then ask me anything 🙂")
        return

    if prompt:
        st.session_state["chat_history"].append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # quick command to show raw text
        if prompt.strip().lower() == "/raw":
            with st.chat_message("assistant"):
                st.markdown("Raw extracted text:")
                st.text_area("Raw", st.session_state.get("all_extracted_text", "<none>"), height=300)
            st.session_state["chat_history"].append({"role": "assistant", "content": "📄 Raw text displayed."})
            st.rerun()

        # build system context
        params_text = "\n".join(f"• **{k}**: {v}" for k, v in st.session_state["extracted_params_dict"].items())
        system = (
            "You are a roofing‑design assistant. Use the parameters below when answering; "
            "ask clarifying questions only when necessary.\n\n" + params_text
        )
        with st.spinner("Thinking …"):
            resp = gemini_api_with_retry("gemini-2.5-flash-preview-04-17", [system, prompt])

        with st.chat_message("assistant"):
            st.markdown(resp.text)
        st.session_state["chat_history"].append({"role": "assistant", "content": resp.text})
        st.rerun()

# ---------------------------------------------------------------------------
# EXTRACT PIPELINE (left column helpers)
# ---------------------------------------------------------------------------

def process_uploaded_files(uploaded_files) -> str:
    """Run through every uploaded file and return concatenated extracted text."""
    all_text = ""
    for f in uploaded_files:
        suffix = f".{f.name.split('.')[-1]}"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(f.getvalue())
            path = tmp.name
        try:
            if f.name.lower().endswith(".eml"):
                header, body, att, inline = process_eml_file(path)
                email_text = header + "\n" + body
                if st.session_state.email_data is None:
                    st.session_state.email_data = {"email_text": email_text, "attachments_data": att}
                extracted = extract_text_from_email(email_text, att, inline)
                all_text += f"\n\nEMAIL FILE: {f.name}\n{extracted}\n{'='*50}\n"
            elif f.name.lower().endswith(".msg"):
                header, body, att, inline = process_msg_file(path)
                email_text = header + "\n" + body
                if st.session_state.email_data is None:
                    st.session_state.email_data = {"email_text": email_text, "attachments_data": att}
                extracted = extract_text_from_email(email_text, att, inline)
                all_text += f"\n\nOUTLOOK EMAIL FILE: {f.name}\n{extracted}\n{'='*50}\n"
            elif f.name.lower().endswith(".pdf"):
                pdf_text = process_pdf_with_gemini(f.getvalue(), f.name)
                all_text += f"\n\nPDF FILE: {f.name}\n{pdf_text}\n{'='*50}\n"
        finally:
            os.remove(path)
    return all_text


# ---------------------------------------------------------------------------
# RESULTS DISPLAY + DOWNLOAD
# ---------------------------------------------------------------------------

def present_results(df: pd.DataFrame, extra_llm_response: str | None = None) -> None:
    """Show DF, create download, cache to session for chat."""
    st.subheader("Extracted Parameters")
    st.dataframe(df, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Parameters")
        if extra_llm_response:
            pd.DataFrame({"Response": [extra_llm_response]}).to_excel(writer, index=False, sheet_name="Full Response")
    buffer.seek(0)

    st.download_button("Download as Excel", buffer, "Technical_Parameters.xlsx", mime="application/vnd.ms-excel")

    # cache for chat
    st.session_state["extracted_params_df"]   = df
    st.session_state["extracted_params_dict"] = df.iloc[0].to_dict()

# ---------------------------------------------------------------------------
# MAIN APP (UI orchestration)
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Design Parameters Extractor", page_icon="📧", layout="wide")

    # Create a three-column layout with wider empty columns on the sides
    left_spacer, center_content, right_spacer = st.columns([1, 3, 1])
    
    # Put all content in the center column
    with center_content:
        # ===== header with centered logo/title =====
        st.markdown("<div style='text-align: center; margin-bottom: 30px;'><h1>📧 Design Parameters Extractor</h1></div>", unsafe_allow_html=True)
        
        # Upload section with a card-like appearance
        st.markdown("""
        <div style='background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h3>1️⃣ Upload & Process</h3>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_files = st.file_uploader("Drag + drop email/PDF files", type=["eml", "msg", "pdf"], accept_multiple_files=True)
        
        # Create a row with columns for centered button placement with more space
        button_col1, button_col2, button_col3 = st.columns([2, 2, 2])
        with button_col2:
            process_clicked = st.button("▶️ Process Files", use_container_width=True)
        
        # Init session keys
        for k, default in {
            "processed_files": False,
            "processing_complete": False,
            "email_data": None,
            "project_name": None,
            "search_results": None,
        }.items():
            st.session_state.setdefault(k, default)
        
        if process_clicked and uploaded_files and not st.session_state.processed_files:
            with st.spinner("Extracting …"):
                all_text = process_uploaded_files(uploaded_files)
            st.session_state.all_extracted_text = all_text
            st.session_state.processed_files   = True
            st.rerun()

        # ------------------------------------------------------------------
        # decide Amendment vs New‑Enquiry (Monday.com)
        # ------------------------------------------------------------------
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
        # ------------------------------------------------------------------
        # RESULTS section - always show the header, conditionally show content
        # ------------------------------------------------------------------
        # Add visual divider before results section
        st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)
        st.markdown("""
        <div style='background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h3>2️⃣ Analysis Results</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Show results only if processing is complete, otherwise show info message
        if not st.session_state.processing_complete:
            st.info("Upload and process files to see analysis results here.")
        else:
            results_data = []
            if st.session_state.enquiry_type == "Amendment" and st.session_state.get("project_details"):
                with st.spinner("Extracting parameters from Monday.com project …"):
                    # Parse the project details to extract parameters
                    print("DEBUG: Extracting parameters from project details")
                    params = extract_parameters_from_monday_project(st.session_state.project_details)
                    print("DEBUG: Extracted params: ", params)

                # Display the extracted parameters
                st.write("The following parameters were extracted from Monday.com:")
                for key, value in params.items():
                    if value and value != "Not found":
                        st.write(f"**{key}:** {value}")

                df = pd.DataFrame([params])
                present_results(df)
            else:
                # New Enquiry – use Gemini to parse `all_extracted_text`
                if hasattr(st.session_state, 'all_extracted_text') and st.session_state.all_extracted_text:
                    # Update the query to include the determined enquiry type
                    query = DEFAULT_QUERY
                    if hasattr(st.session_state, 'enquiry_type') and st.session_state.enquiry_type:
                        # Make sure the query contains instructions to find the Reason for Change
                        if "Reason for Change" in query:
                            # Update the query to specify the determined enquiry type
                            query = query.replace("Reason for Change: (Either 'Amendment' or 'New Enquiry' depending on the context of the email)", 
                                                f"Reason for Change: ({st.session_state.enquiry_type})")
                    
                    # with st.spinner("Analysing Extracted Data …"):
                    resp = query_llm(st.session_state.all_extracted_text, query)
                    
                    # Display LLM response
                    st.subheader("AI Analysis Results")
                    st.markdown(resp)
                    
                    df_row = {}
                    for p in [
                        "Post Code","Drawing Reference","Drawing Title","Revision","Date Received","Company","Contact","Reason for Change","Surveyor","Target U-Value","Target Min U-Value","Fall of Tapered","Tapered Insulation","Decking"]:
                        m = re.search(rf"{p}:?\s*(.*?)(?:\n|$)", resp, re.I)
                        val = m.group(1).strip() if m else "Not found"
                        
                        # Remove leading asterisks from all values
                        val = re.sub(r'^\*+\s*', '', val)
                        
                        # Special processing for specific parameters
                        if p == "Tapered Insulation":
                            val = map_tapered_insulation_value(val)
                        # For Post Code, extract just the postcode area (initial letters)
                        elif p == "Post Code":
                            # First clean up any formatting from the LLM response
                            cleaned_value = re.sub(r'^\s*of Project Location:?\*?\s*', '', val, flags=re.IGNORECASE)
                            cleaned_value = cleaned_value.strip()
                            
                            # Check if the value indicates "not provided" or similar
                            if re.search(r'not\s+provided|not\s+found|none', cleaned_value, re.IGNORECASE):
                                val = "Not provided"
                            else:
                                # Define UK postcode pattern
                                uk_postcode_pattern = r'([A-Z]{1,2})[0-9]'
                                postcode_match = re.search(uk_postcode_pattern, cleaned_value.upper())
                                if postcode_match:
                                    val = postcode_match.group(1)
                                else:
                                    # Keep original value if it doesn't match a postcode pattern
                                    val = cleaned_value
                        
                        df_row[p] = val

                    df = pd.DataFrame([df_row])
                    present_results(df, resp)

            # Reset button
            st.button("Process New Files", on_click=reset_app_state, use_container_width=True)

        # Move chat section OUTSIDE the if block so it's always visible
        # Add visual divider before chat section
        st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)
        st.markdown("""
        <div style='background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h3>3️⃣ Ask about the parameters</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Show chat interface
        show_chat()
        
        # After all content but before the chat input - MOVED INSIDE center_content
        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
        reset_col1, reset_col2, reset_col3 = st.columns([2, 2, 2])
        with reset_col2:
            reset_button = st.button("🔄 Reset App", help="Reset app", on_click=reset_app_state, use_container_width=True)

# ---------------------------------------------------------------------------
# bootstrap
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
