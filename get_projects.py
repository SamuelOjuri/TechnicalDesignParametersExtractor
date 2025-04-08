import os
from dotenv import load_dotenv
from monday_dot_com_interface import MondayDotComInterface
import json

# Load environment variables from .env file
load_dotenv()

# Get API token from environment variables
api_token = os.environ.get("MONDAY_API_TOKEN")

# Check if the token was loaded
if not api_token:
    print("Error: MONDAY_API_TOKEN not found in environment variables.")
    exit(1)

monday_interface = MondayDotComInterface(api_token)

# When displaying project information, extract the title from the column values.
def extract_project_title(column_values):
    project_title = None
    for col in column_values:
        # First check if it's a MirrorValue and if display_value is present.
        if col.get("__typename") == "MirrorValue" and col.get("display_value"):
            project_title = col["display_value"]
        # Otherwise, use the text value if available.
        elif col.get("text"):
            project_title = col["text"]
        if project_title:
            break
    return project_title

# Call the function to get all projects from the Tapered Enquiry Maintenance board
projects, error = monday_interface.get_tapered_enquiry_projects()

# Check if projects were returned successfully
if projects:
    print(f"Retrieved {len(projects)} projects from Tapered Enquiry Maintenance board:")
    
    # Display information about the first 10 projects only
    for i, project in enumerate(projects[:10], 1):
        project_id = project["id"]
        project_name = project["name"]
        # Use the helper function to extract the project title
        project_title = extract_project_title(project.get("column_values", []))
        print(f"{i}. ID: {project_id} - Name: {project_name} - Title: {project_title}")
    
    # Example of how to use the first project in detail
    if projects:
        first_project = projects[0]
        print(f"\nSelected Project: {first_project['name']} (ID: {first_project['id']})")
        
        board_id = "1825117125"  # Board ID for Tapered Enquiry Maintenance
        
        project_details, error = monday_interface.get_item_by_name_on_board(board_id, first_project['name'])
        
        if project_details:
            print("\nDetailed information about the selected project:")
            print(f"ID: {project_details['id']}")
            print(f"Name: {project_details['name']}")
            # Extract the project title with MirrorValue support
            project_title = extract_project_title(project_details.get("column_values", []))
            print(f"Title: {project_title}")
            # (Rest of your logic remains unchanged)
        else:
            print(f"Error retrieving project details: {error}")
else:
    print(f"Error retrieving projects: {error}")
