import requests
import json
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional


class MondayDotComInterface:
    def __init__(self, api_token: str):
        self.api_token = api_token
        self.api_url = "https://api.monday.com/v2/"

    def does_item_exist(self, board_id: str, tp_number: str, revision: str) -> Tuple[bool, str]:
        """
        Checks if an item exists on a board with the given TP number and revision.
        
        Returns a tuple of (exists, error_message)
        """
        null_error = ""
        item = self.get_item_by_name_on_board(board_id, f"{tp_number}_{revision}")
        
        if item is None:
            return False, null_error
        
        return True, ""


    def send_query_to_monday(self, query: str) -> Dict[str, Any]:
        """
        Sends a GraphQL query to the Monday.com API.
        
        Returns the parsed JSON response.
        """
        try:
            headers = {
                "Authorization": self.api_token,
                "Content-Type": "application/json"
            }
            
            response = requests.post(
                self.api_url,
                headers=headers,
                json=json.loads(query),
                timeout=30
            )
            
            return response.json()
        except Exception:
            return None

    def get_id_by_item_name(self, item: str, board_name: str) -> Tuple[Optional[str], str]:
        """
        Gets the ID of an item by its name and board name.
        
        Returns a tuple of (item_id, error_message)
        """
        active_boards_root = self.send_query_to_monday('{"query":"query Boards{boards(state: active, limit: 100000){name id}}"}')
        
        if active_boards_root is not None:
            matching_boards = [b for b in active_boards_root["data"]["boards"] if b["name"] == board_name]
            
            if not matching_boards:
                return None, f"{board_name} Board not found in Monday.com"
            
            items = []
            
            for board in matching_boards:
                query = (
                    '{"query":"{items_page_by_column_values(board_id: '
                    + board["id"]
                    + 'columns: [{ column_id: \\\"name\\\", column_values: [\\\"' + item + '\\\"] }]) '
                    + '{ items { id name subitems {id name column_values {id text type value}}}}}"}'
                )
                found_items = self.send_query_to_monday(query)
                
                if found_items["data"] is None:
                    return None, "Item not found in Monday.com"
                
                if found_items["data"]["items_page_by_column_values"] is not None:
                    items.extend(found_items["data"]["items_page_by_column_values"]["items"])
            
            if len(items) == 1:
                return items[0]["id"], ""
            elif len(items) == 0:
                return None, "Item not found in Monday.com"
            else:
                return None, "Duplicate Item found in Monday.com"
        else:
            return None, "Unable to retrieve Boards from Monday.com"

    def get_contacts_list(self) -> Tuple[Optional[List[Dict[str, Any]]], str]:
        """
        Gets a list of all contacts from the Contacts board.
        
        Returns a tuple of (contacts_list, error_message)
        """
        active_boards_root = self.send_query_to_monday('{"query":"query Boards{boards(state: active, limit: 100000){name id}}"}')
        
        if active_boards_root is not None:
            account_board_id = ""
            account_boards = [b for b in active_boards_root["data"]["boards"] if b["name"] == "Accounts"]
            
            if not account_boards:
                return None, "Accounts Board not found in Monday.com"
            else:
                account_board_id = account_boards[0]["id"]
            
            contact_boards = [b for b in active_boards_root["data"]["boards"] if b["name"] == "Contacts"]
            if not contact_boards:
                return None, "Contacts Board not found in Monday.com"
            
            items = []
            
            for board in contact_boards:
                query = (
                    '{"query":"query next_items_page {boards(ids: ' + board["id"] + 
                    ', limit: 100000) {id name items_page(limit: 500, cursor: null) {cursor items {id name state ' + 
                    'linked_items(link_to_item_column_id: \\\"contact_account\\\" linked_board_id: ' + account_board_id + 
                    ') {id name}}} items_count}}"}'
                )
                found_items = self.send_query_to_monday(query)
                
                if found_items["data"] is None:
                    return None, "Item not found in Monday.com"
                
                if (found_items["data"]["boards"] is not None and 
                    found_items["data"]["boards"][0]["items_page"] is not None):
                    if found_items["data"]["boards"][0]["items_page"]["items"] is not None:
                        active_items = [item for item in found_items["data"]["boards"][0]["items_page"]["items"] 
                                        if item["state"] == "active"]
                        items.extend(active_items)
                    
                    cursor = found_items["data"]["boards"][0]["items_page"]["cursor"]
                    while cursor is not None:
                        next_query = (
                            '{"query":"query next_items_page {boards(ids: ' + board["id"] + 
                            ', limit: 100000) {id name items_page(limit: 500, cursor: \\\"' + cursor + 
                            '\\\") {cursor items {id name state ' + 
                            'linked_items(link_to_item_column_id: \\\"contact_account\\\" linked_board_id: ' + account_board_id + 
                            ') {id name}}} items_count}}"}'
                        )
                        found_items = self.send_query_to_monday(next_query)
                        if found_items["data"]["boards"][0]["items_page"]["items"] is not None:
                            active_items = [item for item in found_items["data"]["boards"][0]["items_page"]["items"] 
                                            if item["state"] == "active"]
                            items.extend(active_items)
                        cursor = found_items["data"]["boards"][0]["items_page"]["cursor"]
            
            if not items:
                return None, "Item not found in Monday.com"
            else:
                return items, ""
        else:
            return None, "Unable to retrieve Boards from Monday.com"

    def get_companies_list(self) -> Tuple[Optional[List[Dict[str, Any]]], str]:
        """
        Gets a list of all companies from the Accounts board.
        
        Returns a tuple of (companies_list, error_message)
        """
        active_boards_root = self.send_query_to_monday('{"query":"query Boards{boards(state: active, limit: 100000){name id}}"}')
        
        if active_boards_root is not None:
            account_boards = [b for b in active_boards_root["data"]["boards"] if b["name"] == "Accounts"]
            if not account_boards:
                return None, "Accounts Board not found in Monday.com"
            
            items = []
            
            for board in account_boards:
                query = (
                    '{"query":"query next_items_page {boards(ids: ' + board["id"] + 
                    ', limit: 100000) {id name items_page(limit: 500, cursor: null) {cursor items {id name state}} items_count}}"}'
                )
                found_items = self.send_query_to_monday(query)
                
                if found_items["data"] is None:
                    return None, "Item not found in Monday.com"
                
                if (found_items["data"]["boards"] is not None and 
                    found_items["data"]["boards"][0]["items_page"] is not None):
                    if found_items["data"]["boards"][0]["items_page"]["items"] is not None:
                        active_items = [item for item in found_items["data"]["boards"][0]["items_page"]["items"] 
                                        if item["state"] == "active"]
                        items.extend(active_items)
                    
                    cursor = found_items["data"]["boards"][0]["items_page"]["cursor"]
                    while cursor is not None:
                        next_query = (
                            '{"query":"query next_items_page {boards(ids: ' + board["id"] + 
                            ', limit: 100000) {id name items_page(limit: 500, cursor: \\\"' + cursor + 
                            '\\\") {cursor items {id name state}} items_count}}"}'
                        )
                        found_items = self.send_query_to_monday(next_query)
                        if found_items["data"]["boards"][0]["items_page"]["items"] is not None:
                            active_items = [item for item in found_items["data"]["boards"][0]["items_page"]["items"] 
                                            if item["state"] == "active"]
                            items.extend(active_items)
                        cursor = found_items["data"]["boards"][0]["items_page"]["cursor"]
            
            if not items:
                return None, "Item not found in Monday.com"
            else:
                return items, ""
        else:
            return None, "Unable to retrieve Boards from Monday.com"

    def get_users_list(self) -> Tuple[Optional[List[Dict[str, Any]]], str]:
        """
        Gets a list of all users from Monday.com.
        
        Returns a tuple of (users_list, error_message)
        """
        user_root = self.send_query_to_monday('{"query":"query Users { users {name id email url}}"}')
        
        if user_root is None or not user_root["data"]["users"]:
            return None, "Users not found in Monday.com"
        else:
            return user_root["data"]["users"], ""

    def get_board_id_for_item(self, item_name: str, board_name: str) -> Tuple[Optional[str], str]:
        """
        Gets the board ID for a specific item and board name.
        
        Returns a tuple of (board_id, error_message)
        """
        active_boards_root = self.send_query_to_monday('{"query":"query Boards{boards(state: active, limit: 100000){name id}}"}')
        
        if (active_boards_root is None or active_boards_root["data"] is None or 
            active_boards_root["data"]["boards"] is None or not active_boards_root["data"]["boards"]):
            return None, "Unable to Query Boards on Monday.com"
        else:
            matching_boards = [b for b in active_boards_root["data"]["boards"] if b["name"] == board_name]
            if matching_boards:
                board = matching_boards[0]
                return board["id"], ""
            else:
                return None, f"Unable to find board Name:{board_name} on Monday.com"

    def get_board_by_id(self, board_id: str) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Gets a board by its ID.
        
        Returns a tuple of (board, error_message)
        """
        active_boards_root = self.send_query_to_monday(
            '{"query":"query Boards{boards(state: active, ids: [' + board_id + ']){name id columns {id settings_str title type}}}"}'
        )
        
        if (active_boards_root is None or active_boards_root["data"] is None or 
            active_boards_root["data"]["boards"] is None or not active_boards_root["data"]["boards"]):
            return None, f"Unable to find board Id:{board_id} on Monday.com"
        else:
            return active_boards_root["data"]["boards"][0], ""

    def get_item_by_column_values(self, board_id: str, search_column_id: str, 
                                 search_value: str, value_column_id: str) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Gets items that match specific column values.
        
        Returns a tuple of (items_page, error_message)
        """
        query = (
            '{"query":"{items_page_by_column_values(board_id: '
            + board_id
            + ' columns: [{ column_id: \\\"' + search_column_id + '\\\", column_values: [\\\"' + search_value + '\\\"] }]) '
            + '{ items { id column_values(ids: \\\"' + value_column_id + '\\\") {value}}}}"}'
        )
        found_items = self.send_query_to_monday(query)
        
        if found_items["data"]["items_page_by_column_values"]["items"]:
            return found_items["data"]["items_page_by_column_values"], ""
        else:
            return None, "No items found in Monday.com"

    def get_item_by_name_on_board(self, board_id: str, name: str) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Gets an item by its name on a specific board.
        
        Returns a tuple of (item, error_message)
        """
        items = []
        
        query = (
            '{"query": '
            '"{ items_page_by_column_values(board_id: ' + board_id + ', columns: ['
            '{ column_id: \\"name\\", column_values: [\\"' + name + '\\"] }'
            ']) { items { id name assets { id name } '
            'column_values { id text __typename ... on MirrorValue { display_value } } '
            'subitems { id name column_values { id text __typename ... on MirrorValue { display_value } } } '
            '} } }"'
            '}'
        )
        
        print(f"DEBUG: Monday query: {query}")
        found_items = self.send_query_to_monday(query)
        print(f"DEBUG: Monday API response: {found_items}")
        
        # Check if a valid response was returned
        if not found_items or "data" not in found_items:
            error_message = "No data returned from Monday.com."
            if found_items and "errors" in found_items:
                error_message += " Errors: " + str(found_items["errors"])
            return None, error_message

        # Process the items if present
        if found_items["data"].get("items_page_by_column_values") is not None:
            items.extend(found_items["data"]["items_page_by_column_values"].get("items", []))
        
        print(f"DEBUG: Found {len(items)} items")
        if len(items) == 1:
            return items[0], ""
        elif len(items) == 0:
            return None, "Item not found in Monday.com"
        else:
            return items[0], "Multiple items found in Monday.com. Using the first match."

    @staticmethod
    def _build_items_page_query(board_id: str,
                                start_date: str,
                                cursor: Optional[str] = None) -> str:
        """
        Helper – returns the exact JSON string you must POST to Monday.com.
        If `cursor` is provided we're building a pagination request, otherwise
        it is the first page.
        """
        # ------- GraphQL block (clean, human‑readable) -------------------------
        gql = f"""
        query {{
        boards(ids: {board_id}) {{
            items_page(
            limit: 500,
            {f'cursor: "{cursor}",' if cursor else ''}
            query_params: {{
                rules: [{{
                column_id: "date9__1",
                compare_value: ["EXACT", "{start_date}"],
                operator: greater_than_or_equals
                }}]
            }}
            ) {{
            cursor
            items {{
                id
                name
                state
                column_values(ids: ["text3__1", "date9__1"]) {{
                id
                text
                __typename
                ... on MirrorValue {{
                    display_value
                }}
                }}
            }}
            }}
        }}
        }}
        """
        # The Monday endpoint expects a JSON payload *whose single member*
        # is the GraphQL string.
        return json.dumps({"query": gql})
    

    def get_tapered_enquiry_projects(
            self,
            start_date: str = "2021-01-01"
    ) -> Tuple[Optional[List[Dict[str, Any]]], str]:
        """
        Return active projects whose **Created** date (column `date9__1`)
        is **on or after** `start_date` (YYYY‑MM‑DD) from the
        "Tapered Enquiry Maintenance" board.
        """
        board_id = "1825117125"
        items: List[Dict[str, Any]] = []

        # ---- first page -------------------------------------------------------
        payload = self._build_items_page_query(board_id, start_date)
        response = self.send_query_to_monday(payload)

        if response is None:
            return None, "No response from Monday.com"

        if "errors" in response:
            return None, f"Monday.com API error: {response['errors']}"

        try:
            page = response["data"]["boards"][0]["items_page"]
        except (KeyError, IndexError, TypeError):
            return None, "Unexpected payload structure from Monday.com"

        # ---- collect items & paginate ----------------------------------------
        def _append_active(page_obj: Dict[str, Any]) -> None:
            for itm in page_obj.get("items", []):
                if itm.get("state") == "active":
                    items.append(itm)

        _append_active(page)
        cursor = page.get("cursor")

        while cursor:
            payload = self._build_items_page_query(board_id, start_date, cursor)
            response = self.send_query_to_monday(payload)

            # More thorough error checking for nested values
            if response is None or "data" not in response:
                break
            
            # Check if boards exists and is not empty
            if not response["data"].get("boards"):
                break
                
            # Access the first board safely
            board = response["data"]["boards"][0] if response["data"]["boards"] else None
            if not board or "items_page" not in board:
                break
                
            page = board["items_page"]
            _append_active(page)
            cursor = page.get("cursor")

        if not items:
            return None, (
                f"No active projects created on or after {start_date} "
                "were found on the Tapered Enquiry Maintenance board."
            )
        return items, ""

    def get_project_by_title(self, board_id: str, project_title: str, limit: int = 10) -> Tuple[Optional[Dict[str, Any]], str]:
        """
        Gets a project by its title (from the "Project Name" column) on a specific board.
        
        Args:
            board_id: The ID of the board to search
            project_title: The title to search for in the "Project Name" column (text3__1)
            limit: Maximum number of results to return
            
        Returns:
            A tuple of (project_details, error_message)
        """
        query = (
            '{"query": "query {'
            'boards(ids: ' + board_id + ') {'
            'items_page('
            'limit: ' + str(limit) + ','
            'query_params: {'
            'rules: ['
            '{'
            'column_id: \\"text3__1\\",'
            'compare_value: [\\"' + project_title + '\\"],'
            'operator: contains_text'
            '}'
            ']'
            '}'
            ') {'
            'items {'
            'id'
            'name'
            'column_values {'
            'id'
            '__typename'
            'column {'
            'title'
            '}'
            'text'
            '... on MirrorValue {'
            'display_value'
            '}'
            '}'
            'subitems {'
            'id'
            'name'
            'column_values {'
            '__typename'
            'column {'
            'title'
            '}'
            '... on MirrorValue {'
            'display_value'
            '}'
            'text'
            '}'
            '}'
            '}'
            '}'
            '}'
            '"}'
        )
        
        found_items = self.send_query_to_monday(query)
        
        # Check if a valid response was returned
        if not found_items or "data" not in found_items:
            error_message = "No data returned from Monday.com."
            if found_items and "errors" in found_items:
                error_message += " Errors: " + str(found_items["errors"])
            return None, error_message

        # Process the items if present
        if (found_items["data"].get("boards") and 
            found_items["data"]["boards"][0].get("items_page") and 
            found_items["data"]["boards"][0]["items_page"].get("items")):
            
            items = found_items["data"]["boards"][0]["items_page"]["items"]
            
            if len(items) == 1:
                return items[0], ""
            elif len(items) == 0:
                return None, f"Project with title '{project_title}' not found"
            else:
                # If multiple items were found, return the first one with a warning
                return items[0], f"Multiple projects found with title '{project_title}'. Returning the first match."
        else:
            return None, f"Project with title '{project_title}' not found"
        
    def check_project_exists(self, sample_project_name: str, similarity_threshold: float = 0.55) -> Dict[str, Any]:
        """
        Checks if a project name exists or is similar to any project in the Tapered Enquiry Maintenance board.
        Uses fuzzy string matching to find similar project names.
        
        Args:
            sample_project_name: The project name to search for
            similarity_threshold: Threshold for string similarity (0.0 to 1.0, where 1.0 is exact match)
            
        Returns:
            A dictionary with:
                - 'exists': Boolean indicating if project exists
                - 'type': 'new' or 'existing'
                - 'matches': List of potential matches with similarity scores
                - 'best_match': The best match if any
                - 'similarity_score': Score of best match
                - 'error': Error message if any
        """
        # Initialize result dictionary
        result = {
            'exists': False, 
            'type': 'new',
            'matches': [],
            'best_match': None,
            'similarity_score': 0.0,
            'error': ''
        }
        
        # Get all projects from Monday.com
        projects, error = self.get_tapered_enquiry_projects()
        
        if error:
            result['error'] = error
            return result
        
        if not projects:
            result['error'] = "No projects found to compare against"
            return result
        
        # Helper function for string similarity using Levenshtein distance
        def similarity(s1: str, s2: str) -> float:
            """Calculate string similarity between 0.0 and 1.0"""
            if not s1 or not s2:
                return 0.0
            
            # Convert both strings to lowercase for case-insensitive comparison
            s1, s2 = s1.lower(), s2.lower()
            
            # Calculate Levenshtein distance
            if len(s1) < len(s2):
                s1, s2 = s2, s1
            
            distances = range(len(s2) + 1)
            for i, c1 in enumerate(s1):
                distances_ = [i + 1]
                for j, c2 in enumerate(s2):
                    if c1 == c2:
                        distances_.append(distances[j])
                    else:
                        distances_.append(1 + min((distances[j], distances[j + 1], distances_[-1])))
                distances = distances_
            
            # Calculate similarity as 1 - normalized_distance
            max_len = max(len(s1), len(s2))
            if max_len == 0:
                return 1.0  # Both strings are empty
            
            return 1 - (distances[-1] / max_len)
        
        # Helper function to extract project title from column values
        def extract_project_title(column_values):
            project_title = None
            for col in column_values:
                # First check if it's a MirrorValue and if display_value is present
                if col.get("__typename") == "MirrorValue" and col.get("display_value"):
                    project_title = col["display_value"]
                # Otherwise, use the text value if available
                elif col.get("text"):
                    project_title = col["text"]
                if project_title:
                    break
            return project_title
        
        # Find similar projects
        matches = []
        for project in projects:
            project_id = project["id"]
            project_name = project["name"]
            
            # Extract the project title from column values
            project_title = extract_project_title(project.get("column_values", []))
            
            # If no title available, use the name as fallback
            if not project_title:
                project_title = project_name
            
            # Calculate similarity score
            sim_score = similarity(sample_project_name, project_title)
            
            if sim_score >= similarity_threshold:
                matches.append({
                    'id': project_id,
                    'name': project_name,
                    'title': project_title,
                    'similarity': sim_score
                })
        
        # Sort matches by similarity score (highest first)
        matches = sorted(matches, key=lambda x: x['similarity'], reverse=True)
        
        # Update result dictionary
        result['matches'] = matches
        
        if matches:
            best_match = matches[0]
            result['exists'] = True
            result['type'] = 'existing'
            result['best_match'] = best_match
            result['similarity_score'] = best_match['similarity']
        
        return result
