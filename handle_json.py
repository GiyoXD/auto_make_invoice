import json

def load_json_file_to_map(file_path: str) -> dict | None:
    """
    Loads JSON data from a file into a Python dictionary (map).

    Assumes the root element of the JSON file is an object (`{...}`).
    Preserves the nested structure of the JSON data.

    Args:
        file_path (str): The path to the JSON file ("test.json").

    Returns:
        dict: A Python dictionary representing the JSON structure.
        None: If the file is not found, cannot be read, contains invalid JSON,
              or the root JSON element is not an object (map).
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        if isinstance(data, dict):
            return data
        else:
            print(f"Error: Root element in '{file_path}' is not an object (map).")
            return None

    except FileNotFoundError:
        print(f"Error: File not found at '{file_path}'")
        return None
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON from file '{file_path}': {e}")
        return None
    except IOError as e:
        print(f"Error reading file '{file_path}': {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

if __name__ == "__main__":
    file_path = "test.json"
    loaded_map = load_json_file_to_map(file_path)
    print(loaded_map)

    if loaded_map:
        print("JSON loaded successfully.")
    else:
        print("Failed to load JSON.")