from pclaw import *
from parse_json import *

def process_bill_matter(data):
    """
    Process matter billing with the provided data.
    Args:
        data (dict): The parsed JSON data containing matter information
    Returns:
        dict: Result of the billing process
    """
    try:
        app = connect_to_pclaw()
        app.set_focus()
        matter = get_matter(data)
        sleep(1)

        bill_matter(app, matter)
        
        return {
            "status": "success",
            "message": f"Successfully billed matter {matter}",
            "matter_id": matter
        }
        
    except Exception as e:
        log(f"Error billing matter: {str(e)}")
        return {"error": str(e)}

def main():
    """Backward compatibility for standalone execution"""
    try:
        data = read_json()
        result = process_bill_matter(data)
        if result.get("error"):
            print(f"Error: {result['error']}")
        else:
            print(result.get("message", "Matter billed successfully"))
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()