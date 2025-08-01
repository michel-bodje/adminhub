# module_registry.py
from config import log

# Import all your converted modules
from emailConfirmation import process_email_confirmation
from emailContract import process_email_contract
from emailFollowup import process_email_followup
from emailReply import process_email_reply
from emailReview import process_email_review
from scheduler import process_scheduler
from new_matter import process_new_matter
from close_matter import process_close_matter
from bill_matter import process_bill_matter
from wordContract import process_word_contract
from wordReceipt import process_word_receipt
from time_entries import process_time_entries

MODULE_REGISTRY = {
    "emailConfirmation": process_email_confirmation,
    "emailContract": process_email_contract,
    "emailFollowup": process_email_followup,
    "emailReply": process_email_reply,
    "emailReview": process_email_review,
    "scheduler": process_scheduler,
    "new_matter": process_new_matter,
    "close_matter": process_close_matter,
    "bill_matter": process_bill_matter,
    #"wordContract": process_word_contract,
    #"wordReceipt": process_word_receipt,
    "time_entries": process_time_entries,
}

def run_module(script_name, json_data):
    """
    Run a module function directly instead of as a subprocess.
    """
    log(f"Running module: {script_name}")
    
    if script_name not in MODULE_REGISTRY:
        return {"error": f"Module not found: {script_name}"}
    
    try:
        module_function = MODULE_REGISTRY[script_name]
        result = module_function(json_data)
        return result
        
    except Exception as e:
        log(f"Error running module {script_name}: {str(e)}")
        return {"error": str(e)}