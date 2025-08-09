import {
  formState,
  getLawyerById,
  collectCaseDetails,
  isValidEmail,
  isValidPhoneNumber,
  formatPhoneNumber,
  fileState,
  ELEMENT_IDS,
} from "./index.js";

/** Action handlers for various form submissions. */
export const ACTIONS_BY_ID = {
  [ELEMENT_IDS.scheduleSubmitBtn]: scheduleAppointment,
  [ELEMENT_IDS.confirmSubmitBtn]: sendConfirmation,
  [ELEMENT_IDS.replySubmitBtn]: sendReply,
  [ELEMENT_IDS.reviewSubmitBtn]: sendReview,
  [ELEMENT_IDS.followupSubmitBtn]: sendFollowup,
  [ELEMENT_IDS.contractSubmitBtn]: sendContract,
  [ELEMENT_IDS.wordContractSubmitBtn]: createContract,
  [ELEMENT_IDS.wordReceiptSubmitBtn]: createReceipt,
  [ELEMENT_IDS.pclawMatterSubmitBtn]: newMatter,
  [ELEMENT_IDS.pclawCloseMatterSubmitBtn]: closeMatter,
  [ELEMENT_IDS.pclawBillMatterSubmitBtn]: billMatter,
  [ELEMENT_IDS.timeEntriesSubmitBtn]: processTimeEntries,
};

/** Submits the current formState to the Python backend. */
async function getForm() {
  const lawyer = getLawyerById(formState.lawyerId);
  const details = collectCaseDetails();

  if (formState.clientEmail && !isValidEmail(formState.clientEmail)) {
    console.error("[AdminHub] Invalid client email address:", formState.clientEmail);
    throw new Error("Invalid client email address.");
  }
  
  formState.clientPhone = formatPhoneNumber(formState.clientPhone);

  if (formState.clientPhone && !isValidPhoneNumber(formState.clientPhone)) {
    console.error("[AdminHub] Invalid client phone number:", formState.clientPhone);
    throw new Error("Invalid client phone number.");
  }

  let response = await window.pywebview.api.format_form(formState, details, lawyer);
  console.log("[AdminHub] Form data prepared for submission:", response.json);
  return response.json;
}

/** Schedules an appointment based on the form state. */
export async function scheduleAppointment() {
  try {
    const json_data = await getForm();
    const result = await window.pywebview.api.run("scheduler", json_data);
    console.log("[AdminHub] Appointment scheduled successfully.");

    // Chain actions without re-submitting the form
    const alsoEmail = document.getElementById("also-email").checked;
    const alsoPclaw = document.getElementById("also-pclaw").checked;

    if (alsoEmail) {
      let formData = JSON.parse(json_data);
      formData.form.teamsMeeting = result.teamsMeeting;
      const emailJsonBlob = JSON.stringify(formData, null, 2);

      await window.pywebview.api.run("emailConfirmation", emailJsonBlob);
      console.log("[AdminHub] Confirmation email prepared and submitted.");
    }
    if (alsoPclaw) {
      await window.pywebview.api.run("new_matter", json_data);
      console.log("[AdminHub] PCLaw matter created successfully.");
    }
  } catch (error) {
    console.error("[AdminHub] Error scheduling appointment:", error);
    alert("Failed to schedule appointment. Please try again.");
    throw error;
  }
}

/** Prepares a confirmation email. */
export async function sendConfirmation() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("emailConfirmation", json_data);
    console.log("[AdminHub] Confirmation email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing confirmation email:", error);
    alert("Failed to prepare confirmation email. Please try again.");
    throw error;
  }
}

/** Prepares a reply email. */
export async function sendReply() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("emailReply", json_data);
    console.log("[AdminHub] Reply email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing reply email:", error);
    alert("Failed to prepare reply email. Please try again.");
    throw error;
  }
}

/** Prepares a review request email. */
export async function sendReview() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("emailReview", json_data);
    console.log("[AdminHub] Review request email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing review request email:", error);
    alert("Failed to prepare review request email. Please try again.");
    throw error;
  }
}

/** Prepares a follow-up email. */
export async function sendFollowup() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("emailFollowup", json_data);
    console.log("[AdminHub] Follow-up email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing follow-up email:", error);
    alert("Failed to prepare follow-up email. Please try again.");
    throw error;
  }
}

/** Prepares a contract email. */
export async function sendContract() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("emailContract", json_data);
    console.log("[AdminHub] Contract email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing contract email:", error);
    alert("Failed to prepare contract email. Please try again.");
    throw error;
  }
}

/** Create a word contract based on the form state. */
export async function createContract() {
  try {
    const json_data = await getForm();
    
    // Create the Word contract
    const result = await window.pywebview.api.run("wordContract", json_data);
    
    if (result.error) {
      throw new Error(result.error);
    }
    
    console.log("[AdminHub] Word contract created successfully.");
    console.log("[AdminHub] Result:", result);
    
    // Check if user wants to also create email
    const alsoContract = document.getElementById("also-contract").checked;
    if (alsoContract && result.pdf_path) {
      
      // Add PDF path to form data for email
      const formData = JSON.parse(json_data);
      formData.form.pdfPath = result.pdf_path;
      const emailJsonBlob = JSON.stringify(formData, null, 2);
      
      // Create contract email with PDF attachment
      await window.pywebview.api.run("emailContract", emailJsonBlob);
      console.log("[AdminHub] Contract email created with PDF attachment.");
    }
    
  } catch (error) {
    console.error("[AdminHub] Error creating word contract:", error);
    alert("Failed to create contract. Please try again.");
    throw error;
  }
}

/** Create a word receipt based on the form state. */
export async function createReceipt() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("wordReceipt", json_data);
    console.log("[AdminHub] Word receipt created successfully.");
  } catch (error) {
    console.error("[AdminHub] Error creating word receipt:", error);
    alert("Failed to create receipt doc. Please try again.");
    throw error;
  }
}

/* Create PCLaw matter based on the form state */
export async function newMatter() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("new_matter", json_data);
    console.log("[AdminHub] PCLaw matter created successfully.");
  } catch (error) {
    console.error("[AdminHub] Error writing PCLaw matter", error);
    alert("Failed to write new PCLaw matter. Please try again.");
    throw error;
  }
}

/** Close the specified PCLaw matter */
export async function closeMatter() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("close_matter", json_data);
    console.log("[AdminHub] PCLaw matter closed successfully.");
  } catch (error) {
    console.error("[AdminHub] Error closing PCLaw matter", error);
    alert("Failed to close PCLaw matter. Please try again.");
    throw error;
  }
}

/** Bill the specified PCLaw matter */
export async function billMatter() {
  try {
    const json_data = await getForm();
    await window.pywebview.api.run("bill_matter", json_data);
    console.log("[AdminHub] PCLaw matter billed successfully.");
  } catch (error) {
    console.error("[AdminHub] Error billing PCLaw matter", error);
    alert("Failed to bill PCLaw matter. Please try again.");
    throw error;
  }
}

export async function processTimeEntries() {
  try {
    const lawyerId = document.getElementById(ELEMENT_IDS.timeEntriesLawyer).value;
    
    if (!lawyerId) {
      throw new Error("Please select a lawyer.");
    }
    if (!fileState.selectedFilePath) {
      throw new Error("Please select a timesheet file.");
    }
    
    const formData = {
      lawyerId: lawyerId,
      filePath: fileState.selectedFilePath,
    };
    
    const json_data = JSON.stringify(formData);

    console.log("[AdminHub] Processing time entries with data:", json_data);
    
    await window.pywebview.api.run("time_entries", json_data);
    console.log("[AdminHub] Time entries processed successfully.");
    
  } catch (error) {
    console.error("[AdminHub] Error processing time entries:", error);
    throw error;
  }
}