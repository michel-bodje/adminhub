import {
  ELEMENT_IDS,
  formState,
  showPage,
  resetPage,
  loadLawyers,
  setLawyers,
  getLawyerById,
  populateLawyerDropdown,
  handleCaseDetails,
  handlePaymentOptions,
  collectCaseDetails,
} from "./index.js";

window.addEventListener('pywebviewready', () => {
  console.log("[AdminHub] pywebview is ready.");
  main();
});

/** Main entry point for the application. */
async function main() {
  const lawyerList = await loadLawyers();
  if (lawyerList.length === 0) {
    console.error("No lawyers loaded");
    return;
  } else {
    console.log("[AdminHub] Lawyers loaded:", lawyerList);
  }
  setLawyers(lawyerList);
  resetPage();
  attachEventListeners();
}

/** Submits the current formState to the Python backend. */
async function submitForm() {
  const lawyer = getLawyerById(formState.lawyerId);
  const details = collectCaseDetails();
  await window.pywebview.api.submit_form(formState, lawyer, details);
}

/** Attaches all event listeners for the application. */
function attachEventListeners() {
  // Form input event listeners
  const formInputs = document.querySelectorAll('input, select, textarea');
  formInputs.forEach((input) => {
    input.addEventListener('change', (event) => {
      /**
       * Update form state
       * and activate events based on input id
       */
      const { id, value } = event.target;
      console.log("changed: %s", input.id); // debug

      // Find the matching ELEMENT_IDS key
      const matchingKey = Object.keys(ELEMENT_IDS).find((key) => ELEMENT_IDS[key] === id);
      console.log("matchingKey: %s", matchingKey); // debug

      if (matchingKey) {
        if (matchingKey.endsWith("LawyerId")) {
          // lawyer dropdown change
          formState.update("lawyerId", value);
        } else if (matchingKey.endsWith("Location")) {
          // location dropdown change
          formState.update("location", value);
        } else if (matchingKey.endsWith("CaseType")) {
          // case type dropdown change
          formState.update("caseType", value);
          populateLawyerDropdown();
          handleCaseDetails();
        } else if (matchingKey.endsWith("ClientName")) {
          // client name input change
          formState.update("clientName", value);
        } else if (matchingKey.endsWith("ClientPhone")) {
          // client phone input change
          formState.update("clientPhone", value);
        } else if (matchingKey.endsWith("ClientEmail")) {
          // client email input change
          formState.update("clientEmail", value);
        } else if (matchingKey.endsWith("ClientTitle")) {
          // client title dropdown change
          formState.update("clientTitle", value);
        } else if (matchingKey.endsWith("ClientLanguage")) {
          // client language dropdown change
          formState.update("clientLanguage", value);
        } else if (matchingKey.endsWith("scheduleMode")) {
          // appointment mode dropdown change
          const manualDate = document.getElementById(ELEMENT_IDS.manualDate);
          const manualTime = document.getElementById(ELEMENT_IDS.manualTime);
          const manualDateLabel = document.querySelector(`label[for=${ELEMENT_IDS.manualDate}]`);
          const manualTimeLabel = document.querySelector(`label[for=${ELEMENT_IDS.manualTime}]`);

          // Show/hide manual date/time inputs based on selected mode          
          if (value === "auto") {
            manualDateLabel.classList.add("hidden");
            manualTimeLabel.classList.add("hidden");
            manualDate.classList.add("hidden");
            manualTime.classList.add("hidden");
            manualDate.required = false;
            manualTime.required = false;
          } else if (value === "manual") {
            manualDateLabel.classList.remove("hidden");
            manualTimeLabel.classList.remove("hidden");
            manualDate.classList.remove("hidden");
            manualTime.classList.remove("hidden");
            manualDate.required = true;
            manualTime.required = true;
          }
        } else if (matchingKey.endsWith("Date")) {
          console.log("Date changed:", value);
          formState.update("appointmentDate", value);
        } else if (matchingKey.endsWith("Time")) {
          console.log("Time changed:", value);
          formState.update("appointmentTime", value); 
        } else if (matchingKey.endsWith("FirstConsultation")) {
          // first consultation checkbox change
          formState.update("isFirstConsultation", event.target.checked);
        } else if (matchingKey.endsWith("refBarreau")) {
          // ref barreau checkbox change
          formState.update("isRefBarreau", event.target.checked);
        } else if (matchingKey.endsWith("existingClient")) {
          // existing client checkbox change
          formState.update("isExistingClient", event.target.checked);
        } else if (matchingKey.endsWith("paymentMade")) {
          // payment checkbox change
          formState.update("isPaymentMade", event.target.checked);
          handlePaymentOptions();
        } else if (matchingKey.endsWith("PaymentMethod")) {
          // payment method dropdown change
          formState.update("paymentMethod", value);
        } else if (matchingKey.endsWith("notes")) {
          // schedule notes textarea change
          formState.update("notes", value);
        } else if (matchingKey.endsWith("Deposit")) {
          // contract deposit input change
          formState.update("depositAmount", value);
        } else if (matchingKey.endsWith("wordContractTitle")) {
          // Show/hide custom contract title input based on selection
          const customTitleInput = document.getElementById(ELEMENT_IDS.customContractTitle);
          if (value === "other") {
            customTitleInput.classList.remove("hidden");
            customTitleInput.required = true;
          } else {
            customTitleInput.classList.add("hidden");
            customTitleInput.required = false;
            formState.update("contractTitle", value);
          }
        } else if (matchingKey.endsWith("wordReceiptReason")) {
          // Show/hide custom receipt reason input based on selection
          const customReasonInput = document.getElementById(ELEMENT_IDS.customReceiptReason);
          if (value === "other") {
            customReasonInput.classList.remove("hidden");
            customReasonInput.required = true;
          } else {
            customReasonInput.classList.add("hidden");
            customReasonInput.required = false;
            formState.update("receiptReason", value);
        }
        } else if (matchingKey.endsWith("customContractTitle")) {
          // Update form state with custom contract title
          formState.update("contractTitle", value);
        } else if (matchingKey.endsWith("customReceiptReason")) {
          // Update form state with custom receipt reason
          formState.update("receiptReason", value);
        }
      }
    });
  });

  // Menu buttons event listeners
  const menuButtons = document.querySelectorAll('.menu-btn');
  menuButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      const { id } = event.target;
      switch (id) {
      case ELEMENT_IDS.outlookMenuBtn:
        showPage(ELEMENT_IDS.outlookMenuPage);
        break;
      case ELEMENT_IDS.wordMenuBtn:
        showPage(ELEMENT_IDS.wordMenuPage);
        break;
      case ELEMENT_IDS.pclawMenuBtn:
        showPage(ELEMENT_IDS.pclawMenuPage);
        break;
      case ELEMENT_IDS.scheduleMenuBtn:
        showPage(ELEMENT_IDS.schedulePage);
        break;
      case ELEMENT_IDS.confirmMenuBtn:
        showPage(ELEMENT_IDS.confirmationPage);
        break;
      case ELEMENT_IDS.contractMenuBtn:
        showPage(ELEMENT_IDS.contractPage);
        break;
      case ELEMENT_IDS.replyMenuBtn:
        showPage(ELEMENT_IDS.replyPage);
        break;
      case ELEMENT_IDS.followupMenuBtn:
        showPage(ELEMENT_IDS.followupPage);
        break;
      case ELEMENT_IDS.reviewMenuBtn:
        showPage(ELEMENT_IDS.reviewPage);
        break;
      case ELEMENT_IDS.wordContractMenuBtn:
        showPage(ELEMENT_IDS.wordContractPage);
        break;
      case ELEMENT_IDS.wordReceiptMenuBtn:
        showPage(ELEMENT_IDS.wordReceiptPage);
        break;
      case ELEMENT_IDS.pclawMatterMenuBtn:
        showPage(ELEMENT_IDS.pclawMatterPage);
        break;
      case ELEMENT_IDS.pclawCloseMatterMenuBtn:
        showPage(ELEMENT_IDS.pclawCloseMatterPage);
        break;
      case ELEMENT_IDS.pclawBillMatterMenuBtn:
        showPage(ELEMENT_IDS.pclawBillMatterPage);
        break;
      default:
        break;
      }
    });
  });

  // Back buttons event listeners
  const backButtons = document.querySelectorAll('.back-btn');
  backButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      event.preventDefault();
      resetPage();
    });
  });

  // Submit buttons event listeners
  const submitButtons = document.querySelectorAll('.submit-btn');
  submitButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      event.preventDefault();
      const { id } = event.target;
      switch (id) {
        case ELEMENT_IDS.scheduleSubmitBtn:
          scheduleAppointment();
          break;
        case ELEMENT_IDS.confirmSubmitBtn:
          sendConfirmation();
          break;
        case ELEMENT_IDS.contractSubmitBtn:
          sendContract();
          break;
        case ELEMENT_IDS.replySubmitBtn:
          sendReply();
          break;
        case ELEMENT_IDS.reviewSubmitBtn:
          sendReview();
          break;
        case ELEMENT_IDS.followupSubmitBtn:
          sendFollowup();
        case ELEMENT_IDS.wordContractSubmitBtn:
          createContract();
          break;
        case ELEMENT_IDS.wordReceiptSubmitBtn:
          createReceipt();
          break;
        case ELEMENT_IDS.pclawMatterSubmitBtn:
          newMatter();
        case ELEMENT_IDS.pclawCloseMatterSubmitBtn:
          closeMatter();
          break;
        case ELEMENT_IDS.pclawBillMatterSubmitBtn:
          billMatter();
          break;
        default:
          break;
      }
    });
  });
}

/** Schedules an appointment based on the form state. */
async function scheduleAppointment() {
  try {
    await submitForm();
    await window.pywebview.api.run("scheduler");
    console.log("[AdminHub] Appointment scheduled successfully.");

    // Chain actions without re-submitting the form
    const alsoEmail = document.getElementById("also-email").checked;
    const alsoPclaw = document.getElementById("also-pclaw").checked;

    if (alsoEmail) {
      await window.pywebview.api.run("emailConfirmation");
      console.log("[AdminHub] Confirmation email prepared and submitted.");
    }
    if (alsoPclaw) {
      await window.pywebview.api.run("new_matter");
      console.log("[AdminHub] PCLaw matter created successfully.");
    }
  } catch (error) {
    console.error("[AdminHub] Error scheduling appointment:", error);
    alert("Failed to schedule appointment. Please try again.");
    throw error;
  }
}

/** Prepares a confirmation email. */
async function sendConfirmation() {
  try {
    await submitForm();
    await window.pywebview.api.run("emailConfirmation");
    console.log("[AdminHub] Confirmation email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing confirmation email:", error);
    alert("Failed to prepare confirmation email. Please try again.");
    throw error;
  }
}

/** Prepares a reply email. */
async function sendReply() {
  try {
    await submitForm();
    await window.pywebview.api.run("emailReply");
    console.log("[AdminHub] Reply email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing reply email:", error);
    alert("Failed to prepare reply email. Please try again.");
    throw error;
  }
}

/** Prepares a review request email. */
async function sendReview() {
  try {
    await submitForm();
    await window.pywebview.api.run("emailReview");
    console.log("[AdminHub] Review request email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing review request email:", error);
    alert("Failed to prepare review request email. Please try again.");
    throw error;
  }
}

/** Prepares a follow-up email. */
async function sendFollowup() {
  try {
    await submitForm();
    await window.pywebview.api.run("emailSuivi");
    console.log("[AdminHub] Follow-up email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing follow-up email:", error);
    alert("Failed to prepare follow-up email. Please try again.");
    throw error;
  }
}

/** Prepares a contract email. */
async function sendContract() {
  try {
    await submitForm();
    await window.pywebview.api.run("emailContract");
    console.log("[AdminHub] Contract email prepared and submitted.");
  } catch (error) {
    console.error("[AdminHub] Error preparing contract email:", error);
    alert("Failed to prepare contract email. Please try again.");
    throw error;
  }
}

/** Create a word contract based on the form state. */
async function createContract() {
  try {
    await submitForm();
    await window.pywebview.api.run("wordContract");
    console.log("[AdminHub] Word contract created successfully.");
    
    const alsoContract = document.getElementById("also-contract").checked;
    if (alsoContract) {
      // Automatically prepare and submit the contract email
      await window.pywebview.api.run("emailContract");
      console.log("[AdminHub] Contract email prepared and submitted.");
    }
  } catch (error) {
    console.error("[AdminHub] Error creating word contract:", error);
    alert("Failed to create contract doc or email. Please try again.");
    throw error;
  }
}

/** Create a word receipt based on the form state. */
async function createReceipt() {
  try {
    await submitForm();
    await window.pywebview.api.run("wordReceipt");
    console.log("[AdminHub] Word receipt created successfully.");
  } catch (error) {
    console.error("[AdminHub] Error creating word receipt:", error);
    alert("Failed to create receipt doc. Please try again.");
    throw error;
  }
}

// Create PCLaw matter based on the form state
async function newMatter() {
  try {
    await submitForm();
    await window.pywebview.api.run("new_matter");
    console.log("[AdminHub] PCLaw matter created successfully.");
  } catch (error) {
    console.error("[AdminHub] Error writing PCLaw matter", error);
    alert("Failed to write new PCLaw matter. Please try again.");
    throw error;
  }
}

async function closeMatter() {
  try {
    const matterId = document.getElementById(ELEMENT_IDS.closeMatterId).value;
    await window.pywebview.api.run("close_matter", matterId);
    console.log("[AdminHub] PCLaw matter closed successfully.");
  } catch (error) {
    console.error("[AdminHub] Error closing PCLaw matter", error);
    alert("Failed to close PCLaw matter. Please try again.");
    throw error;
  }
}

async function billMatter() {
  try {
    const matterId = document.getElementById(ELEMENT_IDS.billMatterId).value;
    await window.pywebview.api.run("bill_matter", matterId);
    console.log("[AdminHub] PCLaw matter billed successfully.");
  } catch (error) {
    console.error("[AdminHub] Error billing PCLaw matter", error);
    alert("Failed to bill PCLaw matter. Please try again.");
    throw error;
  }
}