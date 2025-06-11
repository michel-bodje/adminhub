import {
  ELEMENT_IDS,
  formState,
  showPage,
  resetPage,
  populateLanguageDropdown,
  populatePaymentDropdown,
  populateLawyerDropdown,
  populateContractTitles,
  handleCaseDetails,
  handlePaymentOptions,
  populateCaseTypeDropdown,
} from "./index.js";

/** Main entry point for the application. */
function main() {
  populateCaseTypeDropdown();
  populateLawyerDropdown();

  populateLocationDropdown();
  populateLanguageDropdown();
  
  populatePaymentDropdown();
  populateContractTitles();

  attachEventListeners();
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
        } else if (matchingKey.includes("caseType")) {
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
        } else if (matchingKey.endsWith("ClientLanguage")) {
          // client language dropdown change
          formState.update("clientLanguage", value);
            populateContractTitles();
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
        } else if (matchingKey.endsWith("customContractTitle")) {
          // Update form state with custom contract title
          formState.update("contractTitle", value);
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
        case ELEMENT_IDS.scheduleMenuBtn:
          showPage(ELEMENT_IDS.schedulePage);
          break;
        case ELEMENT_IDS.confirmMenuBtn:
          showPage(ELEMENT_IDS.confirmPage);
          break;
        case ELEMENT_IDS.contractMenuBtn:
          showPage(ELEMENT_IDS.contractPage);
          break;
        case ELEMENT_IDS.replyMenuBtn:
          showPage(ELEMENT_IDS.replyPage);
          break;
        case ELEMENT_IDS.wordContractMenuBtn:
          showPage(ELEMENT_IDS.wordContractPage);
          break;
        case ELEMENT_IDS.wordReceiptMenuBtn:
          showPage(ELEMENT_IDS.wordReceiptPage);
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
        case ELEMENT_IDS.confirmSubmitBtn:
        case ELEMENT_IDS.contractSubmitBtn:
        case ELEMENT_IDS.replySubmitBtn:
        case ELEMENT_IDS.wordContractSubmitBtn:
        case ELEMENT_IDS.wordReceiptSubmitBtn:
          submitForm();
          break;
        default:
          break;
      }
    });
  });
}

/**
 * Submits the current formState to the Python backend.
 * @return {Promise} - Resolves when the submission is complete.
 */
async function submitForm() {
  await window.pywebview.api.submit_form(formState);
}

window.addEventListener('pywebviewready', () => {
  console.log("[LawHub] pywebview is ready.");
  main();
});