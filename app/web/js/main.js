import {
  ELEMENT_IDS,
  MENU_ROUTES,
  formState,
  showPage,
  resetPage,
  loadLawyers,
  setLawyers,
  populateLawyerDropdown,
  handleCaseDetails,
  handlePaymentOptions,
  fileState,
} from "./index.js";
import { ACTIONS_BY_ID } from "./actions.js";

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
  console.log("[AdminHub] Application initialized.");}

/** Attaches all event listeners for the application. */
function attachEventListeners() {
  // Form input event listeners
  const formInputs = document.querySelectorAll('input, select, textarea');

  // Define handlers for each form field
  const FIELD_HANDLERS = {
    MatterId: (value) => formState.update("matterId", value),
    LawyerId: (value) => formState.update("lawyerId", value),
    Location: (value) => formState.update("location", value),
    CaseType: (value) => {
      formState.update("caseType", value);
      populateLawyerDropdown();
      handleCaseDetails();
    },
    ClientName: (value) => formState.update("clientName", value),
    ClientPhone: (value) => formState.update("clientPhone", value),
    ClientEmail: (value) => formState.update("clientEmail", value),
    ClientTitle: (value) => formState.update("clientTitle", value),
    ClientLanguage: (value) => formState.update("clientLanguage", value),
    Date: (value) => formState.update("appointmentDate", value),
    Time: (value) => formState.update("appointmentTime", value),
    FirstConsultation: (value, el) => formState.update("isFirstConsultation", el.checked),
    RefBarreau: (value, el) => formState.update("isRefBarreau", el.checked),
    ExistingClient: (value, el) => formState.update("isExistingClient", el.checked),
    PaymentMade: (value, el) => {
      formState.update("isPaymentMade", el.checked);
      handlePaymentOptions();
    },
    PaymentMethod: (value) => formState.update("paymentMethod", value),
    notes: (value) => formState.update("notes", value),
    Deposit: (value) => formState.update("depositAmount", value),
    wordContractTitle: (value) => {
      const input = document.getElementById(ELEMENT_IDS.customContractTitle);
      if (value === "other") {
        input.classList.remove("hidden");
        input.required = true;
      } else {
        input.classList.add("hidden");
        input.required = false;
        formState.update("contractTitle", value);
      }
    },
    customContractTitle: (value) => formState.update("contractTitle", value),
    wordReceiptReason: (value) => {
      const input = document.getElementById(ELEMENT_IDS.customReceiptReason);
      if (value === "other") {
        input.classList.remove("hidden");
        input.required = true;
      } else {
        input.classList.add("hidden");
        input.required = false;
        formState.update("receiptReason", value);
      }
    },
    customReceiptReason: (value) => formState.update("receiptReason", value),
    scheduleMode: (value) => {
      const d = ELEMENT_IDS.manualDate;
      const t = ELEMENT_IDS.manualTime;
      const dateEl = document.getElementById(d);
      const timeEl = document.getElementById(t);
      const dateLabel = document.querySelector(`label[for=${d}]`);
      const timeLabel = document.querySelector(`label[for=${t}]`);

      if (value === "manual") {
        [dateLabel, timeLabel, dateEl, timeEl].forEach(el => el.classList.remove("hidden"));
        dateEl.required = true;
        timeEl.required = true;
      } else {
        [dateLabel, timeLabel, dateEl, timeEl].forEach(el => el.classList.add("hidden"));
        dateEl.required = false;
        timeEl.required = false;
      }
    }
  };

  formInputs.forEach((input) => {
    input.addEventListener('change', (event) => {
      const { id, value } = event.target;

      const matchingKey = Object.keys(ELEMENT_IDS).find((key) => ELEMENT_IDS[key] === id);
      const handlerEntry = Object.entries(FIELD_HANDLERS).find(([suffix]) => matchingKey && matchingKey.endsWith(suffix));

      if (handlerEntry) {
        const [, handler] = handlerEntry;
        handler(value, event.target);
      } else {
        console.warn(`No handler for input with id '${id}'`);
      }
    });
  });

  // Menu buttons event listeners
  const menuButtons = document.querySelectorAll('.menu-btn');
  menuButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      const pageId = MENU_ROUTES[event.target.id];
      if (pageId) {
        showPage(pageId);
      } else {
        console.warn(`No page mapped for menu button: ${event.target.id}`);
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
      const action = ACTIONS_BY_ID[event.target.id];
      if (action) {
        action();
      } else {
        console.warn(`No action found for submit button: ${event.target.id}`);
      }
    });
  });

  // Select file button event listener
  const selectFileBtn = document.getElementById(ELEMENT_IDS.selectFileBtn);
  if (selectFileBtn) {
    selectFileBtn.addEventListener('click', async () => {
      try {
        const filePath = await window.pywebview.api.select_timesheet_file();
        if (filePath) {
          fileState.selectedFilePath = filePath;
          const pathDisplay = document.getElementById(ELEMENT_IDS.selectedFilePath);
          if (pathDisplay) {
            pathDisplay.textContent = filePath;
          }
          
          // Enable the submit button
          const submitBtn = document.getElementById(ELEMENT_IDS.timeEntriesSubmitBtn);
          if (submitBtn) {
            submitBtn.disabled = false;
          }
        }
      } catch (error) {
        console.error("[AdminHub] Error selecting file:", error);
        alert("Failed to select file.");
      }
    });
  }

  /** Initialize form state with pre-checked checkbox values */
  function initializeFormState() {
    // Get all checkboxes and initialize their state
    const checkboxes = document.querySelectorAll('input[type="checkbox"]');

    checkboxes.forEach((checkbox) => {
      if (checkbox.checked) {
        const { id } = checkbox;
        const matchingKey = Object.keys(ELEMENT_IDS).find((key) => ELEMENT_IDS[key] === id);
        const handlerEntry = Object.entries(FIELD_HANDLERS).find(([suffix]) => matchingKey && matchingKey.endsWith(suffix));

        if (handlerEntry) {
          const [, handler] = handlerEntry;
          handler(checkbox.value, checkbox);
          console.log(`[AdminHub] Initialized ${id} to checked state`);
        }
      }
    });
  }

  initializeFormState();
  console.log("[AdminHub] Event listeners attached.");
}
