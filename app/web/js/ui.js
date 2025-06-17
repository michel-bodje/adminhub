import {
  ELEMENT_IDS,
  formState,
  getLawyers,
  caseTypeHandlers,
  locationRules,
} from "./index.js";

/**
 * Displays the specified page by its ID and hides all other pages.
 *
 * This function assumes that all pages have the class "page" and that the
 * currently visible page has the additional class "active". It removes the
 * "active" class from all pages and adds it to the page with the given ID.
 *
 * @param {string} pageId - The ID of the page to display. The element with this
 * ID should exist in the DOM and have the class "page".
 */
export function showPage(pageId) {
  // Hide all pages
  const pages = document.querySelectorAll(".page");
  pages.forEach(page => page.classList.remove("active"));

  // Show the selected page
  const selectedPage = document.getElementById(pageId);
  if (selectedPage) {
    selectedPage.classList.add("active");
  }

  // Scroll to the top of the page
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/** Utility function to reset the page to its initial state. */
export function resetPage() {
  // Clear all form inputs
  const forms = document.querySelectorAll("form");
  forms.forEach(form => form.reset());
  
  // Clear state
  formState.reset();

  // Hide extra case details fields
  hideExtraFields();

  // Reset payment options
  handlePaymentOptions();

  // Navigate back to main menu
  showPage(ELEMENT_IDS.menuPage);

  // Reset dropdowns
  populateCaseTypeDropdown();
  populateLawyerDropdown();
  populateLocationDropdown();
  populateLanguageDropdown();
  populatePaymentDropdown();
}

/**
 * Shows or hides the loading overlay.
 *
 * @param {boolean} visible - Whether or not to display the overlay.
 */
export function showLoading(visible) {
  const overlay = document.getElementById("loading-overlay");
  overlay.classList.toggle("hidden", !visible);
}

/**
 * Displays a modal error message window with an "OK" button.
 * @param {string} errorMessage - The error message to display.
 */
export function showErrorModal(errorMessage) {
  const modal = document.getElementById("error-modal");
  const messageElement = document.getElementById("error-message");
  const okButton = document.getElementById("error-ok-button");

  if (!modal || !messageElement || !okButton) {
    console.error("Error modal elements not found.");
    return;
  }

  // Set the error message
  messageElement.textContent = errorMessage;

  // Show the modal
  modal.classList.remove("hidden");

  // Add event listener to the OK button to close the modal
  okButton.onclick = () => {
    modal.classList.add("hidden");
  };
}

/**
 * Opens a popup window with the specified content or loads an external HTML file.
 * @param {Object} options - The options for the popup.
 * @param {string} options.title - The title of the popup window.
 * @param {string} options.contentOrFile - The HTML content to display or the path to an external HTML file.
 * @param {boolean} [options.isFile=false] - Whether the second parameter is a file path.
 * @param {string} [options.styles=""] - Additional CSS styles to apply to the popup.
 * @param {number} [options.width=684] - The width of the popup window.
 * @param {number} [options.height=600] - The height of the popup window.
 * @param {string} [options.position="center"] - The position of the popup ("center" or "bottom-right").
 */
export function openPopup({
  title,
  contentOrFile,
  isFile = false,
  styles = "",
  width = 684,
  height = 600,
  position = "center",
}) {
  // Calculate position
  let left = 0;
  let top = 0;

  if (position === "center") {
    left = (window.screen.width - width) / 2;
    top = (window.screen.height - height) / 2;
  } else if (position === "bottom-right") {
    left = window.screen.width - width;
    top = window.screen.height - height;
  }

  const popupWindow = window.open("", title, `width=${width},height=${height},left=${left},top=${top}`);

  const applyStyles = () => {
    const styleElement = popupWindow.document.createElement("style");
    styleElement.textContent = `
      body {
        background-color: #1e1e1e;
        color: #f4f4f4;
        font-family: 'Trebuchet MS', 'Tahoma', 'Arial', sans-serif;
        line-height: 1.6;
        margin: 0;
        padding: 1rem;
      }
      a {
        color: #0078d4;
        text-decoration: none;
      }
      a:hover {
        text-decoration: underline;
      }
      h3 {
        color: #0078d4;
      }
            button {
        background-color: #0078d4;
        color: #fff;
        border: none;
        padding: 0.5rem 1rem;
        font-size: 1rem;
        cursor: pointer;
        border-radius: 4px;
      }
      button:hover {
        background-color: #005a9e;
      }
      ${styles}
    `;
    popupWindow.document.head.appendChild(styleElement);
  };

  if (isFile) {
    // Load the external HTML file
    fetch(contentOrFile)
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Failed to load file: ${contentOrFile}`);
        }
        return response.text();
      })
      .then((html) => {
        popupWindow.document.write(html);
        popupWindow.document.close();
        applyStyles(); // Apply styles after the content is loaded
      })
      .catch((error) => {
        console.error("Error loading file:", error);
        popupWindow.document.write(`
          <html>
            <head><title>Error</title></head>
            <body>
              <h3>Failed to load the requested file.</h3>
              <p>${error.message}</p>
              <button onclick="window.close()">Close</button>
            </body>
          </html>
        `);
        popupWindow.document.close();
        applyStyles(); // Apply styles to the error page
      });
  } else {
    // Use the provided HTML content
    popupWindow.document.write(`
      <html>
        <head>
          <title>${title}</title>
        </head>
        <body>
          ${contentOrFile}
          <button onclick="window.close()">Close</button>
        </body>
      </html>
    `);
    popupWindow.document.close();
    applyStyles(); // Apply styles immediately
  }
}

/**
 * Displays a popup with available calendar slots for the next 2 weeks,
 * with the selected slot at the top. The popup is only displayed when
 * auto-scheduling is enabled.
 * @param {Array<{start: Date, end: Date}>} validSlots - The array of valid slot objects with start and end times.
 * @param {{start: Date, end: Date}} selectedSlot - The selected slot object with start and end times.
 */
export function popupAvailableSlots(validSlots, selectedSlot) {
  let popupContent = "<h3>Next 5 available calendar slots</h3><ul>";

  // Ensure the selected slot is displayed first
  const selectedSlotIndex = validSlots.indexOf(selectedSlot);

  if (selectedSlotIndex !== -1) {
    const slot = validSlots[selectedSlotIndex];
    popupContent += formatSlot(slot, true);
  }

  // Display the next 4 future slots after the selected slot
  let previousDate = selectedSlot.start.toLocaleDateString();
  for (let i = selectedSlotIndex + 1; i < validSlots.length && i <= selectedSlotIndex + 4; i++) {
    const slot = validSlots[i];

    // Insert a line break if the day changes
    if (slot.start.toLocaleDateString() !== previousDate) {
      popupContent += "<br>";
      previousDate = slot.start.toLocaleDateString();
    }

    popupContent += formatSlot(slot);
  }

  popupContent += "</ul>";

  openPopup({
    title: "Available Slots",
    contentOrFile: popupContent,
    width: 480,
    height: 300,
    position: "bottom-right",
  });

  /**
   * Formats a slot into an HTML list item.
   * @param {Object} slot - The slot object with start and end times.
   * @param {boolean} [isSelected=false] - Whether the slot is the selected one.
   * @returns {string} - The formatted HTML string for the slot.
  */
  function formatSlot(slot, isSelected = false) {
    const startDate = slot.start.toLocaleString([], { weekday: "long", year: "numeric", month: "long", day: "numeric" });
    const startTime = slot.start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    const endTime = slot.end.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

    return `<li>${isSelected ? "<strong>Selected Slot:</strong><br>" : ""}${startDate} - ${startTime} to ${endTime}</li>`;
  }
}

/** Utility function to hide all extra form fields and reset their values. */
export function hideExtraFields() {
  // Select all extra fields
  const caseDetailsFields = document.querySelectorAll("div[id$='-details']");

  // Clear their input values
  caseDetailsFields.forEach(field => {
    field.hidden = true;
    const inputs = field.querySelectorAll("input, textarea");
    inputs.forEach(input => input.value = "");
  });

  // Hide the entire section
  const caseDetailsElement = document.getElementById(ELEMENT_IDS.caseDetails);

  if (caseDetailsElement) {
    caseDetailsElement.hidden = true;
  }
}

/** Utility function to clear a dropdown and add a placeholder option.
 * @param {HTMLSelectElement} dropdown - The dropdown element to clear.
 * This function clears all existing options and adds a placeholder option.
 * @param {string} placeholder - The placeholder text for the dropdown.
*/
function emptyDropdown(dropdown, placeholder) {
  // Clear existing options
  dropdown.innerHTML = "";

  // Add placeholder option
  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = placeholder;
  placeholderOption.disabled = true;
  placeholderOption.selected = true;
  placeholderOption.hidden = true;
  dropdown.appendChild(placeholderOption);
}

/** 
 * Utility function to populate a dropdown with options.
 * @param {string} elementId - The ID of the dropdown element.
 * @param {Array} options - The options to populate the dropdown with.
 * @param {string} [placeholder="Select an option"] - The placeholder text for the dropdown.
 */
function populateDropdown(elementId, options, placeholder = "Select an option") {
  // Select the dropdown
  const dropdown = document.getElementById(elementId);

  if (!dropdown) {
    console.error(`Dropdown with ID ${elementId} not found.`);
    return;
  }

  emptyDropdown(dropdown, placeholder);

  // Add options
  options.forEach(option => {
    const opt = document.createElement("option");
    opt.value = option.value;
    opt.textContent = option.label;
    dropdown.appendChild(opt);
  });
}

/** Dynamically loads the client language options. */
export function populateLanguageDropdown() {
  // Language dropdown elements
  const elements = Object.keys(ELEMENT_IDS)
    .filter(key => key.endsWith("ClientLanguage"))
    .map(key => ELEMENT_IDS[key]);

  // Populate dropdowns
  elements.forEach(element => {
    populateDropdown(element, [
      { value: "English", label: "English" },
      { value: "Français", label: "Français" },
    ],
    "Select Language");
  });
}

/** Dynamically loads the payment method options. */
export function populatePaymentDropdown() {
  // Payment dropdown elements
  const elements = Object.keys(ELEMENT_IDS)
    .filter(key => key.endsWith("PaymentMethod"))
    .map(key => ELEMENT_IDS[key]);

  // Populate dropdowns
  elements.forEach(element => {
    populateDropdown(element, [
      { value: "cash", label: "Cash" },
      { value: "cheque", label: "Cheque" },
      { value: "credit", label: "Credit" },
      { value: "e-transfer", label: "E-Transfer" },
    ],
    "Select Payment Method");
  });
}

/** Dynamically loads the case type options. */
export function populateCaseTypeDropdown() {
  const lawyers = getLawyers();
  if (!lawyers.length) return;

  // Collect all unique case types from all lawyers
  let caseTypes = [...new Set(
    lawyers.flatMap(lawyer => lawyer.specialties)
  )].map(caseType => {
    const handlerExists = caseTypeHandlers[caseType];
    const label = handlerExists ? caseTypeHandlers[caseType].label : caseType;
    return { value: caseType, label };
  });

  // Sort case types alphabetically, except for "Other (Specify)" which should be last
  caseTypes = caseTypes.sort((a, b) => {
    if (a.label === "Other (Specify)") return 1;
    if (b.label === "Other (Specify)") return -1;
    return a.label.localeCompare(b.label);
  });

  // Populate the case type dropdown
  populateDropdown(ELEMENT_IDS.caseType, caseTypes, "Select Case Type");
}

/** Dynamically loads the lawyer options based on the selected case type. */
export function populateLawyerDropdown() {
  const lawyers = getLawyers();
  if (!lawyers.length) return;

  // Get the selected case type (if any)
  const selectedCaseType = formState.caseType;

  // Lawyer dropdown elements
  const elements = Object.keys(ELEMENT_IDS)
    .filter(key => key.endsWith("LawyerId"))
    .map(key => ELEMENT_IDS[key]);

  // Filter lawyers based on the selected case type (if applicable)
  const filteredLawyers = selectedCaseType
    ? lawyers.filter(lawyer => lawyer.specialties.includes(selectedCaseType))
    : lawyers; // If no case type, include all lawyers

  // Populate the dropdowns
  elements.forEach(element => {
    populateDropdown(
      element,
      filteredLawyers.map(lawyer => ({
        value: lawyer.id,
        label: `${lawyer.name} (${lawyer.id})`
      })),
      "Select Lawyer"
    );
  });
}

/** Dynamically loads the location options based on the selected lawyer. */
export function populateLocationDropdown() {
  // Location dropdown elements
  const elements = Object.keys(ELEMENT_IDS)
    .filter(key => key.endsWith("Location"))
    .map(key => ELEMENT_IDS[key]);

  const locations = locationRules.locations;

  elements.forEach(element => {
    populateDropdown(
      element,
      locations.map((location) => ({
        value: location,
        label: location.charAt(0).toUpperCase() + location.slice(1),
      })),
      "Select Location"
    );
  });
}

/** Handles the case type dropdown change event. */
export function handleCaseDetails() {
  const caseType = formState.caseType;
  const detailsContainer = document.getElementById(ELEMENT_IDS.caseDetails);

  // Hide all case details fields
  hideExtraFields();

  // Dynamically show the selected case details field using the handler
  if (caseTypeHandlers[caseType]) {
    const handlerId = `${caseType}-details`;
    const detailsElement = document.getElementById(handlerId);
    if (detailsElement) {
      detailsContainer.hidden = false;
      detailsElement.hidden = false;
    }
  } else {
    // If no valid case type is selected, hide the entire case details section
    detailsContainer.hidden = true;
  }
}

/**
 * Toggles the visibility of the payment options dropdown based on the "Payment Made" checkbox.
 */
export function handlePaymentOptions() {
  const paymentOptionsContainer = document.getElementById(ELEMENT_IDS.paymentOptionsContainer);

  if (!paymentOptionsContainer) {
    console.error("Payment options container not found.");
    return;
  }

  if (formState.isPaymentMade) {
    paymentOptionsContainer.hidden = false;
  } else {
    paymentOptionsContainer.hidden = true; 
  }
}
