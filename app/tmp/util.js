import {
  ELEMENT_IDS,
  formState,
} from "../index.js";

/** Handles the case type details based on the selected case type. */
export const caseTypeHandlers = {
  divorce: {
    label: "Divorce / Family Law",
    handler: function () {
      const spouseName = document.getElementById(ELEMENT_IDS.spouseName).value;
      const conflictSearchDone = document.getElementById(ELEMENT_IDS.conflictSearchDoneDivorce).checked;
      return `
        ${this.label}
        <p><strong>Spouse Name:</strong> ${spouseName}</p>
        <p>Conflict Search Done? ${conflictSearchDone ? "✔️" : "❌"}</p>
      `;
    },
  },
  estate: {
    label: "Successions / Estate Law",
    handler: function () {
      const deceasedName = document.getElementById(ELEMENT_IDS.deceasedName).value;
      const executorName = document.getElementById(ELEMENT_IDS.executorName).value;
      const conflictSearchDone = document.getElementById(ELEMENT_IDS.conflictSearchDoneEstate).checked;
      return `
        ${this.label}
        <p><strong>Deceased Name:</strong> ${deceasedName}</p>
        <p><strong>Executor Name:</strong> ${executorName}</p>
        <p>Conflict Search Done? ${conflictSearchDone ? "✔️" : "❌"}</p>
      `;
    },
  },
  employment: {
    label: "Employment Law",
    handler: function () {
      const employerName = document.getElementById(ELEMENT_IDS.employerName).value;
      return `
        ${this.label}
        <p><strong>Employer Name:</strong> ${employerName}</p>
      `;
    },
  },
  contract: {
    label: "Contract Law",
    handler: function () {
      const otherPartyName = document.getElementById(ELEMENT_IDS.otherPartyName).value;
      return `
        ${this.label}
        <p><strong>Other Party:</strong> ${otherPartyName}</p>
      `;
    },
  },
  defamations: {
    label: "Defamations",
    handler: function () {
      const otherPartyName = document.getElementById(ELEMENT_IDS.otherPartyName).value;
      return `
        ${this.label}
      `;
    },
  },
  real_estate: {
    label: "Real Estate",
    handler: function () {
      return `
        ${this.label}
      `;
    },
  },
  name_change: {
    label: "Name Change",
    handler: function () {
      return `
        ${this.label}
      `;
    },
  },
  adoptions: {
    label: "Adoptions",
    handler: function () {
      return `
        ${this.label}
      `;
    },
  },
  mandates: {
    label: "Regimes de Protection / Mandates",
    handler: function () {
      const mandateDetails = document.getElementById(ELEMENT_IDS.mandateDetails).value;
      return `
        ${this.label}
        <p><strong>Mandate Details:</strong> ${mandateDetails}</p>
      `;
    },
  },
  business: {
    label: "Business Law",
    handler: function () {
      const businessName = document.getElementById(ELEMENT_IDS.businessName).value;
      return `
        ${this.label}
        <p><strong>Business Name:</strong> ${businessName}</p>
      `;
    },
  },
  assermentation: {
    label: "Assermentation",
    handler: function () {
      return `
        ${this.label}
      `;
    },
  },
  // A catch-all option for unspecified case types
  common: {
    label: "Other (Specify)",
    handler: function () {
      const commonField = document.getElementById(ELEMENT_IDS.commonField).value;
      if (!commonField) {
        console.error("Please specify the details for the case type.");
        throw new Error("Missing common field details");
      }
      return `
        ${commonField}
      `;
    },
  },
};

/**
 * Returns the case details based on the selected case type.
 * If the case type is not found, throws an error.
 * @returns {string} - The case details as a string.
 * @throws {Error} - If the case type is not found.
 */
export function getCaseDetails() {
  try {
    const details = caseTypeHandlers[formState.caseType]?.handler();
    if (details) {
      return details;
    }
    throw new Error("Selected case type not found.");
  } catch (error) {
    console.error(error.message);
  }
}

/**
 * Checks if all required fields in the form have been filled in and the phone number and email are valid.
 * @returns {boolean} - True if all required fields are valid, false otherwise.
 */
export function isValidInputs() {
  try {
    const caseType = formState.caseType;

    if (!formState.clientName || !formState.clientPhone || !formState.clientEmail || !formState.clientLanguage || !caseType) {
      throw new Error("Please fill in all required fields.");
    }

    if (!isValidPhoneNumber(formState.clientPhone)) {
      throw new Error("Please provide a valid phone number in the format 555-555-5555.");
    }

    if (!isValidEmail(formState.clientEmail)) {
      throw new Error("Please provide a valid email address.");
    }

    return true;
  } catch (error) {
    console.error(error.message);
    return false;
  }
}

/** Utility function to validate an email address.
 * @param {string} email - The email address to validate.
 * @returns {boolean} - True if valid, false otherwise.
 * */
export function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.toLowerCase());
}

/** Utility function to validate an international phone number in E.164 format.
 * @param {string} phone - The phone number to validate.
 * @returns {boolean} - True if valid, false otherwise.
 */
export function isValidPhoneNumber(phone) {
  // E.164 format 
  const phoneRegex = /^\+?[1-9]\d{1,14}$/;

  // Remove spaces and dashes before testing
  return phoneRegex.test(phone.replace(/[\s-]/g, ""));
}

/** Utility function to format a phone number for display.
 * @param {string} phone - The phone number to format.
 * @returns {string} - The formatted phone number.
 */
export function formatPhoneNumber(phone) {
  // Ensure the number starts with a "+"
  if (!phone.startsWith("+")) {
    phone = `+${phone}`;
  }

  // Remove all non-digit characters except "+"
  phone = phone.replace(/[^\d+]/g, "");

  // Add spaces or dashes for readability (e.g., +1 555-555-5555)
  return phone.replace(/(\+\d{1,3})(\d{3})(\d{3})(\d+)/, "$1 $2-$3-$4");
}