import {
  ELEMENT_IDS,
  formState,
} from "./index.js";

export const locationRules = {
  // Centralized list of locations
  locations: ["office", "phone", "teams"],
  
  // List of unavailability for each lawyer
  lawyerUnavailability: {
    DH: {
      office: ["Monday"],
    },
    TG: {
      office: ["Friday"],
    }
  },
};

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
    label: "Regimes de Protection",
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

/** Collects extra case details fields based on the selected case type. */
export function collectCaseDetails() {
  const caseType = formState.caseType;
  const details = {};

  switch (caseType) {
    case "divorce":
      details.spouseName = document.getElementById(ELEMENT_IDS.spouseName)?.value || "";
      details.conflictSearchDone = document.getElementById(ELEMENT_IDS.conflictSearchDoneDivorce)?.checked || false;
      break;
    case "estate":
      details.deceasedName = document.getElementById(ELEMENT_IDS.deceasedName)?.value || "";
      details.executorName = document.getElementById(ELEMENT_IDS.executorName)?.value || "";
      details.conflictSearchDone = document.getElementById(ELEMENT_IDS.conflictSearchDoneEstate)?.checked || false;
      break;
    case "employment":
      details.employerName = document.getElementById(ELEMENT_IDS.employerName)?.value || "";
      break;
    case "contract":
      details.otherPartyName = document.getElementById(ELEMENT_IDS.otherPartyName)?.value || "";
      break;
    case "defamations":
      details.otherPartyName = document.getElementById(ELEMENT_IDS.otherPartyName)?.value || "";
      break;
    case "mandates":
      details.mandateDetails = document.getElementById(ELEMENT_IDS.mandateDetails)?.value || "";
      break;
    case "business":
      details.businessName = document.getElementById(ELEMENT_IDS.businessName)?.value || "";
      break;
    case "common":
      details.commonField = document.getElementById(ELEMENT_IDS.commonField)?.value || "";
      break;
    // Add more case types as needed
    default:
      break;
  }
  return details;
}

/**
 * Formats a phone number for display.
 * Tries to format the number intelligently based on its length and region.
 * @param {string} number - The phone number to format.
 * @param {string} [region="CA"] - The region to format the number for (default: "CA").
 * @returns {string} The formatted phone number.
 */
export function formatPhoneNumber(number, region = "CA") {
  if (!number || typeof number !== "string") return number;

  const digitsOnly = number.replace(/\D/g, "");

  if (digitsOnly.length === 10) {
    return `${digitsOnly.slice(0, 3)}-${digitsOnly.slice(3, 6)}-${digitsOnly.slice(6)}`;
  } else if (digitsOnly.length === 11 && digitsOnly.startsWith("1")) {
    return `+1 ${digitsOnly.slice(1, 4)}-${digitsOnly.slice(4, 7)}-${digitsOnly.slice(7)}`;
  } else if (digitsOnly.length >= 11) {
    return `+${digitsOnly}`;
  }

  return number;
}

/**
 * Utility function to validate a phone number (with optional region).
 * 
 * Only supports Canadian/US numbers (1-10 digits) or international numbers (8-15 digits).
 * @param {string} number - The phone number to validate.
 * @param {string} [region="CA"] - The region to validate the number against (default: "CA").
 * @returns {boolean} True if valid, false otherwise.
 */
export function isValidPhoneNumber(number, region = "CA") {
  if (!number || typeof number !== "string" || number.trim() === "") return false;

  const digitsOnly = number.replace(/\D/g, "");

  let normalized = digitsOnly;
  if (digitsOnly.length === 10) {
    // Add leading '1' for North America
    normalized = "1" + digitsOnly;
  }

  // Must start with non-zero and be 8–15 digits
  return /^[1-9]\d{7,14}$/.test(normalized);
}

/**
 * Utility function to validate an email address.
 * @param {string} email - The email address to validate.
 * @returns {boolean} True if valid, false otherwise.
 */
export function isValidEmail(email) {
  if (!email || typeof email !== "string" || email.trim() === "") return false;

  const lowerEmail = email.toLowerCase();
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(lowerEmail);
}
