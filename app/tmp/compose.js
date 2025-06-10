import {
  formState,
  getLawyer,
  getCaseDetails,
  htmlTemplates,
  loadDocxTemplate,
  isValidEmail,
} from "../index.js";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";

/**
 * Sets the subject in the draft email.
 * @param {string} subject - The email subject.
 */
function setSubject(subject) {
  Office.context.mailbox.item.subject.setAsync(subject, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set subject:", asyncResult.error.message);
    }
  });
}

/**
 * Sets the body in the draft email.
 * @param {string} body - The email body.
 */
function setBody(body) {
  Office.context.mailbox.item.body.setAsync(
    body,
    { coercionType: Office.CoercionType.Html },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set body:", asyncResult.error.message);
      }
    }
  );
}

/**
 * Sets the recipient in the draft email.
 * @param {string} recipient - The recipient's email address.
 */
function setRecipient(recipient) {
  Office.context.mailbox.item.to.setAsync([recipient], (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set recipient:", asyncResult.error.message);
    }
  });
}

/**
 * Sets the attendees in the draft meeting.
 * @param {Array} attendees - An array of attendee objects with displayName and emailAddress.
 */
function setAttendees(attendees) {
  Office.context.mailbox.item.requiredAttendees.setAsync(attendees, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set attendees:", asyncResult.error.message);
    }
  });
}

/**
 * Sets the category in the draft meeting.
 * @param {Array<string>} category - The category string array.
 */
function setCategory(category) {
  // Check if categories API is available
  if (Office.context.mailbox.item.categories) {
    removeAllCategories(() => addCategory(category));
  } else {
    console.error("Categories API is not available.");
  }

  function getCategories(callback) {
    Office.context.mailbox.item.categories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to get categories. Error:", result.error.message);
        callback([]);
      } else {
        callback(result.value || []);
      }
    });
  }

  function removeAllCategories(callback) {
    getCategories((categories) => {
      if (categories.length === 0) {
        callback();
        return;
      }
      const categoriesToRemove = categories.map((category) => category.displayName);
      
      Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, (removeResult) => {
        if (removeResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to remove categories. Error:", removeResult.error.message);
        }
        callback();
      });
    });
  }

  function addCategory(category) {
    Office.context.mailbox.item.categories.addAsync(category, (addResult) => {
      if (addResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set category:", addResult.error.message);
      }
    });
  }
}

/**
 * Sets the location in the draft meeting.
 * @param {string} location - The location to set for the meeting.
 */
function setLocation(location) {
  Office.context.mailbox.item.location.setAsync(location, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set location:", asyncResult.error.message);
    } else {
      console.log("Location set successfully:", location);
    }
  });
}

/**
 * Sets the start and end times for the draft meeting.
 * @param {Date} startTime - The start time of the meeting.
 * @param {Date} endTime - The end time of the meeting.
 */
function setMeetingTimes(startTime, endTime) {
  Office.context.mailbox.item.start.setAsync(startTime, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set start time:", asyncResult.error.message);
    }
  });

  Office.context.mailbox.item.end.setAsync(endTime, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set end time:", asyncResult.error.message);
    }
  });
}

/**
 * Returns the email subject line for the given language and type.
 * @param {string} language - The language of the email.
 * @param {string} type - The type of email.
 * @returns {string} The subject line.
 */
function getSubject(language, type) {
  if (language === "fr") {
    if (type === "contract") {
      return "Contrat de services - Allen Madelin"
    } else if (type === "reply") {
      return "Réponse - Allen Madelin";
    } else {
      return "Confirmation de rendez-vous - Allen Madelin";
    }
  } else {
    if (type === "contract") {
      return "Contract of services - Allen Madelin";
    } else if (type === "reply") {
      return "Reply - Allen Madelin";
    } else {
      return "Confirmation of appointment - Allen Madelin";
    }
  }
}

/**
 * Adds taxes to the given amount.
 * @param {number} amount - The amount to which taxes will be added.
 * @param {Boolean} addFOF - Whether or not to add the file opening fee.
 * @returns {number} The total amount with taxes.
 * @throws {Error} If the amount is not a number.
 */
function addTaxes(amount, addFOF = false) {
  if (isNaN(amount)) {
    console.error("Amount is not a number:", amount);
    throw new Error("Amount is not a valid number.");
  }

  // GST + QST + 100$ file opening fee
  // GST: 5% + QST: 9.975% 
  const fof = 100;
  let total = (amount * (1 + 0.05 + 0.09975));
  if (addFOF) total += fof;

  return total;
}


/**
 * Generates a draft email with the given type and options.
 * @param {string} [type=formState.location] - The type of email to generate.
 * @param {Object} [options={}] - Additional options to customize the email.
 * @param {Object} [options.state=formState] - The state object containing client info.
 * @param {Date} [options.slot=null] - A date and time slot to use for the appointment.
 * @returns {Promise<void>} A promise that resolves if the email is created successfully.
 * @throws {Error} If any of the required information is missing or invalid.
 */
export async function createEmail(type = formState.location, options = {}) {
  const { state = formState, slot = null } = options;
  try {
    const clientEmail = state.clientEmail;
    const language = state.clientLanguage === "Français" ? "fr" : "en";
    let body = htmlTemplates[language][type];
    
    const dateTime = slot
      ? slot
      : (state.appointmentDate && state.appointmentTime
        ? new Date(`${state.appointmentDate}T${state.appointmentTime}`)
        : null
      )
    ;

    if (!isValidEmail(clientEmail)) {
      throw new Error("Please provide a valid email address.");
    }

    if (!body) {
      throw new Error(`No template found for type "${type}" in language "${language}".`);
    }

    // Only validate date and time for appointment confirmations
    if (type === "office" || type === "teams" || type === "phone") {
      if (!dateTime) {
        throw new Error("No appointment date and time provided.");
      }

      const date = dateTime.toLocaleDateString(language == "fr" ? "fr-CA" : "en-US", {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric",
      });
  
      const time = dateTime.toLocaleTimeString(language == "fr" ? "fr-CA" : "en-US", {
        hour: "2-digit",
        minute: "2-digit",
      });

      const rates = state.isFirstConsultation ? 125 : 350;
      const totalRates = addTaxes(rates);
      
      body = body
        .replace("{{date}}", date)
        .replace("{{time}}", time)
        .replace("{{rates}}", Number(rates).toFixed())
        .replace("{{totalRates}}", Number(totalRates).toFixed(2))
      ;
    } else if (type === "contract") {
      // Deposit for contract email
      const depositAmount = state.depositAmount;
      const totalAmount = addTaxes(state.depositAmount, true);

      body = body
        .replace("{{depositAmount}}", Number(depositAmount).toFixed())
        .replace("{{totalAmount}}", Number(totalAmount).toFixed(2))
      ;
    }

    body = body
      .replace("{{lawyerName}}", getLawyer(state.lawyerId).name)
    ;
  
    setSubject(getSubject(language, type));
    setRecipient(clientEmail);
    setBody(body);

  } catch (error) {
    console.error("createEmail:", error);
    throw error; // Rethrow the error for further handling if needed
  }
}

/**
 * Creates a meeting draft with the specified details.
 * @param {{ start: Date, end: Date, location: string }} selectedSlot - The proposed time slot.
 */
export async function createMeeting(selectedSlot) {
  try {
    const lawyer = getLawyer(formState.lawyerId);
    
    // Capitalize location string
    const location = selectedSlot.location.charAt(0).toUpperCase() + selectedSlot.location.slice(1);

    // Construct the case details
    const caseDetails = getCaseDetails();
    
    // Construct the subject and body
    const subject = `${formState.clientName} (ma)`;

    const priceDetails = (() => {
      if (formState.isRefBarreau) {
        return '<u><em>Ref. Barreau ($60+tax)</em></u>';
      } else if (formState.isFirstConsultation) {
        return '<span style="background-color: yellow;">First Consultation ($125+tax)</span>';
      } else {
        return '<span style="background-color: #d3d3d3;">Follow-up ($350+tax)</span>';
      }
    })();
    
    const body = `
      <p>Client:&nbsp;&nbsp;&nbsp;&nbsp;${formState.clientName}<br>

      Phone:&nbsp;&nbsp;&nbsp;${formState.clientPhone}<br>

      Email:&nbsp;&nbsp;&nbsp;&nbsp;${formState.clientEmail}<br>

      Lang:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${formState.clientLanguage}</p>

      ${formState.isExistingClient
        ? '<p style="color: green;"><strong>Existing Client</strong></p>'
        : `${priceDetails}: ${caseDetails}</p>

      <p><strong>Payment</strong>  ${formState.isPaymentMade ? "✔️" : "❌"}<br>
      
      ${formState.isPaymentMade ? `${formState.paymentMethod} (ma)` : ""}</p>
      `
      }
      <p>Notes:<br>
      <span style="font-style: italic">${formState.notes}</span></p>
    `;

    // Set details for the draft meeting
    setCategory([lawyer.name]);
    setSubject(subject);
    setMeetingTimes(selectedSlot.start, selectedSlot.end);
    setAttendees([{ displayName: lawyer.name, emailAddress: lawyer.email }]);
    setLocation(location);
    setBody(body);

  } catch (error) {
    console.error("createMeeting:", error);
    throw error; // Rethrow the error for further handling if needed
  }
}

/**
 * Prepares a confirmation email by storing the necessary data in localStorage
 * and notifies the user to open a new email draft for auto-filling.
 * @param {{ start: Date, end: Date, location: string }} slot - The chosen time slot
 */
export function prepareConfirmation(slot) {
    // Only run if formState and slot are defined
    if (!formState || !slot) {
        console.error("Missing formState or selected slot for confirmation email.");
        return;
    }

    const payload = {
        formState,
        slot
    };

    // Store data in localStorage so the taskpane in new draft can access it
    localStorage.setItem('confirmationPayload', JSON.stringify(payload));

    // Open a new email draft window
    // Office.context.mailbox.displayNewMessageForm({});

    // Notify user
    console.warn("Confirmation data saved. Open a new email draft and the taskpane will auto-fill it.");
}

/**
 * Creates a contract document in Word using the Office.js API.
 * @returns {Promise<void>}
 */
export async function createContract() {
  const {
    clientName,
    clientEmail,
    clientLanguage,
    depositAmount,
    contractTitle,
  } = formState;

  // Basic input validation
  if (!clientName || !clientEmail || !contractTitle || !depositAmount) {
    console.error("One or more inputs are missing.");
    return;
  }

  if (!isValidEmail(clientEmail)) {
    console.error("Invalid email format.");
    return;
  }

  const language = clientLanguage === "Français" ? "fr" : "en";

  try {
    // Load the DOCX template as a binary string
    const templateBinary = await loadDocxTemplate(language, "contract");

    // Initialize PizZip with the template binary
    const zip = new PizZip(templateBinary);

    // Initialize Docxtemplater
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Define placeholders
    const placeholders = {
      clientName,
      clientEmail: "{{clientEmail}}", // Placeholder for email link
      contractTitle,
      depositAmount: Number(depositAmount).toFixed(),
      totalAmount: Number(addTaxes(depositAmount, true)).toFixed(2),
      date: new Date().toLocaleDateString(language === "fr" ? "fr-CA" : "en-US", {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric",
        }),
    };

    // Render the document
    doc.render(placeholders);

    // Generate the processed document as a base64 string
    const base64Template = doc.getZip().generate({ type: "base64" });

    await Word.run(async (context) => {
      // Create the new document
      const newDoc = context.application.createDocument(base64Template);
      context.trackedObjects.add(newDoc);
      await context.sync();
    
      // Search for the placeholder in the new document
      const searchResults = newDoc.body.search("{{clientEmail}}");
      context.load(searchResults, "items");
      await context.sync();
    
      if (searchResults.items.length === 0) {
        console.error("Placeholder {{clientEmail}} not found in the document.");
        return;
      }
    
      // Replace the placeholder with a mailto hyperlink using HTML
      const mailtoHtml = `<a href="mailto:${clientEmail}">${clientEmail}</a>`;
      searchResults.items[0].insertHtml(mailtoHtml, Word.InsertLocation.replace);
      await context.sync();
    
      console.log("Replaced {{clientEmail}} with a mailto hyperlink using HTML.");

      newDoc.open();
      context.trackedObjects.remove(newDoc);
    });
  } catch (error) {
    console.error("Error creating contract:", error);
  }
}

export async function createReceipt() {
  try {
    const {
      clientName,
      clientLanguage,
      lawyerId,
      depositAmount,
      paymentMethod,
    } = formState;

    // Basic input validation
    if (!clientName || !lawyerId || !depositAmount) {
      console.error("One or more inputs are missing.");
      return;
    }

    const language = clientLanguage === "Français" ? "fr" : "en";

    // Load the DOCX template as a binary string
    const templateBinary = await loadDocxTemplate(language, "receipt");

    // Initialize PizZip with the template binary
    const zip = new PizZip(templateBinary);

    // Initialize Docxtemplater
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Define placeholders
    const placeholders = {
      user: "Michel Assi-Bodje",
      reason: "{}",
      clientName,
      paymentMethod,
      depositAmount: Number(depositAmount).toFixed(2),
      lawyerName: getLawyer(lawyerId).name,
      date: new Date().toLocaleDateString(language === "fr" ? "fr-CA" : "en-US", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
      }),
    };

    // Render the document
    doc.render(placeholders);

    // Generate the processed document as a base64 string
    const base64Template = doc.getZip().generate({ type: "base64" });

    await Word.run(async (context) => {
      // Create the new document
      const newDoc = context.application.createDocument(base64Template);
      context.trackedObjects.add(newDoc);
      await context.sync();
      newDoc.open();
      context.trackedObjects.remove(newDoc);
    });
  } catch (error) {
    console.error("Error creating receipt:", error);
  }
}
