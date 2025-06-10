// Hub for the HTA application
attachEventListeners();


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
          if (Office.context.host === Office.HostType.Word) {
            populateContractTitles();
          }
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
        case ELEMENT_IDS.userManualMenuBtn:
          // Open user manual in new tab
          openPopup({
            title: "User Manual",
            contentOrFile: "./user-manual.html",
            isFile: true,
            position: "bottom-right",
          });
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
        case ELEMENT_IDS.wordContractSubmitBtn:
          createContract();
          break;
        case ELEMENT_IDS.wordReceiptSubmitBtn:
          createReceipt();
          break;
        default:
          break;
      }
    });
  });
}

/**
 * Prepares a reply email.
 * @async
 */
async function sendReply() {
  await createEmail("reply");
}

/**
 * Prepares a contract email.
 * @async
 */
async function sendContract() {
  await createEmail("contract");
}

/**
 * Prepares a confirmation email.
 * @async
 */
async function sendConfirmation() {
  await createEmail();
}

/**
 * Returns valid time slots for the appointment within the accepted range.
 * @param {string} scheduleMode - The schedule mode ("auto" or "manual").
 * @async
 * @returns {Promise<Array<{ start: Date, end: Date, location:string }>>} An array of valid time slots.
 */
async function findSlots(scheduleMode) {
  try {
    const lawyer = getLawyer(formState.lawyerId);
    const location = formState.location;

    // === Manual Scheduling Mode ===
    if (scheduleMode === 'manual') {
      const manualDate = document.getElementById(ELEMENT_IDS.manualDate)?.value;
      const manualTime = document.getElementById(ELEMENT_IDS.manualTime)?.value;

      if (!manualDate || !manualTime) {
        throw new Error("Please select a valid date and time for manual scheduling.");
      }

      const start = new Date(`${manualDate}T${manualTime}`);
      const end = new Date(start.getTime() + 60 * 60 * 1000);
      const selectedSlot = { start, end, location };

      // Validate manually selected slot
      if (!isValidSlot(lawyer.id, selectedSlot, [])) { // Pass an empty array for events
        throw new Error("The selected time slot is not valid.");
      }

      return [selectedSlot];
    }

    // === Auto-Scheduling Mode ===
    const events = await fetchCalendarEvents();
    const rawSlots = generateSlots(lawyer, location, events);

    // Validate each slot
    const validated = await Promise.all(
      rawSlots.map(async slot => ({
        slot,
        valid: isValidSlot(lawyer.id, slot, events)
      }))
    );

    const validSlots = validated.filter(r => r.valid).map(r => r.slot);

    if (validSlots.length === 0) {
      throw new Error('No available slots found in the next 2 weeks.');
    }

    console.log('Valid slots:', validSlots);
    return validSlots;
  } catch (error) {
    console.error('Error finding schedule slot:', error);
    showErrorModal(error.message);
  }
}

/**
 * Schedules an appointment based on the selected inputs.
 * @async
 * @throws {Error} if the inputs are invalid or if no available slots are found.
 */
async function scheduleAppointment() {
  try {
    // Show loading spinner
    showLoading(true);

    if (!isValidInputs()) {
      throw new Error("Invalid inputs.");
    }

    const scheduleMode = document.getElementById(ELEMENT_IDS.scheduleMode)?.value;

    // Fetch valid slots based on the schedule mode
    const validSlots = await findSlots(scheduleMode);

    if (!validSlots || validSlots.length === 0) {
      throw new Error("No available slots found.");
    }

    // Select the appropriate slot
    const now = new Date();
    const selectedSlot =
      scheduleMode === "manual"
        ? validSlots[0] // Manual mode returns the chosen slot if valid
        : validSlots.find(slot => !isSameDay(slot.start, now) || slot.start > now) || validSlots[0]; // Auto mode logic: select the next available slot

    console.log("Selected appointment slot:", selectedSlot.start);

    // Display available slots in a popup for auto-scheduling
    if (scheduleMode === "auto") {
      popupAvailableSlots(validSlots, selectedSlot);
    }

    // Create the appointment in the lawyer's calendar
    await createMeeting(selectedSlot);
    console.log("Scheduled appointment successfully.");

    // Prepare confirmation email
    prepareConfirmation(selectedSlot);

  } catch (error) {
    console.error("Scheduling Error:", error);
  } finally {
    // Hide loading spinner
    showLoading(false);
  }
}