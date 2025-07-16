/** The form state object to store user inputs */
class FormState {
  constructor() {
    this.reset();
  }

  /**
   * Resets all form-related properties to their default initial values.
   * This includes clearing text fields, unchecking checkboxes, and setting
   * file and selection fields to their initial states.
   */
  reset() {
    this.matterId = "";
    this.lawyerId = "";
    this.location = "";
    this.clientTitle = "";
    this.clientName = "";
    this.clientPhone = "";
    this.clientEmail = "";
    this.clientLanguage = "";
    this.caseType = "";
    this.appointmentDate = null;
    this.appointmentTime = null;
    this.isFirstConsultation = false;
    this.isRefBarreau = false;
    this.isExistingClient = false;
    this.isPaymentMade = false;
    this.depositAmount = 0;
    this.paymentMethod = "";
    this.contractTitle = "";
    this.receiptReason = "";
    this.notes = "";
  }

  /**
   * Updates the specified property of the form state with a new value.
   * @param {string} property - The name of the property to update.
   * @param {*} value - The new value to assign to the specified property.
   */
  update(property, value) {
    this[property] = value;
  }
}

export const formState = new FormState();