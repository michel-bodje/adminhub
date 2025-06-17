/**
 * @module constants
 * @description This module contains constants used throughout the application,
 * such as HTML element IDs for various components, including pages, buttons,
 * form inputs, and dropdowns.
 * This avoids hardcoding ID strings in multiple places and makes it easier to maintain the code.
 * This way, if you need to change an ID, you only have to do it here.
 */
export const ELEMENT_IDS = {
    // Page IDs
    menuPage: "menu-page",
    outlookMenuPage: "outlook-menu-page",
    wordMenuPage: "word-menu-page",
    pclawMenuPage: "pclaw-menu-page",

    schedulePage: "schedule-page",
    confirmationPage: "confirmation-page",
    replyPage: "reply-page",
    contractPage: "contract-page",
    wordContractPage: "word-contract-page",
    wordReceiptPage: "word-receipt-page",

    // Button IDs
    backBtn: "back-btn",

    outlookMenuBtn: "outlook-menu-btn",
    wordMenuBtn: "word-menu-btn",
    pclawMenuBtn: "pclaw-menu-btn",

    scheduleMenuBtn: "schedule-menu-btn",
    confirmMenuBtn: "conf-menu-btn",
    contractMenuBtn: "contract-menu-btn",
    replyMenuBtn: "reply-menu-btn",
    userManualMenuBtn: "user-manual-btn",
    wordContractMenuBtn: "word-contract-menu-btn",
    wordReceiptMenuBtn: "word-receipt-menu-btn",

    scheduleSubmitBtn: "schedule-appointment-btn",
    confirmSubmitBtn: "send-confirmation-btn",
    contractSubmitBtn: "send-contract-btn",
    replySubmitBtn: "send-reply-btn",
    wordContractSubmitBtn: "generate-word-contract-btn",
    wordReceiptSubmitBtn: "generate-word-receipt-btn",

    // Form input IDs
    scheduleLawyerId: "schedule-lawyer-id",
    confLawyerId: "conf-lawyer-id",
    contractLawyerId: "contract-lawyer-id",
    replyLawyerId: "reply-lawyer-id",
    wordContractLawyerId: "word-contract-lawyer-id",
    wordReceiptLawyerId: "word-receipt-lawyer-id",

    scheduleLocation: "schedule-location",
    confLocation: "conf-location",

    scheduleClientName: "schedule-client-name",
    wordContractClientName: "word-contract-client-name",
    wordReceiptClientName: "word-receipt-client-name",

    scheduleClientPhone: "schedule-client-phone",
    scheduleClientTitle: "schedule-client-title",

    scheduleClientEmail: "schedule-client-email",
    confClientEmail: "conf-client-email",
    contractClientEmail: "contract-client-email",
    replyClientEmail: "reply-client-email",
    wordContractClientEmail: "word-contract-client-email",

    scheduleClientLanguage: "schedule-client-language",
    confClientLanguage: "conf-client-language",
    contractClientLanguage: "contract-client-language",
    replyClientLanguage: "reply-client-language",
    wordContractClientLanguage: "word-contract-client-language",
    wordReceiptClientLanguage: "word-receipt-client-language",

    confDate: "conf-date",
    confTime: "conf-time",
    manualDate: "manual-date",
    manualTime: "manual-time",
    scheduleMode: "schedule-mode",
    
    emailContractDeposit: "email-contract-deposit",
    wordContractDeposit: "word-contract-deposit",
    wordReceiptDeposit: "word-receipt-deposit",
    wordContractTitle: "word-contract-title",
    customContractTitle: "custom-contract-title",
    wordReceiptReason: "word-receipt-reason",
    customReceiptReason: "custom-receipt-reason",

    refBarreau: "ref-barreau",
    existingClient: "existing-client",

    scheduleFirstConsultation: "schedule-first-consultation",
    confFirstConsultation: "conf-first-consultation",

    paymentMade: "payment-made",
    schedulePaymentMethod: "payment-method",
    receiptPaymentMethod: "receipt-payment-method",
    paymentOptionsContainer: "payment-options-container",

    // Case details
    caseType: "case-type",
    caseDetails: "case-details",

    spouseName: "spouse-name",
    deceasedName: "deceased-name",
    executorName: "executor-name",
    employerName: "employer-name",
    businessName: "business-name",
    mandateDetails: "mandate-details",
    otherPartyName: "other-party-name",
    commonField: "common-field",

    conflictSearchDoneDivorce: "conflict-search-done-divorce",
    conflictSearchDoneEstate: "conflict-search-done-estate",

    // Notes
    notes: "schedule-notes",
    notesContainer: "notes-container",

    // Misc
    loadingOverlay: "loading-overlay",
};