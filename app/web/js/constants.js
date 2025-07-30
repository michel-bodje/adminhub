/**
 * @module constants
 * @description This module contains constants used throughout the application,
 * such as HTML element IDs for various components, including pages, buttons,
 * form inputs, and dropdowns.
 * This avoids hardcoding ID strings in multiple places and makes it easier to maintain the code.
 * This way, if you need to change an ID, you only have to do it here.
 */
export const ELEMENT_IDS = {
    // --- General/Misc ---
    loadingOverlay: "loading-overlay",

    // --- Main Menu Pages ---
    menuPage: "menu-page",
    outlookMenuPage: "outlook-menu-page",
    wordMenuPage: "word-menu-page",
    pclawMenuPage: "pclaw-menu-page",

    // --- Main Menu Buttons ---
    outlookMenuBtn: "outlook-menu-btn",
    wordMenuBtn: "word-menu-btn",
    pclawMenuBtn: "pclaw-menu-btn",

    // --- Outlook Submenu Buttons ---
    scheduleMenuBtn: "schedule-menu-btn",
    confirmMenuBtn: "conf-menu-btn",
    replyMenuBtn: "reply-menu-btn",
    followupMenuBtn: "followup-menu-btn",
    reviewMenuBtn: "review-menu-btn",
    contractMenuBtn: "contract-menu-btn",

    // --- Word Submenu Buttons ---
    wordContractMenuBtn: "word-contract-menu-btn",
    wordReceiptMenuBtn: "word-receipt-menu-btn",

    // --- PCLaw Submenu Buttons ---
    pclawMatterMenuBtn: "new-matter-menu-btn",
    pclawCloseMatterMenuBtn: "close-matter-menu-btn",
    pclawBillMatterMenuBtn: "bill-matter-menu-btn",

    // --- Common Buttons ---
    backBtn: "back-btn",

    // --- Schedule Appointment Page ---
    schedulePage: "schedule-page",
    scheduleClientName: "schedule-client-name",
    scheduleClientPhone: "schedule-client-phone",
    scheduleClientEmail: "schedule-client-email",
    scheduleClientLanguage: "schedule-client-language",
    scheduleClientTitle: "schedule-client-title",
    scheduleCaseType: "case-type",
    scheduleLawyerId: "schedule-lawyer-id",
    scheduleLocation: "schedule-location",
    caseDetails: "case-details",
    spouseName: "spouse-name",
    conflictSearchDoneDivorce: "conflict-search-done-divorce",
    deceasedName: "deceased-name",
    executorName: "executor-name",
    conflictSearchDoneEstate: "conflict-search-done-estate",
    employerName: "employer-name",
    otherPartyName: "other-party-name",
    mandateDetails: "mandate-details",
    businessName: "business-name",
    commonField: "common-field",
    scheduleMode: "schedule-mode",
    manualDate: "manual-date",
    manualTime: "manual-time",
    isExistingClient: "existing-client",
    isRefBarreau: "ref-barreau",
    scheduleFirstConsultation: "schedule-first-consultation",
    isPaymentMade: "payment-made",
    paymentOptionsContainer: "payment-options-container",
    schedulePaymentMethod: "payment-method",
    notes: "schedule-notes",
    notesContainer: "notes-container",
    alsoEmail: "also-email",
    alsoPclaw: "also-pclaw",
    scheduleSubmitBtn: "schedule-appointment-btn",

    // --- Email Confirmation Page ---
    confirmationPage: "confirmation-page",
    confClientEmail: "conf-client-email",
    confClientLanguage: "conf-client-language",
    confLawyerId: "conf-lawyer-id",
    confLocation: "conf-location",
    confDate: "conf-date",
    confTime: "conf-time",
    confFirstConsultation: "conf-first-consultation",
    confirmSubmitBtn: "send-confirmation-btn",

    // --- Email Reply Page ---
    replyPage: "reply-page",
    replyClientEmail: "reply-client-email",
    replyClientLanguage: "reply-client-language",
    replyLawyerId: "reply-lawyer-id",
    replySubmitBtn: "send-reply-btn",

    // --- Email Follow-up Page ---
    followupPage: "followup-page",
    followupClientEmail: "followup-client-email",
    followupClientLanguage: "followup-client-language",
    followupSubmitBtn: "send-followup-btn",

    // --- Email Review Page ---
    reviewPage: "review-page",
    reviewClientEmail: "review-client-email",
    reviewClientLanguage: "review-client-language",
    reviewSubmitBtn: "send-review-btn",

    // --- Email Contract Page ---
    contractPage: "contract-page",
    contractClientEmail: "contract-client-email",
    contractClientLanguage: "contract-client-language",
    contractLawyerId: "contract-lawyer-id",
    emailContractDeposit: "email-contract-deposit",
    contractSubmitBtn: "send-contract-btn",

    // --- Word Contract Page ---
    wordContractPage: "word-contract-page",
    wordContractClientName: "word-contract-client-name",
    wordContractClientEmail: "word-contract-client-email",
    wordContractLawyerId: "word-contract-lawyer-id",
    wordContractDeposit: "word-contract-deposit",
    wordContractClientLanguage: "word-contract-client-language",
    wordContractTitle: "word-contract-title",
    customContractTitle: "custom-contract-title",
    alsoContract: "also-contract",
    wordContractSubmitBtn: "generate-word-contract-btn",

    // --- Word Receipt Page ---
    wordReceiptPage: "word-receipt-page",
    wordReceiptClientName: "word-receipt-client-name",
    wordReceiptClientLanguage: "word-receipt-client-language",
    wordReceiptDeposit: "word-receipt-deposit",
    receiptPaymentMethod: "receipt-payment-method",
    wordReceiptReason: "word-receipt-reason",
    customReceiptReason: "custom-receipt-reason",
    wordReceiptLawyerId: "word-receipt-lawyer-id",
    wordReceiptSubmitBtn: "generate-word-receipt-btn",

    // --- PCLaw New Matter Page ---
    pclawMatterPage: "pclaw-matter-page",
    matterClientName: "matter-client-name",
    matterClientPhone: "matter-client-phone",
    matterClientEmail: "matter-client-email",
    matterClientLanguage: "matter-client-language",
    matterClientTitle: "matter-client-title",
    matterCaseType: "matter-case-type",
    matterLawyerId: "matter-lawyer-id",
    matterCaseDetails: "matter-case-details",
    matterSpouseName: "matter-spouse-name",
    matterConflictSearchDoneDivorce: "matter-conflict-search-done-divorce",
    matterDeceasedName: "matter-deceased-name",
    matterExecutorName: "matter-executor-name",
    matterConflictSearchDoneEstate: "matter-conflict-search-done-estate",
    matterEmployerName: "matter-employer-name",
    matterOtherPartyName: "matter-other-party-name",
    matterMandateDetails: "matter-mandate-details",
    matterBusinessName: "matter-business-name",
    matterCommonField: "matter-common-field",
    matterExistingClient: "matter-existing-client",
    matterRefBarreau: "matter-ref-barreau",
    matterFirstConsultation: "matter-first-consultation",
    pclawMatterSubmitBtn: "matter-btn",

    // --- PCLaw Close Matter Page ---
    pclawCloseMatterPage: "pclaw-close-page",
    closeMatterId: "close-matter-id",
    pclawCloseMatterSubmitBtn: "close-matter-btn",

    // --- PCLaw Bill Matter Page ---
    pclawBillMatterPage: "pclaw-bill-page",
    billMatterId: "bill-matter-id",
    pclawBillMatterSubmitBtn: "bill-matter-btn",

    // --- PCLaw Time Entry Page ---
    timeEntriesMenuBtn: "time-entries-menu-btn",
    timeEntriesPage: "pclaw-time-entries-page",
    timeEntriesSubmitBtn: "time-entries-btn",
    timeEntriesLawyer: "time-entries-lawyer-id",
    timeEntriesFile: "time-entries-file",
    selectFileBtn: "select-file-btn",
    selectedFilePath: "selected-file-path",
};

export const MENU_ROUTES = {
  [ELEMENT_IDS.outlookMenuBtn]: ELEMENT_IDS.outlookMenuPage,
  [ELEMENT_IDS.wordMenuBtn]: ELEMENT_IDS.wordMenuPage,
  [ELEMENT_IDS.pclawMenuBtn]: ELEMENT_IDS.pclawMenuPage,
  [ELEMENT_IDS.scheduleMenuBtn]: ELEMENT_IDS.schedulePage,
  [ELEMENT_IDS.confirmMenuBtn]: ELEMENT_IDS.confirmationPage,
  [ELEMENT_IDS.contractMenuBtn]: ELEMENT_IDS.contractPage,
  [ELEMENT_IDS.replyMenuBtn]: ELEMENT_IDS.replyPage,
  [ELEMENT_IDS.followupMenuBtn]: ELEMENT_IDS.followupPage,
  [ELEMENT_IDS.reviewMenuBtn]: ELEMENT_IDS.reviewPage,
  [ELEMENT_IDS.wordContractMenuBtn]: ELEMENT_IDS.wordContractPage,
  [ELEMENT_IDS.wordReceiptMenuBtn]: ELEMENT_IDS.wordReceiptPage,
  [ELEMENT_IDS.pclawMatterMenuBtn]: ELEMENT_IDS.pclawMatterPage,
  [ELEMENT_IDS.pclawCloseMatterMenuBtn]: ELEMENT_IDS.pclawCloseMatterPage,
  [ELEMENT_IDS.pclawBillMatterMenuBtn]: ELEMENT_IDS.pclawBillMatterPage,
  [ELEMENT_IDS.timeEntriesMenuBtn]: ELEMENT_IDS.timeEntriesPage
};
