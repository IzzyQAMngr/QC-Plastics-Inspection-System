/*************************************************************
 * START-UP VERIFICATION — METALS — thin wrappers over the shared engine in
 * StartUpVerification.gs, fixed to department='Metals'. Deviation approval/PFA
 * sign-off share the same Deviation Approvals inbox as Plastics (see
 * getDeviationApprovalFormData / approveDeviation / signPfa in StartUpVerification.gs).
 *************************************************************/

/** Called by MetalsStartUpVerificationView.html on load. */
function getStartUpVerificationFormDataMetals() { return getStartUpVerificationFormData('Metals'); }

/** Public wrapper — same payload shape as saveStartUpVerification. */
function saveStartUpVerificationMetals(payload) { return saveStartUpVerification_(payload, 'Metals'); }
