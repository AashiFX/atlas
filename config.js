/* ATLAS OS configuration (no business data here — this file is public).
   Sign-in is limited to the AashiFX Microsoft tenant (single-tenant app registration). Read-only: User.Read + Files.Read. */
window.ATLAS_CONFIG = {
  msClientId: "5921eced-55f6-4bce-a673-27ad87d5fb05",
  msTenant: "4ffe99f1-8ad7-444f-8543-dd742eaf28c8",
  scopes: ["User.Read", "Files.Read"],
  // How ATLAS finds each workbook in OneDrive: "path" is tried first (when the file sits in its final folder),
  // then each name in "names" is searched and the most recently saved exact match is used.
  records: {
    jamaClientAccounts: { names: ["Jama Client Accounts.xlsx"] },
    masAccounts: { path: "AashiFX/ATLAS Hub/MAS/MAS Accounts/MAS Accounts.xlsx", names: ["MAS Accounts.xlsx", "MAS Accounts v0.4 (script-ready).xlsx"] },
    monthlyFinance: { path: "AashiFX/ATLAS Hub/Jama Star/Monthly Finance Report/Monthly Finance Report 2026.xlsx", names: ["Monthly Finance Report 2026.xlsx", "Monthly Finance Report v0.1 (script-ready).xlsx"] }
  },
  hubSheet: "Hub Status",
  hubRange: "A4:C40"
};
