/* ATLAS OS configuration.
   msClientId: the Application (client) ID of the "ATLAS OS" app registration in Microsoft Entra
   (set during the admin sign-in step). Until it is set, ATLAS runs on sample data. */
window.ATLAS_CONFIG = {
  msClientId: "5921eced-55f6-4bce-a673-27ad87d5fb05",
  msTenant: "4ffe99f1-8ad7-444f-8543-dd742eaf28c8",
  scopes: ["User.Read", "Files.Read"],
  // OneDrive file names ATLAS looks for when you connect your account
  records: {
    jamaClientAccounts: "Jama Client Accounts.xlsx"
  }
};
