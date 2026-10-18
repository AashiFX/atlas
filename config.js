/* ATLAS OS configuration.
   msClientId: the Application (client) ID of the "ATLAS OS" app registration in Microsoft Entra
   (set during the admin sign-in step). Until it is set, ATLAS runs on sample data. */
window.ATLAS_CONFIG = {
  msClientId: "",
  msTenant: "organizations",
  scopes: ["User.Read", "Files.Read"],
  // OneDrive file names ATLAS looks for when you connect your account
  records: {
    jamaClientAccounts: "Jama Client Accounts.xlsx"
  }
};
