// Front-end sign-in config for the supervisor app. These are NOT secrets.
// After you create the "Supervisor Timesheet" app registration, paste its
// Application (client) ID below. It must equal the SUP_CLIENT_ID env var on
// Netlify (the token audience the backend checks). The tenant is your J & D
// McLennan one.
window.SUP_CONFIG = {
  clientId: "7a1c2b98-bb07-4748-adfd-6357da796d6d",
  tenantId: "3efd78a4-4c46-434e-b653-4d0b65d18caa"
};
