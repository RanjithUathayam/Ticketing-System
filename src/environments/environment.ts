// src/environments/environment.ts  (Development)
export const environment = {
  production: false,
  apiUrl: "https://vms.uathayam.in:4300/TICKETING-API/api",
//   apiUrl: 'http://localhost:3001/api',
//   hwLabelPrintMode: 'server' as 'server' | 'local-agent' | 'qz-tray',
  hwLabelPrintMode: 'qz-tray',
  hwLabelAgentUrl: 'http://127.0.0.1:17951',
  hwLabelPrinterName: 'TSC TTP-244 Pro',
};
