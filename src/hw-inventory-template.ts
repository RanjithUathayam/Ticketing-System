// src/hw-inventory-template.ts
//
// Column list mirrors the organization's "asset_import_template.csv" exactly
// (name, spelling, and order). Used to export Hardware Inventory data in the
// same layout so exported files can round-trip through that template.
export const HW_TEMPLATE_COLUMNS: string[] = [
  'Company', 'Commissioned Date', 'Asset Type', 'Make', 'Model', 'Serial No',
  'Host Name', 'Warranty', 'Expiry Date', 'PO Number', 'Vendor', 'Invoice No',
  'Value', 'OS', 'OS Edition', 'OS Serial Number', 'MS Office Version',
  'MS Office Edition', 'Mac ID', 'Wi-Fi Mac', 'Processor', 'RAM Size',
  'RAM Type', 'Storage Size', 'Storage Serial No', 'Monitor', 'Keyboard',
  'Mouse', 'Allocated To', 'Allocated Name', 'Allocation Date',
  'Allocated Dept', 'Allocated Location', 'Asset Category', 'Location Type',
  'Remarks', 'IMEI Number', 'DisplaySize', 'Management IP Address',
  'WAN IP Address', 'LAN IP Address', 'Number of Ports', 'Uplink Ports',
  'Stack Member', 'VPN Type', 'HA Mode', 'Subscription Details',
  'SSID Details', 'Controller Type', 'Band Support', 'Mount Type',
  'Printer Type', 'Printer Consumable', 'Printing Area size', 'Duplex Support',
  'DPI Resolution', 'Printing Speed', 'Cutter Availability',
  'Connectivity Type', 'Scanner Type', 'Supported Type',
  'Bluetooth Pairing ID', 'MDM Enrolled', 'MDM Policy', 'Resolution',
  'Lens Type', 'IR Range', 'PoE Support', 'Surveillance Recorder IP',
  'Connected Cameras', 'Recording Days', 'Rack Number', 'Rack Position',
  'Asset Ref', 'Capacity in kVA', 'IMEI Number 2', 'Battery Capacity',
  'Battery Qty',
];
