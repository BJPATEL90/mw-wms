function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    SHEET_ID: props.getProperty('SHEET_ID'),
    GOOGLE_CLIENT_ID: props.getProperty('GOOGLE_CLIENT_ID'),
    ALLOWED_DOMAIN: props.getProperty('ALLOWED_DOMAIN') || 'mosaicwellness.in'
  };
}

function getSS_() {
  const cfg = getConfig_();
  if (!cfg.SHEET_ID) throw new Error('SHEET_ID missing in Script Properties');
  return SpreadsheetApp.openById(cfg.SHEET_ID);
}
