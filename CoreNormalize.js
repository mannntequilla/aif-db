function normalizeReferralSource_(leadReferralSource, leadType) {
  const referral = String(leadReferralSource || '').trim();

  return referral;
}

function normalizeText_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeScheduledEventType_(eventType) {
  return String(firstNonEmpty_(eventType || ''))
    .trim()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}
