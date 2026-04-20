function normalizeReferralSource_(leadReferralSource, leadType) {
  const referral = String(leadReferralSource || '').trim();
  const type = String(leadType || '').trim();

  if (!referral && type === 'Existing Client') {
    return 'Existing Client';
  }

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
