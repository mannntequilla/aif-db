function normalizeReferralSource_(leadReferralSource, leadType) {
  const referral = String(leadReferralSource || '').trim();
  const type = String(leadType || '').trim();

  if (!referral && type === 'Existing Client') {
    return 'Existing Client';
  }

  return referral;
}
