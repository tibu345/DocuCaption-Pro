export function friendlyErrorMessage(error: unknown, fallback = 'Something went wrong. Please try again.'): string {
  const raw = error instanceof Error ? error.message : typeof error === 'string' ? error : fallback;
  const message = stripHtml(raw).replace(/\s+/g, ' ').trim();
  const lower = message.toLowerCase();

  if (/email rate limit|rate limit.*email|2 emails per hour/.test(lower)) {
    return 'We have sent several account emails recently. Please wait before requesting another email, or continue with Google.';
  }
  if (/invalid login credentials|invalid credentials/.test(lower)) {
    return 'Email or password is incorrect.';
  }
  if (/email not confirmed|confirm.*email/.test(lower)) {
    return 'Please confirm your email before signing in.';
  }
  if (/missing authentication token|invalid or expired authentication token|jwt/.test(lower)) {
    return 'Your session expired. Please sign in again.';
  }
  if (/failed to fetch|econnrefused|backend unavailable|service.*starting|networkerror/.test(lower)) {
    return 'The document service is not reachable. Start the backend, then try again.';
  }
  if (/caption request is too large|payloadtoolarge|request entity too large/.test(lower)) {
    return 'The caption request is too large. Try fewer figures at once or use smaller images.';
  }
  if (/quota|resource_exhausted|billing|api key|permission|gemini|caption service failed/.test(lower)) {
    return 'Generated captions are temporarily unavailable. Fallback captions can still be edited before export.';
  }
  if (/no caption credits/.test(lower)) {
    return 'You have no caption credits remaining.';
  }
  if (/monthly document limit/.test(lower)) {
    return 'You have reached your monthly document limit.';
  }
  if (/monthly export limit/.test(lower)) {
    return 'You have reached your monthly export limit.';
  }

  return message || fallback;
}

function stripHtml(value: string): string {
  return value.replace(/<[^>]*>/g, ' ');
}
