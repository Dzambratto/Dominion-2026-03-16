export default function handler(req, res) {
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const tenantId = process.env.MICROSOFT_TENANT_ID || 'common';
  if (!clientId) {
    return res.status(500).json({ error: 'Microsoft OAuth not configured. Add MICROSOFT_CLIENT_ID to Vercel environment variables.' });
  }
  const userId = req.query.userId || '';
  // Determine the origin dynamically so this works from any Vercel preview URL
  const proto = req.headers['x-forwarded-proto'] || 'https';
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  const origin = `${proto}://${host}`;
  const redirectUri = `${origin}/api/auth/microsoft/callback`;
  // Encode userId + origin in state so callback can redirect back to the right domain
  const state = Buffer.from(JSON.stringify({ userId, origin })).toString('base64url');
  const params = new URLSearchParams({
    client_id: clientId,
    redirect_uri: redirectUri,
    response_type: 'code',
    scope: 'openid email profile offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadBasic',
    state,
    response_mode: 'query',
  });
  return res.redirect(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params.toString()}`);
}
