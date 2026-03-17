export default async function handler(req, res) {
  const { code, state: rawState, error } = req.query;

  // Decode state to get userId and origin
  let userId = '';
  let origin = 'https://getdominiontech.com';
  try {
    const decoded = JSON.parse(Buffer.from(rawState || '', 'base64url').toString());
    userId = decoded.userId || '';
    origin = decoded.origin || origin;
  } catch {
    userId = rawState || '';
  }

  if (error) {
    return res.redirect(`${origin}/?oauth_error=${encodeURIComponent(error)}`);
  }
  if (!code) {
    return res.redirect(`${origin}/?oauth_error=no_code`);
  }

  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const tenantId = process.env.MICROSOFT_TENANT_ID || 'common';
  if (!clientId || !clientSecret) {
    return res.redirect(`${origin}/?oauth_error=not_configured`);
  }

  const proto = req.headers['x-forwarded-proto'] || 'https';
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  const redirectUri = `${proto}://${host}/api/auth/microsoft/callback`;

  try {
    // Exchange code for tokens
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
      }),
    });
    const tokens = await tokenRes.json();
    if (tokens.error) {
      return res.redirect(`${origin}/?oauth_error=${encodeURIComponent(tokens.error)}`);
    }

    // Get user email from Microsoft Graph
    const profileRes = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${tokens.access_token}` },
    });
    const profile = await profileRes.json();
    const email = profile.mail || profile.userPrincipalName || '';

    // Store tokens in a secure cookie keyed by userId
    const cookieKey = `ms_tokens_${(userId || 'anon').replace(/[^a-zA-Z0-9]/g, '_')}`.slice(0, 64);
    const tokenPayload = {
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expiry: Date.now() + (tokens.expires_in || 3600) * 1000,
    };
    const encoded = Buffer.from(JSON.stringify(tokenPayload)).toString('base64');
    res.setHeader('Set-Cookie', [
      `${cookieKey}=${encoded}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=2592000`,
    ]);

    return res.redirect(
      `${origin}/?oauth_success=microsoft&email=${encodeURIComponent(email)}&userId=${encodeURIComponent(userId)}`
    );
  } catch (err) {
    console.error('Microsoft OAuth callback error:', err);
    return res.redirect(`${origin}/?oauth_error=server_error`);
  }
}
