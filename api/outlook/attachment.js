/**
 * /api/outlook/attachment
 *
 * Downloads a specific Outlook message attachment via Microsoft Graph API.
 * Returns the attachment as base64 for AI extraction.
 */

function getTokensFromCookie(cookies, userId) {
  const cookieKey = `ms_tokens_${(userId || 'anon').replace(/[^a-zA-Z0-9]/g, '_')}`.slice(0, 64);
  const cookieStr = cookies || '';
  const match = cookieStr.match(new RegExp(`(?:^|;\\s*)${cookieKey}=([^;]+)`));
  if (!match) return null;
  try {
    return JSON.parse(Buffer.from(match[1], 'base64').toString());
  } catch {
    return null;
  }
}

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { userId, messageId, attachmentId } = req.query;
  if (!userId || !messageId || !attachmentId) {
    return res.status(400).json({ error: 'userId, messageId, and attachmentId are required' });
  }

  const tokenData = getTokensFromCookie(req.headers.cookie, userId);
  if (!tokenData) {
    return res.status(401).json({ error: 'not_connected' });
  }

  try {
    const attRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/${attachmentId}`,
      { headers: { Authorization: `Bearer ${tokenData.access_token}` } }
    );
    const att = await attRes.json();

    if (att.error) {
      return res.status(400).json({ error: att.error.code, message: att.error.message });
    }

    // contentBytes is already base64 in Graph API response
    return res.status(200).json({
      name: att.name,
      contentType: att.contentType,
      data: att.contentBytes,
    });
  } catch (err) {
    console.error('Outlook attachment error:', err);
    return res.status(500).json({ error: 'download_failed', message: err.message });
  }
}
