const Imap = require('node-imap');
const { simpleParser } = require('mailparser');

const DEFAULT_SCAN_LIMIT = 500;
const MAX_SCAN_LIMIT = 2000;

function getParams(req) {
    return req.method === 'GET' ? req.query : req.body || {};
}

function parseBoolean(value) {
    return value === true || value === 'true' || value === '1' || value === 1 || value === 'yes';
}

function parseKeywords(value) {
    if (Array.isArray(value)) {
        return value.map(item => String(item).trim()).filter(Boolean);
    }

    if (typeof value !== 'string') {
        return [];
    }

    const trimmed = value.trim();
    if (!trimmed) {
        return [];
    }

    try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) {
            return parsed.map(item => String(item).trim()).filter(Boolean);
        }
    } catch (error) {
        // Not JSON. Fall back to common separators.
    }

    return trimmed
        .split(/[\n,，]+/)
        .map(item => item.trim())
        .filter(Boolean);
}

function parseScanLimit(value) {
    const limit = Number.parseInt(value, 10);
    if (!Number.isFinite(limit) || limit <= 0) {
        return DEFAULT_SCAN_LIMIT;
    }
    return Math.min(limit, MAX_SCAN_LIMIT);
}

function normalizeMailboxForImap(mailbox) {
    if (!mailbox || mailbox === 'INBOX' || mailbox === 'inbox') {
        return 'INBOX';
    }

    if (mailbox === 'Junk' || mailbox === 'junk' || mailbox === 'junkemail') {
        return 'Junk';
    }

    return mailbox;
}

function normalizeMailboxForGraph(mailbox) {
    if (!mailbox || mailbox === 'INBOX' || mailbox === 'inbox') {
        return 'inbox';
    }

    if (mailbox === 'Junk' || mailbox === 'junk' || mailbox === 'junkemail') {
        return 'junkemail';
    }

    return mailbox;
}

function stripHtml(html) {
    if (!html) {
        return '';
    }

    return String(html)
        .replace(/<style[\s\S]*?<\/style>/gi, ' ')
        .replace(/<script[\s\S]*?<\/script>/gi, ' ')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function truncate(value, maxLength) {
    if (!value) {
        return '';
    }

    const text = String(value);
    if (text.length <= maxLength) {
        return text;
    }

    return `${text.slice(0, maxLength)}...`;
}

function buildSearchText(message) {
    return [
        message.send,
        message.subject,
        message.text,
        stripHtml(message.html)
    ].join('\n').toLowerCase();
}

function getMatchedKeywords(message, keywords, matchMode) {
    const searchText = buildSearchText(message);
    const matchedKeywords = keywords.filter(keyword => searchText.includes(keyword.toLowerCase()));

    if (matchMode === 'all') {
        return matchedKeywords.length === keywords.length ? matchedKeywords : [];
    }

    return matchedKeywords;
}

async function getAccessToken(refreshToken, clientId, scope) {
    const body = {
        client_id: clientId,
        grant_type: 'refresh_token',
        refresh_token: refreshToken
    };

    if (scope) {
        body.scope = scope;
    }

    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams(body).toString()
    });

    const responseText = await response.text();
    if (!response.ok) {
        throw new Error(`Token request failed: ${response.status}, response: ${responseText}`);
    }

    try {
        return JSON.parse(responseText);
    } catch (error) {
        throw new Error(`Failed to parse token response: ${error.message}, response: ${responseText}`);
    }
}

function hasGraphScope(scopeText, scopeName) {
    return String(scopeText || '')
        .split(/\s+/)
        .some(scope => scope === scopeName || scope.endsWith(`/${scopeName}`));
}

async function getGraphToken(refreshToken, clientId) {
    const data = await getAccessToken(refreshToken, clientId, 'https://graph.microsoft.com/.default');
    const scopeText = data.scope || '';
    const canRead = hasGraphScope(scopeText, 'Mail.Read') || hasGraphScope(scopeText, 'Mail.ReadWrite');

    return {
        accessToken: data.access_token,
        canRead,
        canDelete: hasGraphScope(scopeText, 'Mail.ReadWrite')
    };
}

function toGraphMessage(item, accountEmail, mailbox) {
    const from = item.from && item.from.emailAddress ? item.from.emailAddress : {};

    return {
        id: item.id,
        uid: item.id,
        accountEmail,
        mailbox,
        send: from.address || from.name || '',
        subject: item.subject || '',
        text: item.bodyPreview || stripHtml(item.body && item.body.content),
        html: item.body && item.body.content ? truncate(item.body.content, 10000) : '',
        date: item.receivedDateTime || item.createdDateTime || ''
    };
}

async function fetchGraphMessages(accessToken, mailbox, limit, accountEmail) {
    const messages = [];
    const pageSize = Math.min(100, limit);
    const encodedMailbox = encodeURIComponent(normalizeMailboxForGraph(mailbox));
    let nextUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/${encodedMailbox}/messages?$top=${pageSize}&$select=id,from,subject,bodyPreview,body,createdDateTime,receivedDateTime&$orderby=receivedDateTime desc`;

    while (nextUrl && messages.length < limit) {
        const response = await fetch(nextUrl, {
            method: 'GET',
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        const responseText = await response.text();
        if (!response.ok) {
            throw new Error(`Graph mail request failed: ${response.status}, response: ${responseText}`);
        }

        const data = JSON.parse(responseText);
        for (const item of data.value || []) {
            if (messages.length >= limit) {
                break;
            }
            messages.push(toGraphMessage(item, accountEmail, mailbox));
        }

        nextUrl = data['@odata.nextLink'] || '';
    }

    return messages;
}

async function deleteGraphMessages(accessToken, messages) {
    let deleted = 0;

    for (const message of messages) {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(message.id)}`, {
            method: 'DELETE',
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(`Graph delete failed for message ${message.id}: ${response.status}, response: ${responseText}`);
        }

        deleted += 1;
    }

    return deleted;
}

async function searchWithGraph(options) {
    let graphToken;
    try {
        graphToken = await getGraphToken(options.refreshToken, options.clientId);
    } catch (error) {
        console.warn('Graph token unavailable, falling back to IMAP:', error.message);
        return null;
    }

    if (!graphToken.canRead) {
        return null;
    }

    if (options.deleteMatches && !graphToken.canDelete) {
        console.warn('Graph token has no Mail.ReadWrite scope, falling back to IMAP for delete.');
        return null;
    }

    const scannedMessages = await fetchGraphMessages(
        graphToken.accessToken,
        options.mailbox,
        options.scanLimit,
        options.email
    );

    const matchedMessages = scannedMessages
        .map(message => ({
            ...message,
            matchedKeywords: getMatchedKeywords(message, options.keywords, options.matchMode)
        }))
        .filter(message => message.matchedKeywords.length > 0);

    if (options.deleteMatches && matchedMessages.length > 0) {
        const deleted = await deleteGraphMessages(graphToken.accessToken, matchedMessages);
        return {
            provider: 'graph',
            scanned: scannedMessages.length,
            deleted,
            messages: matchedMessages.map(message => ({ ...message, deleted: true }))
        };
    }

    return {
        provider: 'graph',
        scanned: scannedMessages.length,
        deleted: 0,
        messages: matchedMessages
    };
}

function generateAuthString(user, accessToken) {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

function connectImap(config) {
    return new Promise((resolve, reject) => {
        const imap = new Imap(config);

        imap.once('ready', () => resolve(imap));
        imap.once('error', reject);
        imap.connect();
    });
}

function openBox(imap, mailbox, readOnly) {
    return new Promise((resolve, reject) => {
        imap.openBox(mailbox, readOnly, (err, box) => {
            if (err) {
                reject(err);
                return;
            }
            resolve(box);
        });
    });
}

function searchAll(imap) {
    return new Promise((resolve, reject) => {
        imap.search(['ALL'], (err, results) => {
            if (err) {
                reject(err);
                return;
            }
            resolve(results || []);
        });
    });
}

function markAsDeleted(imap, uids) {
    return new Promise((resolve, reject) => {
        if (uids.length === 0) {
            resolve();
            return;
        }

        imap.addFlags(uids, ['\\Deleted'], err => {
            if (err) {
                reject(err);
                return;
            }
            resolve();
        });
    });
}

function expungeDeleted(imap) {
    return new Promise((resolve, reject) => {
        imap.expunge(err => {
            if (err) {
                reject(err);
                return;
            }
            resolve();
        });
    });
}

function fetchImapMessages(imap, uids, accountEmail, mailbox) {
    if (uids.length === 0) {
        return Promise.resolve([]);
    }

    return new Promise((resolve, reject) => {
        const messages = [];
        const parseTasks = [];
        const fetcher = imap.fetch(uids, { bodies: '', markSeen: false });

        fetcher.on('message', msg => {
            const chunks = [];
            let attrs = {};

            msg.on('body', stream => {
                stream.on('data', chunk => {
                    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
                });
                stream.once('error', reject);
            });

            msg.once('attributes', attributes => {
                attrs = attributes || {};
            });

            msg.once('end', () => {
                const raw = Buffer.concat(chunks);
                const task = simpleParser(raw, {
                    skipImageLinks: true,
                    skipTextToHtml: true,
                    skipTextLinks: true
                }).then(mail => {
                    const html = mail.html || '';
                    messages.push({
                        id: String(attrs.uid || ''),
                        uid: attrs.uid,
                        accountEmail,
                        mailbox,
                        send: mail.from ? mail.from.text : '',
                        subject: mail.subject || '',
                        text: truncate(mail.text || stripHtml(html), 10000),
                        html: truncate(html, 10000),
                        date: mail.date || ''
                    });
                });
                parseTasks.push(task);
            });
        });

        fetcher.once('error', reject);
        fetcher.once('end', async () => {
            try {
                await Promise.all(parseTasks);
                messages.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
                resolve(messages);
            } catch (error) {
                reject(error);
            }
        });
    });
}

async function searchWithImap(options) {
    const tokenData = await getAccessToken(options.refreshToken, options.clientId);
    const authString = generateAuthString(options.email, tokenData.access_token);
    const imap = await connectImap({
        user: options.email,
        xoauth2: authString,
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
        tlsOptions: {
            rejectUnauthorized: false
        }
    });

    try {
        const mailbox = normalizeMailboxForImap(options.mailbox);
        await openBox(imap, mailbox, !options.deleteMatches);

        const allUids = await searchAll(imap);
        const scanUids = allUids.slice(-options.scanLimit).reverse();
        const scannedMessages = await fetchImapMessages(imap, scanUids, options.email, mailbox);
        const matchedMessages = scannedMessages
            .map(message => ({
                ...message,
                matchedKeywords: getMatchedKeywords(message, options.keywords, options.matchMode)
            }))
            .filter(message => message.matchedKeywords.length > 0);

        if (options.deleteMatches && matchedMessages.length > 0) {
            const matchedUids = matchedMessages.map(message => message.uid).filter(Boolean);
            await markAsDeleted(imap, matchedUids);
            await expungeDeleted(imap);
            return {
                provider: 'imap',
                scanned: scannedMessages.length,
                deleted: matchedUids.length,
                messages: matchedMessages.map(message => ({ ...message, deleted: true }))
            };
        }

        return {
            provider: 'imap',
            scanned: scannedMessages.length,
            deleted: 0,
            messages: matchedMessages
        };
    } finally {
        imap.end();
    }
}

module.exports = async (req, res) => {
    if (req.method === 'OPTIONS') {
        res.status(204).end();
        return;
    }

    if (req.method !== 'GET' && req.method !== 'POST') {
        res.status(405).json({ error: 'Method not allowed' });
        return;
    }

    const params = getParams(req);
    const expectedPassword = process.env.PASSWORD;

    if (expectedPassword && params.password !== expectedPassword) {
        res.status(401).json({
            error: 'Authentication failed. Please provide valid credentials.'
        });
        return;
    }

    const refreshToken = params.refresh_token;
    const clientId = params.client_id;
    const email = params.email;
    const mailbox = params.mailbox || 'INBOX';
    const keywords = parseKeywords(params.keywords || params.keyword);
    const matchMode = params.match_mode === 'all' ? 'all' : 'any';
    const scanLimit = parseScanLimit(params.scan_limit || params.limit);
    const deleteMatches = parseBoolean(params.delete_matches || params.delete);

    if (!refreshToken || !clientId || !email) {
        res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, or email' });
        return;
    }

    if (keywords.length === 0) {
        res.status(400).json({ error: 'Missing required parameter: keywords' });
        return;
    }

    try {
        const options = {
            refreshToken,
            clientId,
            email,
            mailbox,
            keywords,
            matchMode,
            scanLimit,
            deleteMatches
        };

        const graphResult = await searchWithGraph(options);
        const result = graphResult || await searchWithImap(options);

        res.status(200).json({
            email,
            mailbox,
            provider: result.provider,
            keywords,
            match_mode: matchMode,
            scan_limit: scanLimit,
            scanned: result.scanned,
            matched: result.messages.length,
            deleted: result.deleted,
            delete_matches: deleteMatches,
            messages: result.messages
        });
    } catch (error) {
        console.error('Error searching emails:', error);
        res.status(error.statusCode || 500).json({
            error: 'Error searching emails',
            details: error.message
        });
    }
};
