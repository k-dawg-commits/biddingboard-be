const { google } = require('googleapis');
const parseBidInfo = require("../services/parseBidInfo.js")

function parseAiResponse(response) {
    return JSON.parse(response.replace(/```json/g, "").replace(/```/g, ""))
}

async function getUnreadMails(auth, maxResults = 10) {
    const gmail = google.gmail({ version: 'v1', auth });

    const res = await gmail.users.messages.list({
        userId: 'me',
        q: 'is:unread',
        maxResults
    });

    const messages = res.data.messages || [];
    const decodeBase64 = (str) => {
        const cleaned = str.replace(/-/g, '+').replace(/_/g, '/');
        return Buffer.from(cleaned, 'base64').toString('utf-8');
    };

    const extractHeader = (headers, name) => {
        const header = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
        return header ? header.value : '';
    };

    const emails = [];

    let index = 0
    for (const message of messages) {
        index++
        const mailContent = await gmail.users.messages.get({
            userId: 'me',
            id: message.id,
            format: 'full'
        });

        const payload = mailContent.data.payload;
        const headers = payload.headers;

        const fromHeader = extractHeader(headers, 'From');
        const subject = extractHeader(headers, 'Subject');
        const date = extractHeader(headers, 'Date').substring(0, 25);

        const from = fromHeader.includes('<')
            ? fromHeader.substring(fromHeader.indexOf('<') + 1, fromHeader.length - 1)
            : fromHeader;

        let body = '';

        if (payload.body && payload.body.data) {
            body = decodeBase64(payload.body.data).trim();
        } else if (payload.parts) {
            const part = payload.parts.find(part =>
                part.mimeType === 'text/plain' || part.mimeType === 'text/html'
            );
            if (part && part.body && part.body.data) {
                body = decodeBase64(part.body.data).trim();
            }
        }

        try {
            const response = await parseBidInfo(body);

            let parsedInfo;
            try {
                const parsed = parseAiResponse(response);
                parsedInfo = parsed?.parsedInfo;
            } catch (err) {
                console.error(`Failed to parse AI response for email ${index}:`, err.message);
            }

            if (parsedInfo) {
                emails.push({ id: index, body, from, subject, date, parsedInfo });
            } else {
                console.warn(`No parsedInfo for email ${index}.`);
            }
        } catch (err) {
            console.error(`Error processing email ${index}:`, err.message);
        }

    }

    return emails;
}

module.exports = getUnreadMails;
