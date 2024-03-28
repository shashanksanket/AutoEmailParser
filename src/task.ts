import { google, gmail_v1 } from 'googleapis';
import { Client } from '@microsoft/microsoft-graph-client';
import { OpenAI } from 'openai';

interface EmailReply {
    subject: string;
    body: string;
}

interface AnalyzedResponse {
    label: string;
    extractedMailContent: string;
    replyMail: EmailReply;
}

const analyzeEmail = async (emailContent: string): Promise<string> => {
    try {
        const openai = new OpenAI({
            apiKey: process.env.OPENAI_API_KEY || '',
            dangerouslyAllowBrowser: true,
        });
        const response = await openai.chat.completions.create({
            messages: [
                {
                    role: 'system',
                    content: `aAs a recruiter, I need to automate the process of analyzing incoming emails and generating appropriate replies based on their content. Here's the task:\n1. Analyze the content of the email and determine its relevance to job inquiries.\n2. Assign a label to each email based on its relevance: [Interested, Not Interested, More Information, Others].\n3. If the email is related to job inquiries, generate a reply email with appropriate subject and body.\n4. The reply email should be in the following JSON format: {"label":"", "extractedMailContent":"", "replyMail":{"subject":"", "body":""}} .Note:extractedMailContent should have the content of mail received not html tags \n5. If the email is not related to job inquiries, ignore it and assign the label as "other".\n6. Do not generate a body for emails labeled as "other".\n7. Ensure that you use proper name to refer any person you can get names of the candidate from there emailid or inside emailcontent refer me  that is any recuruiter with any random name if you can't get a name from mail contents \n\nPrompt: Analyze the given email content and generate a reply email accordingly. Consider yourself as a recruiter receiving job inquiries and tailor the reply based on the content of the email. \n This is the email Content
                    ${emailContent}`,
                },
            ],
            model: 'gpt-3.5-turbo',
        });
        return response.choices[0].message.content || "";
    } catch (error) {
        throw error;
    }
};

const sendReplyUsingGoogle = async (gmail: gmail_v1.Gmail, message: gmail_v1.Schema$Message, replyMail: EmailReply): Promise<void> => {
    try {
        const messageId = message.id ?? '';
        const res = await gmail.users.messages.get({
            userId: 'me',
            id: messageId,
            format: 'metadata',
            metadataHeaders: ['Subject', 'From'],
        });
        const from = res.data.payload?.headers?.find((header) => header.name == 'From')?.value || '';
        const replyTo = from.match(/<(.*)>/)?.[1] || '';
        const replySubject = replyMail.subject;
        const replyBody = replyMail.body.split('.').join('\n');

        const rawMessage = [
            `From: me`,
            `To: ${replyTo}`,
            `Subject: ${replySubject}`,
            `In-Reply-To: ${messageId}`,
            `References: ${messageId}`,
            ``,
            replyBody,
        ].join('\n');

        const encodedMessage = Buffer.from(rawMessage).toString('base64').replace(/\+/g, '-').replace(/\//g, '-').replace(/=+$/, '');

        await gmail.users.messages.send({
            userId: 'me',
            requestBody: {
                raw: encodedMessage,
            },
        });
    } catch (error) {
        console.error('Error sending reply email:', error);
        throw error;
    }
};


const sendReplyUsingOutlook = async (client: Client, message: any, replyMail: EmailReply): Promise<void> => {
    try {
        const fromEmailAddress = message.sender?.emailAddress?.address || '';

        const replyEmail = {
            message: {
                subject: replyMail.subject,
                body: {
                    contentType: 'Text',
                    content: replyMail.body,
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: fromEmailAddress,
                        },
                    },
                ],
            },
            saveToSentItems: 'false',
        };

        await client.api(`/me/sendMail`).post(replyEmail);
    } catch (error) {
        console.error('Error sending reply email:', error);
        throw error;
    }
};

async function getGmailLabelId(gmail: gmail_v1.Gmail, labelName: string) {
    let labelId = null;
    const labels = await gmail.users.labels.list({ userId: 'me' });
    const label = labels.data.labels.find(label => label.name === labelName);

    if (label) {
        labelId = label.id;
    } else {
        const createdLabel = await gmail.users.labels.create({
            userId: 'me',
            requestBody: {
                name: labelName,
                labelListVisibility: 'labelShow',
                messageListVisibility: 'show'
            }
        });
        labelId = createdLabel.data.id;
    }

    return labelId;
}

async function addOutlookCategory(graph: Client, message: any, category: string) {
    await graph.api(`/me/messages/${message.id}`).post({
        categories: [category]
    });
}

const fetchAndSendEmailGoogle = async (oauth2Client: any): Promise<any> => {
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    const messages = await (gmail.users.messages.list as any)({ userId: 'me', maxResults: 1, q: 'is:unread', orderBy: 'date desc' });

    const messageContents = messages.data?.messages ? await Promise.all(messages.data.messages.map(async (message: any) => {
        const messageData = await gmail.users.messages.get({ userId: 'me', id: message.id });
        return messageData.data;
    })) : [];
    const sentEmail: any[] = [];
    const messageReplied: string[] = [];
    for (const message of messageContents) {
        const analyzedResponse = await analyzeEmail(message.snippet || "");
        const response: AnalyzedResponse = JSON.parse(analyzedResponse);
        const replyMail = response.replyMail;
        const extractedMailContent = response.extractedMailContent;
        messageReplied.push(extractedMailContent);
        if (response.label == "Interested") {
            console.log("here")
            const labelId = await getGmailLabelId(gmail, 'Interested');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        else if (response.label == "Not Interested") {
            console.log("here")
            const labelId = await getGmailLabelId(gmail, 'Not Interested');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        else if (response.label == "More Information") {
            console.log("here")
            const labelId = await getGmailLabelId(gmail, 'More Information');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        else if (response.label == "Others") {
            console.log("here")
            const labelId = await getGmailLabelId(gmail, 'Others');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        if (replyMail?.body) {
            await sendReplyUsingGoogle(gmail, message, replyMail);
            sentEmail.push([`${response.label.toUpperCase()} - Mail is of ${response.label.toUpperCase()} category`, replyMail]);
        } else {
            sentEmail.push(["OTHER", "Mail is in other category no reply sent"]);
        }

        // Mark the email as read
        await gmail.users.messages.modify({
            userId: 'me',
            id: message.id,
            requestBody: {
                removeLabelIds: ['UNREAD']
            }
        });
    }
    return { message: messageReplied.length>0?"Success":"No New Mails", email: messageReplied.length > 0 ? messageReplied : ["No Mails"], replies: sentEmail.length > 0 ? sentEmail : ["No Mails"] };
};

const fetchAndSendEmailOutlook = async (graph: Client): Promise<any> => {
    const messages = await graph.api("/me/messages").filter("isRead eq false").orderby("receivedDateTime desc").top(1).get();
    const messageContents = messages.value ? await Promise.all(messages.value.map(async (message: any) => {
        const messageData = await graph.api(`/me/messages/${message.id}`).get();
        return messageData;
    })) : [];

    const sentEmail: any[] = [];
    const messageReplied: string[] = [];
    for (const message of messageContents) {
        const analyzedResponse = await analyzeEmail(message.body.content);
        const response: AnalyzedResponse = JSON.parse(analyzedResponse);
        const replyMail = response.replyMail;
        const extractedMailContent = response.extractedMailContent;
        messageReplied.push(extractedMailContent);
        if (response.label === "Interested" || response.label === "Not Interested" || response.label === "More Information" || response.label === "Others") {
            await addOutlookCategory(graph, message, response.label);
        }
        if (replyMail.body) {
            await sendReplyUsingOutlook(graph, message, replyMail);
            sentEmail.push([`${response.label.toUpperCase()} - Mail is of ${response.label.toUpperCase()} category`, replyMail]);
        } else {
            sentEmail.push([response.label.toUpperCase(), "Mail is in other category no reply sent"]);
        }

        await graph.api(`/me/messages/${message.id}`).update({
            isRead: true
        });
    }
    return { message: "Success", email: messageReplied.length > 0 ? messageReplied : ["No Mails"], replies: sentEmail.length > 0 ? sentEmail : ["No Mails"] };
};


export { fetchAndSendEmailGoogle, fetchAndSendEmailOutlook, sendReplyUsingGoogle, sendReplyUsingOutlook };
